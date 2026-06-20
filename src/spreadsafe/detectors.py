from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Any, ClassVar

import yaml


@dataclass
class Config:
    locale: str = "pl"
    redact_free_text: bool = True
    preserve_status_values: bool = True
    max_sample_rows_per_sheet: int = 500
    sensitive_columns: list[str] = field(default_factory=list)
    safe_enum_columns: list[str] = field(default_factory=list)
    deny_columns: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class Detection:
    label: str
    value: str
    start: int
    end: int
    confidence: float
    source: str


@dataclass(frozen=True)
class Decision:
    action: str
    label: str | None = None
    reasons: tuple[str, ...] = ()
    detections: tuple[Detection, ...] = ()


def load_config(path: Path | None) -> Config:
    if path is None or not path.exists():
        return Config()
    loaded = yaml.safe_load(path.read_text(encoding="utf-8"))
    if loaded is None:
        raw: Mapping[str, Any] = {}
    elif not isinstance(loaded, Mapping):
        raise ValueError("spreadsafe.yml must contain a mapping")
    else:
        raw = loaded
    return Config(
        locale=str(raw.get("locale", "pl")),
        redact_free_text=_bool_value(raw.get("redact_free_text"), "redact_free_text", True),
        preserve_status_values=_bool_value(
            raw.get("preserve_status_values"),
            "preserve_status_values",
            True,
        ),
        max_sample_rows_per_sheet=_positive_int(
            raw.get("max_sample_rows_per_sheet"),
            "max_sample_rows_per_sheet",
            500,
        ),
        sensitive_columns=_string_list(raw.get("sensitive_columns"), "sensitive_columns"),
        safe_enum_columns=_string_list(raw.get("safe_enum_columns"), "safe_enum_columns"),
        deny_columns=_string_list(raw.get("deny_columns"), "deny_columns"),
    )


class Detector:
    _presidio_analyzer: ClassVar[Any | None] = None
    _presidio_failed: ClassVar[bool] = False

    def __init__(self, config: Config) -> None:
        self.config = config

    def detect_text(self, value: str, *, include_presidio: bool = True) -> list[Detection]:
        detections: list[Detection] = []
        for label, pattern in PATTERNS.items():
            for match in pattern.finditer(value):
                detected = match.group(0)
                detections.append(
                    Detection(
                        label=label,
                        value=detected,
                        start=match.start(),
                        end=match.end(),
                        confidence=0.9,
                        source="regex",
                    )
                )
        if include_presidio:
            detections.extend(self._analyze_with_presidio(value))
        return _dedupe_detections(detections)

    def detect_path(self, value: str) -> list[Detection]:
        detections: list[Detection] = []
        for part in Path(value).parts:
            if part in {"", "."}:
                continue
            path_part = Path(part)
            detections.extend(self.detect_text(part, include_presidio=False))
            if path_part.suffix:
                detections.extend(self.detect_text(path_part.stem))
            else:
                detections.extend(self.detect_text(part))
        return _dedupe_detections(detections)

    def classify_cell(self, column_name: str | None, value: Any) -> Decision:
        if value is None or value == "":
            return Decision("preserve", reasons=("blank",))

        normalized_column = _normalize(column_name or "")
        if normalized_column in {_normalize(column) for column in self.config.deny_columns}:
            return Decision("redact", reasons=("config_deny_column",))
        if normalized_column in {_normalize(column) for column in self.config.sensitive_columns}:
            configured_label = self._label_from_column(normalized_column)
            if configured_label == "FREE_TEXT":
                return Decision("redact", reasons=("config_sensitive_column", "free_text"))
            return Decision(
                "tokenize",
                configured_label or "VALUE",
                reasons=("config_sensitive_column",),
            )
        column_label = self._label_from_column(normalized_column)
        if isinstance(value, str) and value.startswith("="):
            if column_label == "FREE_TEXT" and self.config.redact_free_text:
                return Decision("redact", reasons=("formula", "free_text"))
            if column_label is not None:
                return Decision("tokenize", column_label, reasons=("formula", "column_heuristic"))
            return Decision("preserve", reasons=("formula",))
        text = str(value)
        detections = tuple(self.detect_text(text))

        if self.config.redact_free_text and _is_free_text(normalized_column, text, detections):
            return Decision("redact", reasons=("free_text",), detections=detections)
        if detections:
            return Decision("tokenize", detections[0].label, reasons=("detected_sensitive",), detections=detections)
        if normalized_column in {_normalize(column) for column in self.config.safe_enum_columns}:
            return Decision("preserve", reasons=("config_safe_enum",))
        if self.config.preserve_status_values and _looks_like_status_column(normalized_column):
            return Decision("preserve", reasons=("status_enum",))
        if column_label is not None:
            if column_label == "FREE_TEXT" and self.config.redact_free_text:
                return Decision("redact", reasons=("free_text",), detections=detections)
            return Decision("tokenize", column_label, reasons=("column_heuristic",), detections=detections)
        if _looks_like_contact_column(normalized_column) and _looks_like_person_name(text):
            return Decision("tokenize", "PERSON", reasons=("likely_person_name",), detections=detections)
        if _looks_like_safe_enum(text):
            return Decision("preserve", reasons=("short_enum",))
        return Decision("preserve", reasons=("no_sensitive_signal",))

    def _label_from_column(self, normalized_column: str) -> str | None:
        if any(term in normalized_column for term in ("email", "e-mail", "mail")):
            return "EMAIL"
        if any(term in normalized_column for term in ("phone", "tel", "telefon")):
            return "PHONE"
        if "pesel" in normalized_column:
            return "PESEL"
        if "regon" in normalized_column:
            return "REGON"
        if any(term in normalized_column for term in ("nip", "vat", "tax")):
            return "NIP"
        if any(term in normalized_column for term in ("iban", "bank", "account")):
            return "IBAN"
        if any(term in normalized_column for term in ("invoice", "faktura", "fv")):
            return "INVOICE_ID"
        if any(term in normalized_column for term in ("company", "client", "customer", "firma", "kontrahent")):
            return "COMPANY"
        if any(term in normalized_column for term in ("name", "person", "imie", "nazwisko")):
            return "PERSON"
        if any(term in normalized_column for term in ("note", "comment", "opis", "uwagi")):
            return "FREE_TEXT"
        return None

    def _analyze_with_presidio(self, value: str) -> list[Detection]:
        if self._presidio_failed:
            return []
        if Detector._presidio_analyzer is None:
            try:
                from presidio_analyzer import AnalyzerEngine

                Detector._presidio_analyzer = AnalyzerEngine()
            except Exception:
                Detector._presidio_failed = True
                return []
        locales = [self.config.locale]
        if self.config.locale != "en":
            locales.append("en")
        results: list[Any] = []
        for locale in locales:
            try:
                results = Detector._presidio_analyzer.analyze(
                    text=value,
                    language=locale,
                    entities=["PERSON", "EMAIL_ADDRESS", "PHONE_NUMBER", "IBAN_CODE"],
                )
            except Exception:
                continue
            break
        else:
            return []
        detections: list[Detection] = []
        for result in results:
            label = _presidio_label(result.entity_type)
            if label is None:
                continue
            detected = value[result.start : result.end]
            detections.append(
                Detection(
                    label=label,
                    value=detected,
                    start=result.start,
                    end=result.end,
                    confidence=float(result.score),
                    source="presidio",
                )
            )
        return detections


def non_overlapping_detections(detections: list[Detection] | tuple[Detection, ...]) -> list[Detection]:
    ordered = sorted(
        detections,
        key=lambda item: (
            item.start,
            _detection_priority(item.label),
            -(item.end - item.start),
        ),
    )
    selected: list[Detection] = []
    for detection in ordered:
        if any(detection.start < existing.end and detection.end > existing.start for existing in selected):
            continue
        selected.append(detection)
    return selected


def _bool_value(raw: Any, field_name: str, default: bool) -> bool:
    if raw is None:
        return default
    if not isinstance(raw, bool):
        raise ValueError(f"{field_name} must be a boolean")
    return raw


def _positive_int(raw: Any, field_name: str, default: int) -> int:
    if raw is None:
        return default
    if isinstance(raw, bool):
        raise ValueError(f"{field_name} must be a positive integer")
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be a positive integer") from exc
    if value < 1:
        raise ValueError(f"{field_name} must be a positive integer")
    return value


def _string_list(raw: Any, field_name: str) -> list[str]:
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise ValueError(f"{field_name} must be a list of strings")
    if not all(isinstance(item, str) for item in raw):
        raise ValueError(f"{field_name} must be a list of strings")
    return raw


def _presidio_label(entity_type: str) -> str | None:
    return {
        "PERSON": "PERSON",
        "EMAIL_ADDRESS": "EMAIL",
        "PHONE_NUMBER": "PHONE",
        "IBAN_CODE": "IBAN",
    }.get(entity_type)


def _detection_priority(label: str) -> int:
    priorities = {
        "EMAIL": 0,
        "IBAN": 1,
        "PESEL": 2,
        "VAT_ID": 3,
        "REGON": 4,
        "NIP": 5,
        "INVOICE_ID": 6,
        "PHONE": 7,
        "COMPANY": 8,
    }
    return priorities.get(label, 99)


PATTERNS: dict[str, re.Pattern[str]] = {
    "EMAIL": re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE),
    "PHONE": re.compile(r"(?<!\d)(?:\+48[\s-]?)?(?:\d[\s-]?){9}(?!\d)"),
    "PESEL": re.compile(r"\b\d{11}\b"),
    "NIP": re.compile(r"\b(?:PL)?\d{10}\b", re.IGNORECASE),
    "REGON": re.compile(r"\b\d{9}(?:\d{5})?\b"),
    "IBAN": re.compile(r"\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b", re.IGNORECASE),
    "INVOICE_ID": re.compile(r"\b(?:FV|INV|FA|FS|VAT)[-/]?\d{2,4}(?:[-/]\d{1,8})+\b", re.IGNORECASE),
    "VAT_ID": re.compile(r"\b(?:VAT|NIP)[\s:-]*(?:PL)?\d{10}\b", re.IGNORECASE),
    "COMPANY": re.compile(r"\b[\w .&-]{2,}?\s+(?:sp\.?\s*z\s*o\.?o\.?|s\.a\.|llc|ltd|gmbh)\b", re.IGNORECASE),
}


def _dedupe_detections(detections: list[Detection]) -> list[Detection]:
    seen: set[tuple[str, str, int, int]] = set()
    unique: list[Detection] = []
    for detection in detections:
        key = (detection.label, detection.value, detection.start, detection.end)
        if key not in seen:
            unique.append(detection)
            seen.add(key)
    return unique


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def _looks_like_status_column(column_name: str) -> bool:
    return any(term in column_name for term in ("status", "state", "stage", "etap"))


def _looks_like_contact_column(column_name: str) -> bool:
    return any(
        term in column_name
        for term in ("contact", "kontakt", "owner", "opiekun", "person", "name", "imie", "nazwisko")
    )


def _looks_like_person_name(value: str) -> bool:
    words = value.strip().split()
    if len(words) not in {2, 3}:
        return False
    return all(_looks_like_name_word(word) for word in words)


def _looks_like_name_word(value: str) -> bool:
    parts = re.split(r"[-']", value)
    return all(
        part and (part[0].isupper() and part[1:].islower() or part.isupper())
        for part in parts
    )


def _looks_like_safe_enum(value: str) -> bool:
    return bool(re.fullmatch(r"[A-Z0-9 _.-]{1,24}", value)) and not any(char.isdigit() for char in value)


def _is_free_text(column_name: str, value: str, detections: tuple[Detection, ...]) -> bool:
    if any(term in column_name for term in ("note", "comment", "opis", "uwagi")) and detections:
        return True
    return len(value.split()) >= 5 and len(detections) >= 2
