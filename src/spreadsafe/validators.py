from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import csv
import re
from typing import Any
import zipfile

from openpyxl import load_workbook

from spreadsafe.detectors import Config, Detector, load_config
from spreadsafe.sanitizer import (
    _data_validation_has_external_reference,
    _data_validation_text,
    _looks_like_csv_formula,
    _looks_like_amount_column,
    _looks_like_date_column,
    _looks_like_external_formula,
    _parse_decimal,
    _parse_iso_date,
)

ALLOWED_PACKAGE_ROOT_ENTRIES = {".gitignore", ".spreadsafe-package", "reports", "sanitized"}


@dataclass
class ValidationResult:
    passed: bool
    issues: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def validate_output(output_dir: Path, config: Config | None = None) -> ValidationResult:
    sanitized_dir = output_dir / "sanitized"
    issues: list[str] = []
    warnings: list[str] = []
    detector = Detector(config or load_config(None))
    if output_dir.is_symlink():
        issues.append(f"Package directory is a symlink: {output_dir}")
    if output_dir.exists() and output_dir.is_dir():
        _validate_package_root(output_dir, detector, issues)
    if sanitized_dir.is_symlink():
        issues.append(f"Sanitized directory is a symlink: {sanitized_dir}")
        return ValidationResult(False, issues, warnings)
    if not sanitized_dir.exists():
        issues.append(f"Missing sanitized directory: {sanitized_dir}")
        return ValidationResult(False, issues, warnings)
    if not sanitized_dir.is_dir():
        issues.append(f"Sanitized path is not a directory: {sanitized_dir}")
        return ValidationResult(False, issues, warnings)
    state_dir = output_dir / "state"
    if state_dir.exists():
        issues.append(f"Package contains local re-identification state: {state_dir}")

    for file_path in sorted(sanitized_dir.rglob("*")):
        if file_path.is_symlink():
            issues.append(f"Symlink is not allowed in sanitized package: {file_path}")
            continue
        if file_path.is_dir():
            continue
        _validate_path(file_path, sanitized_dir, detector, issues)
        if file_path.suffix.lower() == ".xlsx":
            _validate_xlsx(file_path, detector, issues, warnings)
        elif file_path.suffix.lower() == ".csv":
            _validate_csv_formulas(file_path, issues)
            _validate_csv_values(file_path, detector, issues)
        else:
            issues.append(f"Unexpected sanitized file type: {file_path}")
    reports_dir = output_dir / "reports"
    if reports_dir.exists():
        if reports_dir.is_symlink():
            issues.append(f"Reports directory is a symlink: {reports_dir}")
            return ValidationResult(False, issues, warnings)
        for file_path in sorted(reports_dir.rglob("*")):
            if file_path.is_symlink():
                issues.append(f"Symlink is not allowed in reports: {file_path}")
                continue
            if file_path.is_file() and file_path.suffix.lower() in {".md", ".json", ".csv", ".txt"}:
                _validate_text_file(file_path, detector, issues)
    return ValidationResult(not issues, issues, warnings)


def _validate_package_root(output_dir: Path, detector: Detector, issues: list[str]) -> None:
    for child in sorted(output_dir.iterdir()):
        if child.name not in ALLOWED_PACKAGE_ROOT_ENTRIES:
            issues.append(f"Unexpected package root entry: {child}")
            continue
        if child.is_symlink():
            issues.append(f"Symlink is not allowed in package root: {child}")
            continue
        if child.name == ".gitignore":
            if child.is_file():
                _validate_text_file(child, detector, issues)
            else:
                issues.append(f"Package .gitignore is not a file: {child}")
        elif child.name == ".spreadsafe-package":
            if not child.is_file():
                issues.append(f"Package marker is not a file: {child}")


def _validate_xlsx(
    file_path: Path,
    detector: Detector,
    issues: list[str],
    warnings: list[str],
) -> None:
    for part_name in _external_link_part_names(file_path):
        issues.append(f"{file_path.name}: external workbook link metadata remains: {part_name}")
    for part_name in _unsupported_ooxml_part_names(file_path):
        issues.append(f"{file_path.name}: unsupported workbook payload remains: {part_name}")
    workbook = load_workbook(file_path, data_only=False)
    for field_name, value in _document_property_values(workbook).items():
        for detection in detector.detect_text(value):
            if _is_safe_generated_value(detection.value):
                continue
            issues.append(
                f"{file_path.name}: document property {field_name} contains "
                f"{detection.label}: {detection.value}"
            )
    for property_name, value in _custom_document_property_values(workbook).items():
        for detection in detector.detect_text(value):
            if _is_safe_generated_value(detection.value):
                continue
            issues.append(
                f"{file_path.name}: custom document property {property_name} contains "
                f"{detection.label}: {detection.value}"
            )
    for name, defined_name in workbook.defined_names.items():
        defined_text = " ".join(
            str(value)
            for value in (name, defined_name.attr_text, defined_name.comment, defined_name.description)
            if value is not None
        )
        for detection in detector.detect_text(defined_text):
            if _is_safe_generated_value(detection.value):
                continue
            issues.append(
                f"{file_path.name}: defined name {name} contains "
                f"{detection.label}: {detection.value}"
            )
    for worksheet in workbook.worksheets:
        headers = [str(cell.value or f"Column {cell.column}") for cell in worksheet[1]]
        for column_index, header in enumerate(headers, start=1):
            if _csv_header_is_sensitive(detector, header):
                issues.append(
                    f"{file_path.name}:{worksheet.title}: row 1 column {column_index} "
                    "contains sensitive header"
                )
        for detection in detector.detect_text(worksheet.title):
            issues.append(
                f"{file_path.name}:{worksheet.title}: sheet title contains "
                f"{detection.label}: {detection.value}"
            )
        for location, value in _worksheet_header_footer_values(worksheet).items():
            for detection in detector.detect_text(value):
                if _is_safe_generated_value(detection.value):
                    continue
                issues.append(
                    f"{file_path.name}:{worksheet.title}: {location} contains "
                    f"{detection.label}: {detection.value}"
                )
        if worksheet.sheet_state != "visible":
            warnings.append(f"{file_path.name}:{worksheet.title}: hidden sheet requires manual review")
        for validation in worksheet.data_validations.dataValidation:
            if _data_validation_has_external_reference(validation):
                issues.append(
                    f"{file_path.name}:{worksheet.title}: "
                    "data validation contains external workbook reference"
                )
            validation_text = _data_validation_text(validation)
            for detection in detector.detect_text(validation_text):
                if _is_safe_generated_value(detection.value):
                    continue
                issues.append(
                    f"{file_path.name}:{worksheet.title}: data validation contains "
                    f"{detection.label}: {detection.value}"
                )
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.comment is not None:
                    issues.append(f"{file_path.name}:{worksheet.title}:{cell.coordinate}: comment remains")
                if cell.hyperlink is not None:
                    issues.append(f"{file_path.name}:{worksheet.title}:{cell.coordinate}: hyperlink remains")
                value = cell.value
                if not isinstance(value, str):
                    header = headers[cell.column - 1] if cell.column - 1 < len(headers) else ""
                    if _value_violates_config_sensitive(detector, header, value):
                        issues.append(
                            f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                            "configured sensitive value remains"
                        )
                        continue
                    if _is_safe_structured_xlsx_value(header, value):
                        continue
                    if value is not None:
                        for detection in detector.detect_text(str(value)):
                            if _is_safe_generated_value(detection.value):
                                continue
                            issues.append(
                                f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                                f"{detection.label} remains: {detection.value}"
                            )
                    continue
                if value.startswith("="):
                    header = headers[cell.column - 1] if cell.column - 1 < len(headers) else ""
                    if _formula_is_in_sensitive_column(detector, header):
                        issues.append(
                            f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                            "formula remains in sensitive column"
                        )
                    if _looks_like_external_formula(value):
                        issues.append(
                            f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                            "formula contains external workbook reference"
                        )
                    for detection in detector.detect_text(value):
                        if _is_safe_generated_value(detection.value):
                            continue
                        issues.append(
                            f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                            f"formula contains {detection.label}: {detection.value}"
                        )
                    continue
                header = headers[cell.column - 1] if cell.column - 1 < len(headers) else ""
                if _value_violates_config_sensitive(detector, header, value):
                    issues.append(
                        f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                        "configured sensitive value remains"
                    )
                    continue
                for detection in detector.detect_text(value):
                    if _is_safe_generated_value(detection.value):
                        continue
                    issues.append(
                        f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
                        f"{detection.label} remains: {detection.value}"
                    )


def _validate_text_file(file_path: Path, detector: Detector, issues: list[str]) -> None:
    text = file_path.read_text(encoding="utf-8", errors="ignore")
    for detection in detector.detect_text(text):
        if _is_safe_generated_value(detection.value):
            continue
        issues.append(f"{file_path.name}: {detection.label} remains: {detection.value}")


def _formula_is_in_sensitive_column(detector: Detector, header: str) -> bool:
    decision = detector.classify_cell(header, "placeholder")
    return decision.action in {"redact", "tokenize"} and bool(
        {"column_heuristic", "config_sensitive_column"} & set(decision.reasons)
    )


def _external_link_part_names(file_path: Path) -> list[str]:
    names: list[str] = []
    with zipfile.ZipFile(file_path, "r") as workbook_zip:
        for name in workbook_zip.namelist():
            if name.startswith("xl/externalLinks/"):
                names.append(name)
                continue
            if name == "xl/workbook.xml":
                content = workbook_zip.read(name)
                if b"<externalReferences" in content:
                    names.append(name)
            elif name == "xl/_rels/workbook.xml.rels":
                content = workbook_zip.read(name)
                if b"externalLink" in content or b"externalLinks/" in content:
                    names.append(name)
    return names


def _unsupported_ooxml_part_names(file_path: Path) -> list[str]:
    names: list[str] = []
    with zipfile.ZipFile(file_path, "r") as workbook_zip:
        for name in workbook_zip.namelist():
            if name.startswith(
                (
                    "xl/media/",
                    "xl/drawings/",
                    "xl/charts/",
                    "xl/embeddings/",
                    "xl/pivotCache/",
                    "xl/pivotTables/",
                    "xl/printerSettings/",
                    "xl/tables/",
                    "customXml/",
                )
            ):
                names.append(name)
    return names


def _validate_path(
    file_path: Path,
    sanitized_dir: Path,
    detector: Detector,
    issues: list[str],
) -> None:
    relative = file_path.relative_to(sanitized_dir).as_posix()
    for detection in detector.detect_path(relative):
        issues.append(f"{relative}: path contains {detection.label}: {detection.value}")


def _validate_csv_formulas(file_path: Path, issues: list[str]) -> None:
    with file_path.open(newline="", encoding="utf-8-sig") as handle:
        for row_index, row in enumerate(csv.reader(handle), start=1):
            for column_index, value in enumerate(row, start=1):
                if _looks_like_csv_formula(value):
                    issues.append(
                        f"{file_path.name}: row {row_index} column {column_index} contains CSV formula"
                    )


def _validate_csv_values(file_path: Path, detector: Detector, issues: list[str]) -> None:
    with file_path.open(newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))
    headers = rows[0] if rows else []
    for row_index, row in enumerate(rows, start=1):
        if row_index > 1 and len(row) > len(headers):
            issues.append(f"{file_path.name}: row {row_index} contains extra columns")
        for column_index, value in enumerate(row, start=1):
            header = headers[column_index - 1] if row_index > 1 and column_index <= len(headers) else ""
            if row_index == 1 and _csv_header_is_sensitive(detector, value):
                issues.append(
                    f"{file_path.name}: row 1 column {column_index} contains sensitive header"
                )
            if row_index > 1 and _value_violates_config_sensitive(detector, header, value):
                issues.append(
                    f"{file_path.name}: row {row_index} column {column_index} "
                    "configured sensitive value remains"
                )
                continue
            if _is_safe_structured_csv_value(header, value):
                continue
            detections = detector.detect_text(value)
            if detections:
                for detection in detections:
                    if _is_safe_generated_value(detection.value):
                        continue
                    issues.append(
                        f"{file_path.name}: row {row_index} column {column_index} "
                        f"{detection.label} remains: {detection.value}"
                    )
                continue


def _is_safe_structured_csv_value(header: str, value: str) -> bool:
    stripped = value.strip()
    if not stripped:
        return True
    if _looks_like_amount_column(header):
        parsed_amount = _parse_decimal(stripped)
        return (
            parsed_amount is not None
            and parsed_amount.is_finite()
            and any(separator in stripped for separator in (".", ","))
        )
    if _looks_like_date_column(header):
        return _parse_iso_date(stripped) is not None
    return False


def _is_safe_structured_xlsx_value(header: str, value: Any) -> bool:
    if isinstance(value, int) and not isinstance(value, bool):
        return False
    return isinstance(value, float) and _looks_like_amount_column(header) and not value.is_integer()


def _csv_header_is_sensitive(detector: Detector, value: str) -> bool:
    if _is_safe_generated_value(value) or value == "[REDACTED_TEXT]":
        return False
    decision = detector.classify_cell(value, value)
    return decision.action in {"redact", "tokenize"} and bool(
        {"column_heuristic", "config_sensitive_column"} & set(decision.reasons)
    )


def _value_violates_config_sensitive(detector: Detector, header: str, value: Any) -> bool:
    text = str(value)
    if not text.strip() or text in {"[REDACTED_TEXT]", "[REDACTED_FORMULA]"}:
        return False
    if _is_safe_generated_value(text):
        return False
    decision = detector.classify_cell(header, value)
    return decision.action in {"redact", "tokenize"} and "config_sensitive_column" in decision.reasons


def _is_safe_generated_value(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"(?:EMAIL|PHONE|IBAN|PESEL|NIP|REGON|VAT_ID) \d{4}|"
            r"VALUE \d{4}|"
            r"INV-FAKE-\d{4}|"
            r"(?:Company|Person|Sheet) \d{4}|"
            r"(?:file|directory)_\d{4}",
            value,
        )
    )


def _document_property_values(workbook: Any) -> dict[str, str]:
    properties = workbook.properties
    values: dict[str, str] = {}
    for field_name in (
        "creator",
        "lastModifiedBy",
        "title",
        "subject",
        "description",
        "keywords",
        "category",
        "contentStatus",
        "identifier",
        "language",
    ):
        value = getattr(properties, field_name, None)
        if isinstance(value, str) and value:
            values[field_name] = value
    return values


def _custom_document_property_values(workbook: Any) -> dict[str, str]:
    custom_properties = getattr(workbook, "custom_doc_props", None)
    values: dict[str, str] = {}
    for prop in getattr(custom_properties, "props", []):
        name = getattr(prop, "name", None)
        value = getattr(prop, "value", None)
        parts = [str(item) for item in (name, value) if item is not None]
        if name is not None and parts:
            values[str(name)] = " ".join(parts)
    return values


def _worksheet_header_footer_values(worksheet: Any) -> dict[str, str]:
    values: dict[str, str] = {}
    for container_name in (
        "oddHeader",
        "oddFooter",
        "evenHeader",
        "evenFooter",
        "firstHeader",
        "firstFooter",
    ):
        container = getattr(worksheet, container_name, None)
        if container is None:
            continue
        for section_name in ("left", "center", "right"):
            section = getattr(container, section_name, None)
            text = getattr(section, "text", None) if section is not None else None
            if isinstance(text, str) and text:
                values[f"{container_name}.{section_name}"] = text
    return values
