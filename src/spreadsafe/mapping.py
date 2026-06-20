from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal
import hashlib
from typing import Any


@dataclass
class PseudonymMapper:
    seed: str = "spreadsafe"
    mappings: dict[str, dict[str, str]] = field(default_factory=lambda: defaultdict(dict))
    counters: dict[str, int] = field(default_factory=lambda: defaultdict(int))

    def token(self, label: str, value: Any) -> str:
        original = str(value)
        bucket = self.mappings.setdefault(label, {})
        if original in bucket:
            return bucket[original]
        self.counters[label] = self.counters.get(label, 0) + 1
        token = self._make_token(label, self.counters[label])
        bucket[original] = token
        return token

    def shift_date(self, value: date | datetime) -> date | datetime:
        offset = self._stable_int("date-offset", 220, 620)
        shifted = value - timedelta(days=offset)
        if isinstance(value, datetime) and not isinstance(shifted, datetime):
            return datetime.combine(shifted, value.time())
        return shifted

    def perturb_amount(self, value: int | float | Decimal) -> float:
        numeric = float(value)
        if numeric == 0:
            return 0.0
        multiplier_basis = self._stable_int(f"amount:{numeric}", 87, 113)
        if multiplier_basis == 100:
            multiplier_basis = 101
        perturbed = round(numeric * (multiplier_basis / 100), 2)
        if perturbed == 0:
            return 0.01 if numeric > 0 else -0.01
        if perturbed == numeric:
            perturbed = round(numeric + (0.01 if numeric > 0 else -0.01), 2)
        return perturbed

    def as_dict(self) -> dict[str, dict[str, str]]:
        return {label: dict(values) for label, values in self.mappings.items()}

    def _make_token(self, label: str, index: int) -> str:
        if label == "EMAIL":
            return f"EMAIL {index:04d}"
        if label == "PHONE":
            return f"PHONE {index:04d}"
        if label == "PERSON":
            return f"Person {index:04d}"
        if label == "COMPANY":
            return f"Company {index:04d}"
        if label == "IBAN":
            return f"IBAN {index:04d}"
        if label in {"PESEL", "NIP", "REGON", "VAT_ID"}:
            return f"{label} {index:04d}"
        if label == "INVOICE_ID":
            return f"INV-FAKE-{index:04d}"
        if label == "FILE":
            return f"file_{index:04d}"
        if label == "DIRECTORY":
            return f"directory_{index:04d}"
        if label == "SHEET":
            return f"Sheet {index:04d}"
        return f"{label} {index:04d}"

    def _stable_int(self, key: str, minimum: int, maximum: int) -> int:
        digest = hashlib.sha256(f"{self.seed}:{key}".encode("utf-8")).hexdigest()
        span = maximum - minimum + 1
        return minimum + (int(digest[:8], 16) % span)
