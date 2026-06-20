from pathlib import Path
from types import SimpleNamespace

import pytest

from spreadsafe.scanner import scan_xlsx


def test_scan_xlsx_handles_missing_workbook_security(monkeypatch: pytest.MonkeyPatch) -> None:
    workbook = SimpleNamespace(security=None, worksheets=[])
    monkeypatch.setattr("spreadsafe.scanner.load_workbook", lambda *_args, **_kwargs: workbook)

    report = scan_xlsx(Path("workbook.xlsx"))

    assert report.warnings == []
