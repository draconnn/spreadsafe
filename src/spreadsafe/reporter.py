from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import json
import re

from spreadsafe.detectors import Config, Detector, load_config, non_overlapping_detections
from spreadsafe.scanner import WorkbookReport


def write_reports(
    reports: list[WorkbookReport],
    risks: list[str],
    reports_dir: Path,
    config: Config | None = None,
) -> None:
    reports_dir.mkdir(parents=True, exist_ok=True)
    detector = Detector(config or load_config(None))
    (reports_dir / "workbook-report.json").write_text(
        _redact_sensitive_text(
            json.dumps([asdict(report) for report in reports], indent=2, ensure_ascii=False),
            detector,
        ),
        encoding="utf-8",
    )
    (reports_dir / "workbook-report.md").write_text(
        _redact_sensitive_text(_markdown_report(reports), detector),
        encoding="utf-8",
    )
    (reports_dir / "risk-report.md").write_text(
        _redact_sensitive_text(_risk_report(reports, risks), detector),
        encoding="utf-8",
    )


def _redact_sensitive_text(text: str, detector: Detector) -> str:
    redacted = text
    for detection in sorted(
        non_overlapping_detections(detector.detect_text(text)),
        key=lambda item: item.start,
        reverse=True,
    ):
        redacted = redacted[: detection.start] + f"[REDACTED_{detection.label}]" + redacted[detection.end :]
    return re.sub(r"\bSPREADSAFE_SHEET_\d{4}\b", "[REDACTED_SHEET]", redacted)


def _markdown_report(reports: list[WorkbookReport]) -> str:
    lines = ["# Workbook Structure Report", ""]
    for report in reports:
        lines.extend([f"## {Path(report.path).name}", "", f"- Type: `{report.kind}`"])
        for warning in report.warnings:
            lines.append(f"- Warning: {warning}")
        for sheet in report.sheets:
            lines.extend(
                [
                    "",
                    f"### Sheet: {sheet.name}",
                    f"- Hidden: {sheet.hidden}",
                    f"- Dimensions: {sheet.rows} rows x {sheet.columns} columns",
                    f"- Headers: {', '.join(sheet.headers)}",
                ]
            )
            if sheet.formulas:
                lines.append("- Formulas:")
                lines.extend(f"  - `{formula}`" for formula in sheet.formulas[:50])
            if sheet.data_validations:
                lines.append("- Data validations:")
                lines.extend(f"  - `{validation}`" for validation in sheet.data_validations)
            for column in sheet.columns_report:
                enum = f"; enum: {', '.join(column.enum_values[:10])}" if column.enum_values else ""
                lines.append(
                    f"- Column {column.index}: `{column.name}` ({column.inferred_type}, "
                    f"{column.sample_count} values, {column.null_count} blanks{enum})"
                )
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _risk_report(reports: list[WorkbookReport], risks: list[str]) -> str:
    lines = ["# Sanitization Risk Report", ""]
    if not risks and not any(report.warnings for report in reports):
        lines.append("No residual risks were detected by automated checks.")
        return "\n".join(lines) + "\n"
    for report in reports:
        for warning in report.warnings:
            lines.append(f"- {Path(report.path).name}: {warning}")
        for sheet in report.sheets:
            for warning in sheet.warnings:
                lines.append(f"- {Path(report.path).name}:{sheet.name}: {warning}")
            for comment in sheet.comments:
                lines.append(f"- {Path(report.path).name}:{sheet.name}:{comment}: comment present")
            for hyperlink in sheet.hyperlinks:
                lines.append(f"- {Path(report.path).name}:{sheet.name}:{hyperlink}: hyperlink present")
    lines.extend(f"- {risk}" for risk in risks)
    return "\n".join(lines) + "\n"
