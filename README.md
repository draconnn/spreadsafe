# spreadsafe

Local spreadsheet sanitizer for creating Codex-safe handoff packages from `.xlsx` and `.csv`
files.

## Install and run

```bash
uv run spreadsafe package ./client-files --out ./codex-safe
```

The package command runs scanning, sanitization, report generation, and validation.

## Outputs

- `sanitized/` contains sanitized `.xlsx` and `.csv` files.
- `reports/workbook-report.md` describes sheets, columns, formulas, enums, and workbook shape.
- `reports/workbook-report.json` contains the same structure in machine-readable form.
- `reports/risk-report.md` lists redactions and manual-review warnings.
Only `sanitized/` and `reports/` are generated for Codex handoff. Re-identification
state is not emitted by default.

## Optional config

Create `spreadsafe.yml` in the input directory:

```yaml
locale: pl
redact_free_text: true
preserve_status_values: true
max_sample_rows_per_sheet: 500
sensitive_columns: []
safe_enum_columns: []
deny_columns: []
```
