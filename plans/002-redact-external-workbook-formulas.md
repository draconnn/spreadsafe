# Plan 002: Redact External Workbook Formula References

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat -- src/spreadsafe/sanitizer.py src/spreadsafe/scanner.py src/spreadsafe/validators.py tests/test_package_integration.py`
> This repository had no commits when the plan was written, so there is no base SHA. Compare the "Current state" excerpts below against the live code before proceeding. If the relevant code no longer matches, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `NO_HEAD` (initial repository has no commits), 2026-06-19

## Why This Matters

The project intentionally treats formulas as structure, but formulas can also contain external workbook references. A formula like `='[prod-budget.xlsx]Sheet1'!A1` leaks the production workbook name into both the sanitized XLSX and the Markdown report, while the current risk report can still say no residual risks were detected. Because external references are explicitly called out as warn/reject candidates in the original tool requirements, v1 should conservatively redact these formulas and validate that none remain.

## Current State

- `src/spreadsafe/sanitizer.py` preserves non-sensitive formulas unchanged.
- `src/spreadsafe/scanner.py` includes formulas in reports.
- `src/spreadsafe/validators.py` only fails formulas when the regex detector finds known sensitive text.

Relevant excerpts:

```python
# src/spreadsafe/sanitizer.py:181-185
if isinstance(value, str) and value.startswith("="):
    if self.detector.detect_text(value):
        self.risks.append(f"{location}: formula redacted")
        return "[REDACTED_FORMULA]"
    return value
```

```python
# src/spreadsafe/scanner.py:84-86
if isinstance(cell.value, str) and cell.value.startswith("="):
    formula = "[SENSITIVE_FORMULA]" if detector.detect_text(cell.value) else cell.value
    formulas.append(f"{cell.coordinate}: {formula}")
```

```python
# src/spreadsafe/validators.py:117-124
if value.startswith("="):
    for detection in detector.detect_text(value):
        ...
        issues.append(
            f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
            f"formula contains {detection.label}: {detection.value}"
        )
```

Existing test pattern:

```python
# tests/test_package_integration.py:172-188
def test_package_redacts_formula_when_formula_contains_sensitive_literal(tmp_path: Path) -> None:
    ...
    assert sanitized["Orders"]["A2"].value == "[REDACTED_FORMULA]"
    report_text = (output_dir / "reports" / "workbook-report.md").read_text(encoding="utf-8")
    assert "jan@example.com" not in report_text
```

## Commands You Will Need

| Purpose | Command | Expected on success |
|---|---|---|
| Tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:
- `src/spreadsafe/sanitizer.py`
- `src/spreadsafe/scanner.py`
- `src/spreadsafe/validators.py`
- `tests/test_package_integration.py`
- `plans/README.md`

**Out of scope**:
- Do not build a full Excel formula parser.
- Do not change the default behavior for ordinary internal formulas like `=D2*1.23`.
- Do not add support for `.xlsm`, macros, pivot caches, or embedded objects.

## Git Workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep the change limited to the in-scope files.

## Steps

### Step 1: Add a Formula Classifier Helper

In `src/spreadsafe/sanitizer.py`, add a small helper near `_looks_like_csv_formula()`:

```python
def _looks_like_external_formula(value: str) -> bool:
    ...
```

Minimum behavior:

- Return `True` only for strings beginning with `=`.
- Return `True` for external workbook reference syntax containing a quoted or unquoted bracketed workbook segment, such as `='[prod-budget.xlsx]Sheet1'!A1` or `=[prod-budget.xlsx]Sheet1!A1`.
- Return `True` for formulas containing obvious URI-like external references such as `http://`, `https://`, or `file://`.
- Return `False` for normal formulas such as `=D2*1.23` and `=SUM(A1:A5)`.

Keep this helper in `sanitizer.py` so `validators.py` can import it similarly to `_looks_like_csv_formula`.

**Verify**: Add focused tests indirectly through the package tests in Step 2; do not add a new test file unless necessary.

### Step 2: Redact External Formulas During Sanitization

In `Sanitizer._sanitize_value()`, update the formula branch so external formulas are redacted before ordinary preservation:

```python
if isinstance(value, str) and value.startswith("="):
    if self.detector.detect_text(value) or _looks_like_external_formula(value):
        self.risks.append(f"{location}: formula redacted")
        return "[REDACTED_FORMULA]"
    return value
```

Add a regression test near `test_package_redacts_formula_when_formula_contains_sensitive_literal`:

- Create `A2` with `='[prod-budget.xlsx]Sheet1'!A1`.
- Run `package_directory()`.
- Assert sanitized `A2` is `[REDACTED_FORMULA]`.
- Assert `workbook-report.md` and `risk-report.md` do not contain `prod-budget.xlsx`.
- Assert `validate_output(output_dir).passed`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 3: Avoid Reporting External Formula Payloads

In `src/spreadsafe/scanner.py`, import `_looks_like_external_formula` from `spreadsafe.sanitizer` or duplicate only if importing would create a cycle. There is no current cycle: `sanitizer.py` does not import `scanner.py`.

Update formula reporting so external formulas are reported as `[EXTERNAL_FORMULA]` or `[SENSITIVE_FORMULA]`, not verbatim. Example:

```python
if detector.detect_text(cell.value):
    formula = "[SENSITIVE_FORMULA]"
elif _looks_like_external_formula(cell.value):
    formula = "[EXTERNAL_FORMULA]"
else:
    formula = cell.value
```

**Verify**: The test from Step 2 confirms report text does not contain the workbook name.

### Step 4: Fail Validation If External Formula References Remain

In `src/spreadsafe/validators.py`, import `_looks_like_external_formula` and update `_validate_xlsx()` formula validation:

- If a string formula looks external, append an issue such as `"{file_path.name}:{worksheet.title}:{cell.coordinate}: formula contains external workbook reference"`.
- Do not include the external workbook path/name in the issue text.
- Continue checking known detector hits as it does today.

Add a validator-only test near existing formula/validation tests:

- Create an output package manually with `sanitized/external.xlsx`.
- Put `='[prod-budget.xlsx]Sheet1'!A1` in a cell.
- Assert `validate_output(output_dir).passed` is false and one issue contains `external workbook reference`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 5: Run Full Gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, mypy reports no issues.

## Test Plan

- Add one package-level test proving external workbook formulas are redacted and not echoed in reports.
- Add one validator-level test proving a remaining external workbook formula fails validation.
- Preserve the existing test that ordinary formulas like `=D2*1.23` remain unchanged.

## Done Criteria

- [ ] External workbook formula references are redacted in sanitized XLSX files.
- [ ] External formula payloads are not echoed in reports.
- [ ] Validation fails if an external formula reference remains.
- [ ] Existing ordinary formula preservation behavior is unchanged.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row for Plan 002 is updated.

## STOP Conditions

Stop and report back if:

- Implementing this requires parsing arbitrary Excel formula grammar.
- OpenPyXL represents external links somewhere other than formulas and the needed change expands beyond formula detection.
- The change breaks the existing formula-preservation test.
- Any in-scope code differs materially from the excerpts above before you begin.

## Maintenance Notes

This is a conservative v1 guard, not a complete OOXML external-link sanitizer. Future work can add lower-level inspection for workbook external link parts, but reviewers should keep this plan focused on formula references and reports.
