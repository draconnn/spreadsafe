# Plan 001: Remove and Validate XLSX Custom Document Properties

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat -- src/spreadsafe/sanitizer.py src/spreadsafe/validators.py tests/test_package_integration.py`
> This repository had no commits when the plan was written, so there is no base SHA. Compare the "Current state" excerpts below against the live code before proceeding. If the relevant code no longer matches, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `NO_HEAD` (initial repository has no commits), 2026-06-19

## Why this matters

The sanitizer clears core workbook properties and defined names, but it does not inspect OpenPyXL custom document properties. A workbook can contain a custom property such as a contact email, client name, or internal project code; today that metadata survives in `sanitized/*.xlsx`, and `spreadsafe validate` still passes. This is a privacy leak in the supported `.xlsx` format.

## Current State

- `src/spreadsafe/sanitizer.py` sanitizes workbooks and clears core document properties, but not `workbook.custom_doc_props`.
- `src/spreadsafe/validators.py` validates core document properties through `_document_property_values()`, but not custom document properties.
- `tests/test_package_integration.py` already has adjacent tests for document properties and defined names. Match that style.

Relevant excerpts:

```python
# src/spreadsafe/sanitizer.py:48-54
def sanitize_xlsx(self, source: Path, destination: Path) -> None:
    shutil.copy2(source, destination)
    workbook = load_workbook(destination, data_only=False)
    _clear_document_properties(workbook)
    if workbook.defined_names:
        self.risks.append(f"{source.name}: workbook defined names removed")
        workbook.defined_names.clear()
```

```python
# src/spreadsafe/validators.py:57-65
workbook = load_workbook(file_path, data_only=False)
for field_name, value in _document_property_values(workbook).items():
    for detection in detector.detect_text(value):
        if _is_safe_generated_value(detection.value):
            continue
        issues.append(
            f"{file_path.name}: document property {field_name} contains "
            f"{detection.label}: {detection.value}"
        )
```

Existing test pattern:

```python
# tests/test_package_integration.py:492-511
def test_package_clears_sensitive_xlsx_document_properties(tmp_path: Path) -> None:
    ...
    assert sanitized.properties.creator == "spreadsafe"
    assert sanitized.properties.lastModifiedBy == "spreadsafe"
    assert sanitized.properties.title is None
    assert validate_output(output_dir).passed
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
- `src/spreadsafe/validators.py`
- `tests/test_package_integration.py`
- `plans/README.md`

**Out of scope**:
- Do not change token formats in `src/spreadsafe/mapping.py`.
- Do not add encrypted pseudonym-map output.
- Do not change how normal workbook core properties are represented except as needed to keep existing behavior.

## Git Workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep the change limited to the in-scope files.

## Steps

### Step 1: Add a Failing Regression Test for Sanitization

In `tests/test_package_integration.py`, add a test near `test_package_clears_sensitive_xlsx_document_properties`. Use OpenPyXL's custom property support:

```python
from openpyxl.packaging.custom import StringProperty
```

Create a workbook with:

```python
workbook.custom_doc_props.append(StringProperty(name="Contact", value="jan@example.com"))
```

Run `package_directory(input_dir, output_dir)`, load the sanitized workbook, and assert:

- `result.passed` is true.
- `list(sanitized.custom_doc_props)` is empty, or at minimum no property name/value contains the original email.
- `validate_output(output_dir).passed` is true.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` should fail before implementation because the custom property survives.

### Step 2: Clear Custom Document Properties During Sanitization

In `src/spreadsafe/sanitizer.py`, add a helper, for example `_clear_custom_document_properties(workbook: Any) -> bool`, and call it inside `sanitize_xlsx()` immediately after `_clear_document_properties(workbook)`.

Expected behavior:

- If `workbook.custom_doc_props` exists and contains properties, clear all of them conservatively.
- Return `True` when properties were removed.
- Add a risk entry such as `"{source.name}: custom document properties removed"` when clearing happens.

OpenPyXL exposes `workbook.custom_doc_props.props` as the mutable list. Use that public object carefully:

```python
custom_props = getattr(workbook, "custom_doc_props", None)
if custom_props is not None and getattr(custom_props, "props", None):
    custom_props.props.clear()
```

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes the new sanitization test.

### Step 3: Validate Custom Document Properties

In `src/spreadsafe/validators.py`, add validation for `workbook.custom_doc_props` inside `_validate_xlsx()`, near the existing core document property loop.

Expected behavior:

- Iterate over every custom property.
- Inspect both the property name and its `value` when present.
- If `Detector.detect_text()` finds anything, append an issue such as: `"{file_path.name}: custom document property Contact contains EMAIL: <value>"`.
- Do not quote secrets other than dummy test values in tests.

Add a second test near `test_validate_rejects_sensitive_xlsx_document_properties` that creates a workbook directly under `sanitized/` with a custom property containing `jan@example.com`, calls `validate_output()`, and asserts validation fails with `custom document property` in the issue text.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 4: Run Full Gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, mypy reports no issues.

## Test Plan

- Add one package-level test proving custom document properties are removed from sanitized XLSX files.
- Add one validator test proving a custom property leak in `sanitized/` fails validation.
- Model both tests after the existing document-property tests in `tests/test_package_integration.py`.

## Done Criteria

- [ ] Sanitized `.xlsx` files no longer retain custom document properties.
- [ ] Validation fails if a sanitized workbook contains sensitive custom document property names or values.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] No files outside the in-scope list are modified.
- [ ] `plans/README.md` status row for Plan 001 is updated.

## STOP Conditions

Stop and report back if:

- OpenPyXL in this environment does not expose `workbook.custom_doc_props` or its `props` list.
- Clearing custom properties corrupts workbook save/load in OpenPyXL.
- The fix requires inspecting raw OOXML zip parts directly; that is larger than this plan.
- Any in-scope code differs materially from the excerpts above before you begin.

## Maintenance Notes

Future XLSX metadata work should use the same pattern: sanitize the metadata surface and add a validator check for the same surface. Reviewers should confirm the risk report does not echo custom property values.
