# Plan 007: Harden XLSX Formula Policy Beyond the Current Denylist

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 3884113..HEAD -- src/spreadsafe/sanitizer.py src/spreadsafe/validators.py tests/test_package_integration.py`
> If any in-scope file changed since this plan was written, compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `3884113`, 2026-06-20

## Why this matters

The project intentionally preserves local spreadsheet formulas as workbook structure, but it also treats formula execution surfaces as a safety concern. The current XLSX formula check is a narrow denylist for a few functions and external workbook references. A broader external-capable formula family can remain in the sanitized workbook and pass validation, which weakens the guarantee that sanitized XLSX files are safe to hand off.

## Current state

- `src/spreadsafe/sanitizer.py` redacts formulas only when `Detector` finds sensitive text, when the column is sensitive, or when `_looks_like_external_formula` returns true.
- `src/spreadsafe/validators.py` uses the same `_looks_like_external_formula` helper to reject residual formulas.
- Tests preserve local formulas and reject specific external forms.

Key excerpts:

```python
# src/spreadsafe/sanitizer.py:209-218
if isinstance(value, str) and value.startswith("="):
    decision = self.detector.classify_cell(column_name, value)
    if (
        decision.action != "preserve"
        or self.detector.detect_text(value)
        or _looks_like_external_formula(value)
    ):
        self.risks.append(f"{location}: formula redacted")
        return "[REDACTED_FORMULA]"
    return value

# src/spreadsafe/sanitizer.py:321-330
def _looks_like_external_formula(value: str) -> bool:
    if not value.startswith("="):
        return False
    if re.search(
        r"(?:^|[^A-Z0-9_])(?:_xlfn\.)?(?:HYPERLINK|WEBSERVICE|FILTERXML|IMAGE)\s*\(",
        value,
        re.IGNORECASE,
    ):
        return True
    return _looks_like_external_reference(value)
```

Existing tests to preserve:

```python
# tests/test_package_integration.py:481-497
def test_package_preserves_local_structured_reference_formulas(tmp_path: Path) -> None:
    ...
    assert sanitized.active["B2"].value == "=[@Amount]*1.23"
    assert sanitized.active["C2"].value == "=Table1[[#Totals],[Amount]]"
```

Repo conventions:

- Formula helper functions live near `_looks_like_external_formula` in `src/spreadsafe/sanitizer.py`.
- Validator imports sanitizer helpers instead of duplicating formula parsing.
- Tests build minimal workbooks with `openpyxl.Workbook` in `tests/test_package_integration.py`.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Focused tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` | selected tests pass |
| Full tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:

- `src/spreadsafe/sanitizer.py`
- `src/spreadsafe/validators.py` only if imports or validation messaging need adjustment
- `tests/test_package_integration.py`
- `plans/README.md`

**Out of scope**:

- Do not remove all formula preservation.
- Do not implement a full Excel parser.
- Do not change CSV formula handling unless a shared helper makes that unavoidable.
- Do not add network access or external files to tests.

## Git workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep changes limited to the in-scope files.

## Steps

### Step 1: Add regression tests for an external-capable XLSX formula family

In `tests/test_package_integration.py`, add two tests near the existing external formula tests:

- Package flow: create a workbook with an `RTD`-family formula, which is external-capable and not currently covered by `HYPERLINK`, `WEBSERVICE`, `FILTERXML`, `IMAGE`, or bracketed external references. Use harmless inert fixture strings such as `server` and `topic`. Assert the sanitized cell is `[REDACTED_FORMULA]` and `validate_output(output_dir).passed`.
- Validate flow: create a sanitized workbook containing the same formula and assert `validate_output` fails with a formula/external reference issue.

Do not include shell commands, local file paths, credentials, or live URLs in the fixture formula.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` should fail before implementation because the formula remains.

### Step 2: Centralize formula risk names

In `src/spreadsafe/sanitizer.py`, replace the inline function-name regex with a named constant such as `EXTERNAL_FORMULA_FUNCTIONS`.

Include the existing covered functions and the new external-capable family from Step 1. Keep the helper conservative and readable:

- It should still detect `_xlfn.`-prefixed functions.
- It should remain case-insensitive.
- It should not flag local arithmetic formulas or structured references covered by `test_package_preserves_local_structured_reference_formulas`.

**Verify**: focused tests still fail only if implementation is incomplete, not because local formulas are redacted.

### Step 3: Keep sanitizer and validator behavior aligned

Because `src/spreadsafe/validators.py` imports `_looks_like_external_formula`, updating the helper should make both package-time redaction and validation-time rejection use the same policy. Confirm no duplicate formula policy appears in `validators.py`.

Run:

```bash
rg -n "HYPERLINK|WEBSERVICE|FILTERXML|IMAGE|RTD|_looks_like_external_formula" src/spreadsafe
```

Expected result: formula risk names are centralized in `sanitizer.py`, and validators use `_looks_like_external_formula`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 4: Run full gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, and mypy reports no issues.

## Test plan

- Add one package regression test and one validator regression test for the newly covered XLSX formula family.
- Preserve existing tests for local structured formulas.
- Preserve existing tests for known external formula functions and external workbook references.

## Done criteria

- [ ] The new external-capable formula fixture is redacted during packaging.
- [ ] The same fixture is rejected by validation when already present in sanitized output.
- [ ] Local arithmetic and structured reference formulas remain preserved.
- [ ] Formula risk names are centralized in `src/spreadsafe/sanitizer.py`.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row updated.

## STOP conditions

Stop and report back if:

- Preserving local formulas and blocking the new external-capable fixture cannot both be achieved with a small helper change.
- The fix requires a full formula parser.
- Openpyxl rewrites the fixture formula in a way that makes the test invalid.
- Any verification command fails twice after a reasonable fix attempt.

## Maintenance notes

Formula policy should remain conservative and explicit. If future work needs more formula preservation, add allowlisted local formula cases with tests before relaxing external checks.
