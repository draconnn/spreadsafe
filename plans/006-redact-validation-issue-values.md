# Plan 006: Redact Sensitive Values From Validation Issues

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 3884113..HEAD -- src/spreadsafe/validators.py src/spreadsafe/cli.py tests/test_package_integration.py tests/test_reporter.py`
> If any in-scope file changed since this plan was written, compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: MED
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `3884113`, 2026-06-20

## Why this matters

`validate_output` currently includes detected sensitive values in validation issue strings. The CLI prints those strings on validation failure, so a command intended to prevent unsafe handoff can echo the data it found into terminal logs, automation logs, or chat transcripts. The validator should report the location and sensitive type without reproducing the value.

## Current state

- `src/spreadsafe/validators.py` builds issue strings for document properties, defined names, sheet titles, worksheet cells, CSV cells, and text reports.
- `src/spreadsafe/cli.py` prints each validation issue for the `validate` and `package` commands.
- `tests/test_package_integration.py` currently asserts raw dummy values appear in validation output.

Key excerpts:

```python
# src/spreadsafe/validators.py:117-120
issues.append(
    f"{file_path.name}: document property {field_name} contains "
    f"{detection.label}: {detection.value}"
)

# src/spreadsafe/validators.py:220-222
issues.append(
    f"{file_path.name}:{worksheet.title}:{cell.coordinate}: "
    f"formula contains {detection.label}: {detection.value}"
)

# src/spreadsafe/validators.py:342-344
issues.append(
    f"{file_path.name}: row {row_index} column {column_index} "
    f"{detection.label} remains: {detection.value}"
)
```

Repo conventions:

- Tests use `tmp_path`, `CliRunner`, and direct `validate_output` calls in `tests/test_package_integration.py`.
- Error handling in CLI commands catches `ValueError` as exit code 2; validation failures are exit code 1.
- Keep messages concise and deterministic; existing tests usually assert substrings rather than whole output.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Focused tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py tests/test_reporter.py -q` | selected tests pass |
| Full tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:

- `src/spreadsafe/validators.py`
- `tests/test_package_integration.py`
- `tests/test_reporter.py` only if a shared redaction helper needs direct test coverage
- `plans/README.md`

**Out of scope**:

- Do not weaken detection.
- Do not change token formats in `src/spreadsafe/mapping.py`.
- Do not redact safe generated placeholders such as `EMAIL 0001`; those are already explicitly allowed by `_is_safe_generated_value`.
- Do not change report redaction behavior except where validation issue formatting reuses a helper.

## Git workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep changes limited to the in-scope files.

## Steps

### Step 1: Add a validation issue formatter

In `src/spreadsafe/validators.py`, add a small helper near `_is_safe_generated_value`:

- It should accept a `Detection`.
- It should return only the label, for example `EMAIL`.
- If useful for readability, use wording such as `contains EMAIL` or `EMAIL remains`; do not append `detection.value`.

Replace every validation issue that currently appends `detection.value` with a redacted form. Search with:

```bash
rg -n "detection\\.value" src/spreadsafe/validators.py
```

Expected after the change: no validation issue string includes `detection.value`. If `detection.value` remains only inside `_is_safe_generated_value` checks, that is acceptable.

**Verify**: `rg -n "detection\\.value" src/spreadsafe/validators.py` shows only safe generated-value checks or no matches.

### Step 2: Update validation tests

Update existing tests that assert raw dummy values in validation failures. At minimum:

- `test_validation_fails_when_obvious_pii_remains`
- `test_validate_command_returns_failure_exit_code_for_leaks`
- any XLSX validator tests that assert a value after `contains <LABEL>` or `<LABEL> remains`

New expectations:

- The issue still names the file/cell/row location.
- The issue still names the sensitive label, such as `EMAIL`.
- The raw sensitive value is absent from `ValidationResult.issues`, `stdout`, and `stderr`.

Use the existing tests around `tests/test_package_integration.py:100-123` as the structural pattern.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 3: Run full gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, and mypy reports no issues.

## Test plan

- Update existing validation leak tests to assert labels remain and values are absent.
- Add one focused assertion for direct `validate_output` and one for `spreadsafe validate` CLI stderr.
- If an existing test already covers a path, prefer updating it over adding duplicate cases.

## Done criteria

- [ ] `validate_output` issue strings do not include raw detected values.
- [ ] `spreadsafe validate` failure output does not include raw detected values.
- [ ] Tests cover both returned issues and CLI stderr.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row updated.

## STOP conditions

Stop and report back if:

- The live validator no longer includes raw values in issue strings.
- Removing values makes a test unable to identify the failing location.
- A fix requires changing `Detector` classification behavior.
- Any verification command fails twice after a reasonable fix attempt.

## Maintenance notes

Reviewers should scrutinize error messages for accidental value interpolation. Future validator additions should follow the same rule: report location and label, never the detected value.
