# Plan 004: Reject Invalid Scalar Config Lists

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat -- src/spreadsafe/detectors.py tests/test_detectors.py tests/test_package_integration.py`
> This repository had no commits when the plan was written, so there is no base SHA. Compare the "Current state" excerpts below against the live code before proceeding. If the relevant code no longer matches, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: correctness
- **Planned at**: commit `NO_HEAD` (initial repository has no commits), 2026-06-19

## Why This Matters

The optional `spreadsafe.yml` config is meant to let users force-sensitive or deny-listed columns. Today a common YAML typo, `sensitive_columns: Email`, is parsed as the string `"Email"` and then converted with `list(...)`, producing `["E", "m", "a", "i", "l"]`. The CLI does not fail, but the user's intended override is silently ignored. A privacy tool should reject malformed policy config instead of pretending it applied it.

## Current State

- `src/spreadsafe/detectors.py` loads config with direct `list(raw.get(..., []))` conversions.
- CLI commands catch `ValueError` and exit with code 2, which is the right path for malformed config.

Relevant excerpt:

```python
# src/spreadsafe/detectors.py:43-52
raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
return Config(
    locale=str(raw.get("locale", "pl")),
    redact_free_text=bool(raw.get("redact_free_text", True)),
    preserve_status_values=bool(raw.get("preserve_status_values", True)),
    max_sample_rows_per_sheet=int(raw.get("max_sample_rows_per_sheet", 500)),
    sensitive_columns=list(raw.get("sensitive_columns", [])),
    safe_enum_columns=list(raw.get("safe_enum_columns", [])),
    deny_columns=list(raw.get("deny_columns", [])),
)
```

CLI error-handling pattern:

```python
# src/spreadsafe/cli.py:70-74
try:
    result = package_directory(input_dir, out)
except ValueError as exc:
    typer.echo(f"error: {exc}", err=True)
    raise typer.Exit(2) from exc
```

## Commands You Will Need

| Purpose | Command | Expected on success |
|---|---|---|
| Tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:
- `src/spreadsafe/detectors.py`
- `tests/test_detectors.py`
- `tests/test_package_integration.py`
- `plans/README.md`

**Out of scope**:
- Do not add a full schema-validation dependency.
- Do not redesign the config file format.
- Do not change defaults for valid absent config fields.

## Git Workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep the change limited to the in-scope files.

## Steps

### Step 1: Add Typed Config List Parsing

In `src/spreadsafe/detectors.py`, add a helper near `load_config()`:

```python
def _string_list(raw: object, field_name: str) -> list[str]:
    ...
```

Behavior:

- `None` -> `[]`.
- `list[str]` -> same strings.
- Any list containing non-strings -> raise `ValueError(f"{field_name} must be a list of strings")`.
- A scalar string -> raise `ValueError(f"{field_name} must be a list of strings")`.
- Any other type -> raise the same style of `ValueError`.

Then update `load_config()` to use this helper for:

- `sensitive_columns`
- `safe_enum_columns`
- `deny_columns`

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_detectors.py -q` should still pass existing tests after implementation.

### Step 2: Add Unit Tests for Config Parsing

In `tests/test_detectors.py`, add tests for:

- Valid list config still parses:
  ```yaml
  sensitive_columns:
    - Email
  ```
  Assert `load_config(...).sensitive_columns == ["Email"]`.
- Scalar config is rejected:
  ```yaml
  sensitive_columns: Email
  ```
  Assert `ValueError` and message contains `sensitive_columns must be a list of strings`.
- Non-string list entries are rejected:
  ```yaml
  deny_columns:
    - 123
  ```
  Assert `ValueError`.

Use `tmp_path` and write a temporary `spreadsafe.yml`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_detectors.py -q` passes.

### Step 3: Add CLI-Level Coverage for Friendly Failure

In `tests/test_package_integration.py`, add a Typer runner test near the other CLI invalid-path tests:

- Create input directory with `spreadsafe.yml` containing `sensitive_columns: Email`.
- Add a minimal `clients.csv`.
- Run `runner.invoke(app, ["package", str(input_dir), "--out", str(output_dir)])`.
- Assert `exit_code == 2`.
- Assert stderr contains `sensitive_columns must be a list of strings`.
- Assert stderr does not contain `Traceback`.

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

- Unit-level config parser tests for valid lists, scalar strings, and non-string list items.
- CLI-level test proving malformed config returns exit code 2 without traceback.
- Existing config-sensitive tests, especially scan config behavior, must remain green.

## Done Criteria

- [ ] Scalar strings no longer become character lists for config list fields.
- [ ] Malformed config exits through the existing friendly `ValueError` CLI path.
- [ ] Valid config examples from the README still work.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row for Plan 004 is updated.

## STOP Conditions

Stop and report back if:

- Existing tests rely on scalar config strings being accepted.
- Adding this validation requires changing Typer command signatures.
- You find additional config fields with the same problem and the fix expands beyond list fields.
- Any in-scope code differs materially from the excerpts above before you begin.

## Maintenance Notes

This should stay lightweight. If config grows beyond a few fields, a future plan can introduce schema validation, but this plan should only make the existing fields fail safely.
