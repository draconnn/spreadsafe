# Plan 008: Validate Output From the `sanitize` Command

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 3884113..HEAD -- src/spreadsafe/cli.py tests/test_package_integration.py`
> If any in-scope file changed since this plan was written, compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: plans/006-redact-validation-issue-values.md
- **Category**: correctness
- **Planned at**: commit `3884113`, 2026-06-20

## Why this matters

`spreadsafe package` runs validation before reporting success, but `spreadsafe sanitize` writes a package marker and exits successfully without validating its output. Since `sanitize` is a public command, users can mistake it for the same safety gate without reports. It should validate the generated `sanitized/` directory and fail with exit code 1 when validation finds issues.

## Current state

Relevant files:

- `src/spreadsafe/cli.py` defines `sanitize`, `validate`, `package`, and `package_directory`.
- `tests/test_package_integration.py` tests command exit codes through `CliRunner`.

Key excerpts:

```python
# src/spreadsafe/cli.py:49-66
@app.command()
def sanitize(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        _ensure_input_directory(input_dir)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config)
        sanitized_dir = out / "sanitized"
        _ensure_input_output_do_not_overlap(input_dir, out)
        _ensure_output_is_owned_or_empty(out)
        _clear_generated_directory(sanitized_dir)
        _clear_generated_directory(out / "reports")
        sanitized_dir.mkdir(parents=True, exist_ok=True)
        sanitizer.sanitize_directory(input_dir, sanitized_dir)
        (out / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    typer.echo(f"Wrote sanitized files to {out / 'sanitized'}")

# src/spreadsafe/cli.py:121-123
write_reports(reports, sanitizer.risks, reports_dir, config)
(output_dir / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
return validate_output(output_dir, config)
```

Repo conventions:

- Validation failures use exit code 1.
- Invalid input/config/preparation errors use exit code 2.
- CLI tests assert no traceback and inspect stderr.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Focused tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` | selected tests pass |
| Full tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:

- `src/spreadsafe/cli.py`
- `tests/test_package_integration.py`
- `plans/README.md`

**Out of scope**:

- Do not make `sanitize` write reports; that is `package` behavior.
- Do not change sanitizer internals.
- Do not change `validate_output` policy except as required by Plan 006 if it has not landed.
- Do not alter package command behavior.

## Git workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep changes limited to the in-scope files.

## Steps

### Step 1: Add a CLI regression test for sanitize validation failure

In `tests/test_package_integration.py`, add a test that proves `sanitize` calls `validate_output` and converts failed validation into exit code 1.

Use `monkeypatch` to replace `spreadsafe.cli.validate_output` with a small function returning `ValidationResult(False, issues=["redacted validation issue"])`. This avoids relying on a sanitizer bug to produce an invalid package.

The test should:

- Create an input directory with a simple safe CSV.
- Invoke `runner.invoke(app, ["sanitize", str(input_dir), "--out", str(output_dir)])`.
- Assert exit code 1.
- Assert stderr contains `redacted validation issue`.
- Assert stdout does not claim success.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` should fail before implementation because `sanitize` does not call validation.

### Step 2: Validate after marker write

In `src/spreadsafe/cli.py`, after `sanitize` writes `.spreadsafe-package`, call:

```python
result = validate_output(out, config)
```

After the `try` block, mirror package command handling:

- Print each `result.issues` entry to stderr prefixed with `error:`.
- Exit with code 1 if `not result.passed`.
- Print each warning to stderr before the success message.
- Print the existing success message only after validation passes.

Keep `ValueError` handling unchanged.

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

- New monkeypatched CLI test for failed sanitize validation.
- Existing sanitize success tests must still pass.
- Existing package and validate command tests must still pass.

## Done criteria

- [ ] `sanitize` calls `validate_output(out, config)` after writing the package marker.
- [ ] `sanitize` exits 1 and prints validation issues when validation fails.
- [ ] `sanitize` prints its success message only when validation passes.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row updated.

## STOP conditions

Stop and report back if:

- Plan 006 has not landed and validation issues still include raw sensitive values.
- The change appears to require generating reports from `sanitize`.
- The monkeypatch test cannot target `spreadsafe.cli.validate_output` cleanly.
- Any verification command fails twice after a reasonable fix attempt.

## Maintenance notes

After this lands, `sanitize`, `scan`, and `package` all have a validation gate. Plan 009 can then safely consolidate duplicate CLI flow without changing safety semantics.
