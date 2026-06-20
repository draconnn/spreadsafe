# Plan 009: Deduplicate CLI Package-Like Command Flow

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 3884113..HEAD -- src/spreadsafe/cli.py tests/test_package_integration.py`
> If any in-scope file changed since this plan was written, compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: plans/008-validate-sanitize-command-output.md
- **Category**: tech-debt
- **Planned at**: commit `3884113`, 2026-06-20

## Why this matters

`scan` and `package_directory` duplicate the same safety-sensitive pipeline: prepare output, sanitize, scan sanitized files, write reports, write marker, and validate. Duplicated command flow makes future privacy fixes easier to apply to one path and miss in another. Consolidating the shared flow keeps behavior consistent while preserving each command's public message.

## Current state

Relevant files:

- `src/spreadsafe/cli.py` contains all CLI commands and helper functions.
- `tests/test_package_integration.py` covers `scan`, `sanitize`, and `package` command behavior.

Key excerpts:

```python
# src/spreadsafe/cli.py:18-46
@app.command()
def scan(input_dir: Path, out: Path = typer.Option(..., "--out")) -> None:
    try:
        _ensure_input_directory(input_dir)
        _ensure_input_output_do_not_overlap(input_dir, out)
        config = load_config(input_dir / "spreadsafe.yml")
        sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
        _ensure_output_is_owned_or_empty(out)
        ...
        write_reports(reports, sanitizer.risks, reports_dir, config)
        (out / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
        result = validate_output(out, config)
        if not result.passed:
            raise ValueError("; ".join(result.issues))
    except ValueError as exc:
        typer.echo(f"error: {exc}", err=True)
        raise typer.Exit(2) from exc
    typer.echo(f"Wrote reports to {reports_dir}")

# src/spreadsafe/cli.py:102-123
def package_directory(input_dir: Path, output_dir: Path) -> ValidationResult:
    _ensure_input_directory(input_dir)
    _ensure_input_output_do_not_overlap(input_dir, output_dir)
    config = load_config(input_dir / "spreadsafe.yml")
    sanitizer = Sanitizer(config, PseudonymMapper(seed=str(input_dir.resolve())))
    ...
    write_reports(reports, sanitizer.risks, reports_dir, config)
    (output_dir / ".spreadsafe-package").write_text("spreadsafe\n", encoding="utf-8")
    return validate_output(output_dir, config)
```

Repo conventions:

- Public CLI functions are thin wrappers around helpers where possible; `package_command` delegates to `package_directory`.
- Validation failures in `package_command` print each issue and exit 1.
- Preparation/config errors are `ValueError` and exit 2.

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
- `tests/test_package_integration.py` only if existing tests need small expectation updates
- `plans/README.md`

**Out of scope**:

- Do not change generated file layout.
- Do not change `package_directory` return type.
- Do not change `scan`, `sanitize`, or `package` command names/options.
- Do not refactor scanner, sanitizer, reporter, or validator modules.

## Git workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep changes limited to the in-scope files.

## Steps

### Step 1: Make validation result printing reusable

In `src/spreadsafe/cli.py`, extract the repeated command-side validation handling into a small helper, for example:

```python
def _exit_if_validation_failed(result: ValidationResult) -> None:
    ...
```

It should:

- Print warnings to stderr.
- Print issues to stderr with `error:` prefixes.
- Raise `typer.Exit(1)` when `not result.passed`.

Use it from `package_command` and, because Plan 008 is a dependency, from `sanitize` too.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 2: Make `scan` delegate to `package_directory`

Replace the duplicated body of `scan` with:

- A `try` block that calls `package_directory(input_dir, out)`.
- The same validation handling helper used by `package_command`.
- The existing success message shape: `Wrote reports to {out / "reports"}`.

Keep `ValueError` handling as exit code 2. Do not raise `ValueError` for validation failures; use exit code 1 consistently with `package_command`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_package_integration.py -q` passes.

### Step 3: Confirm no duplicate package flow remains

Run:

```bash
rg -n "scan_directory\\(|write_reports\\(|PseudonymMapper\\(seed=str\\(input_dir\\.resolve\\(\\)\\)\\)" src/spreadsafe/cli.py
```

Expected result:

- `scan_directory`, `write_reports`, and seeded `PseudonymMapper` are used inside `package_directory`.
- `scan` and `package_command` are thin wrappers.

If imports become unused, remove them from `cli.py`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` reports all checks passed.

### Step 4: Run full gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, and mypy reports no issues.

## Test plan

- Existing `scan` command tests must still pass.
- Existing `package` command tests must still pass.
- Existing `sanitize` command tests from Plan 008 must still pass.
- Add a new test only if existing coverage does not assert scan validation failure behavior.

## Done criteria

- [ ] `scan` delegates to `package_directory`.
- [ ] `scan`, `sanitize`, and `package` share validation result handling where applicable.
- [ ] Validation failures exit 1; input/config/preparation failures exit 2.
- [ ] No unused imports remain in `src/spreadsafe/cli.py`.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row updated.

## STOP conditions

Stop and report back if:

- Plan 008 has not landed.
- Preserving current scan success output requires duplicating the package pipeline.
- Existing tests reveal intentionally different validation semantics between `scan` and `package`.
- Any verification command fails twice after a reasonable fix attempt.

## Maintenance notes

After this plan, future changes to package preparation should happen in `package_directory` first. CLI wrappers should remain small and should not duplicate sanitizer/report/validator orchestration.
