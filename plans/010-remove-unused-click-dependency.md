# Plan 010: Remove the Unused Direct `click` Dependency

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 3884113..HEAD -- pyproject.toml uv.lock src tests`
> If any in-scope file changed since this plan was written, compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: migration
- **Planned at**: commit `3884113`, 2026-06-20

## Why this matters

`click` remains declared as a direct runtime dependency even though the source imports `typer`, not `click`. Typer will continue to resolve its own Click dependency transitively. Removing the unused direct dependency keeps the runtime manifest closer to actual imports and avoids carrying direct dependency responsibility for a package the project does not use directly.

## Current state

Relevant files:

- `pyproject.toml` declares runtime dependencies.
- `uv.lock` records the resolved dependency graph.
- Source imports `typer`, `openpyxl`, `presidio_analyzer`, and `yaml`; it does not import `click`.

Key excerpt:

```toml
# pyproject.toml:11-16
dependencies = [
  "click>=8.4.1",
  "openpyxl>=3.1",
  "presidio-analyzer>=2.2",
  "pyyaml>=6",
  "typer>=0.12",
]
```

Audit command output at plan time:

```bash
rg -n "import click|from click|import typer|from typer" pyproject.toml src tests -S
```

This showed `click` only in `pyproject.toml`; `typer` is imported in `src/spreadsafe/cli.py` and `tests/test_package_integration.py`.

Repo conventions:

- Runtime dependencies are declared in `pyproject.toml`.
- Lockfile is `uv.lock`.
- Previous dependency cleanup is tracked in `plans/005-prune-unused-runtime-dependencies.md`.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Usage search | `rg -n '"click|import click|from click|click\\.' src tests README.md pyproject.toml` | only the manifest dependency matches |
| Lock update | `UV_CACHE_DIR=/tmp/codex-uv-cache uv lock` | exits 0 and updates `uv.lock` if needed |
| Tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:

- `pyproject.toml`
- `uv.lock`
- `plans/README.md`

**Out of scope**:

- Do not remove `typer`.
- Do not remove transitive Click from `uv.lock` if Typer still requires it.
- Do not remove dev dependencies.
- Do not change CLI code.

## Git workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep changes limited to the in-scope files.

## Steps

### Step 1: Reconfirm Click is not directly used

Run:

```bash
rg -n '"click|import click|from click|click\\.' src tests README.md pyproject.toml
```

Expected result: the only match is the dependency declaration in `pyproject.toml`. If source, tests, or README use Click directly, stop and report.

### Step 2: Remove the direct dependency and refresh the lock

Edit `pyproject.toml` and remove only:

```toml
"click>=8.4.1",
```

Then run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv lock
```

Expected result: command exits 0. `uv.lock` may still contain Click as a transitive dependency of Typer; that is expected and acceptable.

**Verify**: `rg -n '"click>=8\\.4\\.1"|import click|from click|click\\.' pyproject.toml src tests README.md` returns no matches.

### Step 3: Run full gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, and mypy reports no issues.

## Test plan

No new tests are required for a manifest-only dependency cleanup. The full existing suite is the regression check.

## Done criteria

- [ ] `pyproject.toml` no longer declares `click` directly.
- [ ] `uv.lock` is refreshed with `uv lock`.
- [ ] Source, tests, and README do not directly import or reference Click APIs.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row updated.

## STOP conditions

Stop and report back if:

- Any direct Click import or Click API call exists in source, tests, or README.
- Removing direct Click changes Typer resolution in a way that breaks the CLI tests.
- `uv lock` needs network access or fails to resolve with the existing constraints.
- Any verification command fails twice after a reasonable fix attempt.

## Maintenance notes

If future code starts using Click directly, re-add it as a direct dependency in the same change that introduces the import. Until then, Typer should own the Click version constraint.
