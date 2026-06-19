# Plan 005: Prune Runtime Dependencies That Remain Unused

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat -- pyproject.toml uv.lock src tests`
> This repository had no commits when the plan was written, so there is no base SHA. Compare the "Current state" excerpts below against the live code before proceeding. If the relevant code no longer matches, treat it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: plans/003-wire-presidio-name-detection.md
- **Category**: migration
- **Planned at**: commit `NO_HEAD` (initial repository has no commits), 2026-06-19

## Why This Matters

The package declares several relatively heavy runtime dependencies that are not imported anywhere in source or tests. Unused runtime dependencies increase install time, lockfile size, and supply-chain surface. This plan should run after Plan 003, because Plan 003 may intentionally make `presidio-analyzer` reachable; after that, remove only dependencies that still have no runtime use.

## Current State

`pyproject.toml` declares these runtime dependencies:

```toml
# pyproject.toml:11-19
dependencies = [
  "faker>=26",
  "openpyxl>=3.1",
  "pandas>=2.2",
  "presidio-analyzer>=2.2",
  "presidio-anonymizer>=2.2",
  "pyyaml>=6",
  "typer>=0.12",
]
```

Audit evidence at plan time:

```bash
rg -n "import (faker|pandas|presidio)|from (faker|pandas|presidio)|Faker|AnalyzerEngine|AnonymizerEngine|DataFrame|pd\\." src tests pyproject.toml README.md
```

At plan time, the only matches were the dependency declarations in `pyproject.toml`. If Plan 003 has landed, this command may now show Presidio imports; that is expected. Remove only dependencies that still have no source/test import or documented runtime purpose.

## Commands You Will Need

| Purpose | Command | Expected on success |
|---|---|---|
| Usage search | `rg -n "faker|pandas|presidio|Faker|AnalyzerEngine|AnonymizerEngine|DataFrame|pd\\." src tests README.md pyproject.toml` | shows actual use or only manifest entries |
| Lock update | `UV_CACHE_DIR=/tmp/codex-uv-cache uv lock` | exits 0 and updates `uv.lock` |
| Tests | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q` | all tests pass |
| Lint | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .` | `All checks passed!` |
| Typecheck | `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests` | `Success: no issues found` |

## Scope

**In scope**:
- `pyproject.toml`
- `uv.lock`
- `README.md` only if dependency-related setup text becomes inaccurate
- `plans/README.md`

**Out of scope**:
- Do not remove dependencies used by Plan 003 or any other landed source code.
- Do not replace the package manager or remove `uv.lock`.
- Do not rewrite sanitizer logic to avoid dependencies as part of this plan.
- Do not remove dev dependencies unless they are clearly unused and the operator asks.

## Git Workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep the change limited to the in-scope files.

## Steps

### Step 1: Confirm Plan 003 Outcome

Open `src/spreadsafe/detectors.py` and check whether Presidio is now imported or lazily imported.

- If Plan 003 has not been executed, STOP. This plan depends on that decision.
- If Plan 003 was rejected and the maintainer decided not to use Presidio, include `presidio-analyzer` and `presidio-anonymizer` in the removal candidates.
- If Plan 003 landed, keep `presidio-analyzer` if it is imported. Remove `presidio-anonymizer` only if it remains unused.

**Verify**:

```bash
rg -n "presidio|AnalyzerEngine|AnonymizerEngine" src tests README.md pyproject.toml
```

Expected result depends on Plan 003:

- If Presidio is wired, at least one source match exists for `presidio` or `AnalyzerEngine`.
- If Presidio is not wired by explicit decision, only manifest matches may exist.

### Step 2: Identify Unused Runtime Dependencies

Run:

```bash
rg -n "faker|pandas|presidio|Faker|AnalyzerEngine|AnonymizerEngine|DataFrame|pd\\." src tests README.md pyproject.toml
```

Decide removal candidates:

- Remove `faker` if there are no source/test imports or README instructions for it.
- Remove `pandas` if there are no source/test imports or README instructions for it.
- Remove `presidio-anonymizer` if there are no source/test imports or README instructions for it.
- Remove `presidio-analyzer` only if Plan 003 was explicitly rejected or superseded.

Do not remove `openpyxl`, `pyyaml`, or `typer`; they are used directly.

**Verify**: Record the command output in your handoff summary. It should justify every removed dependency.

### Step 3: Update Manifest and Lockfile

Edit `pyproject.toml` to remove only the confirmed-unused dependencies. Keep the dependency list sorted in its current simple style.

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv lock
```

Expected result: command exits 0 and updates `uv.lock`.

**Verify**:

```bash
git diff -- pyproject.toml uv.lock
```

Expected result: only dependency removals and lockfile consequences. No source files should be changed.

### Step 4: Run Full Gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, mypy reports no issues.

## Test Plan

No new tests are required if this plan only removes unused dependencies. The verification is the full existing suite plus a usage search proving removed packages are not referenced.

## Done Criteria

- [ ] Every removed dependency has no source/test/README use after Plan 003's outcome is accounted for.
- [ ] `uv.lock` is regenerated with `uv lock`.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `git diff -- pyproject.toml uv.lock` shows only dependency cleanup.
- [ ] `plans/README.md` status row for Plan 005 is updated.

## STOP Conditions

Stop and report back if:

- Plan 003 has not been executed or explicitly rejected.
- A dependency appears unused but is required by a documented public install path or optional feature.
- `uv lock` changes unrelated package groups in a way you cannot explain.
- Any file outside the in-scope list must change to keep tests passing.

## Maintenance Notes

This plan intentionally follows Plan 003. If future work adds pandas-based CSV inference, Faker-backed pseudonyms, or Presidio anonymization, those dependencies can be reintroduced with tests that prove they are used.
