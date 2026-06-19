# Plan 003: Wire Presidio Into Text Detection for Person Names

> **Executor instructions**: Follow this plan step by step. Run every verification command and confirm the expected result before moving to the next step. If anything in the "STOP conditions" section occurs, stop and report; do not improvise. When done, update the status row for this plan in `plans/README.md`, unless a reviewer told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat -- pyproject.toml src/spreadsafe/detectors.py tests/test_detectors.py tests/test_package_integration.py`
> This repository had no commits when the plan was written, so there is no base SHA. Compare the "Current state" excerpts below against the live code before proceeding. If the relevant code no longer matches, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `NO_HEAD` (initial repository has no commits), 2026-06-19

## Why This Matters

`pyproject.toml` declares `presidio-analyzer` and `presidio-anonymizer`, but the detector currently uses only local regex and column-name heuristics. That means obvious person names in generic columns, for example `Contact -> Jan Kowalski`, can remain unchanged because there is no person-name regex and no Presidio call. The project is a privacy sanitizer, so advertised PII recognition should either exist or be removed; this plan implements the useful path.

## Current State

- `pyproject.toml` declares Presidio dependencies.
- `src/spreadsafe/detectors.py` contains all detection logic and never imports Presidio.
- The detector only tokenizes person names when the column name itself looks like a person/name field.

Relevant excerpts:

```toml
# pyproject.toml:15-16
"presidio-analyzer>=2.2",
"presidio-anonymizer>=2.2",
```

```python
# src/spreadsafe/detectors.py:59-74
def detect_text(self, value: str) -> list[Detection]:
    detections: list[Detection] = []
    for label, pattern in PATTERNS.items():
        for match in pattern.finditer(value):
            detected = match.group(0)
            detections.append(
                Detection(
                    label=label,
                    value=detected,
                    start=match.start(),
                    end=match.end(),
                    confidence=0.9,
                    source="regex",
                )
            )
    return _dedupe_detections(detections)
```

```python
# src/spreadsafe/detectors.py:125-128
if any(term in normalized_column for term in ("company", "client", "customer", "firma", "kontrahent")):
    return "COMPANY"
if any(term in normalized_column for term in ("name", "person", "imie", "nazwisko")):
    return "PERSON"
```

Existing test style:

```python
# tests/test_detectors.py:8-25
def test_detects_polish_and_generic_identifiers() -> None:
    detector = Detector(load_config(None))
    results = detector.detect_text(...)
    assert {...}.issubset(labels(results))
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
- `pyproject.toml` and `uv.lock` only if dependency changes are required
- `plans/README.md`

**Out of scope**:
- Do not call any external API or download language models at runtime.
- Do not add encrypted pseudonym state output.
- Do not overhaul token naming in `src/spreadsafe/mapping.py`.
- Do not implement a full natural-language PII classifier beyond Presidio integration and minimal safe fallbacks.

## Git Workflow

- Do not create or switch branches unless the operator explicitly instructs you.
- Do not commit or push unless the operator explicitly instructs you.
- Keep the change limited to the in-scope files.

## Steps

### Step 1: Add a Lazy Presidio Analyzer Wrapper

In `src/spreadsafe/detectors.py`, add an optional lazy analyzer initialization. It must not fail package import if Presidio or its NLP engine cannot initialize.

Target shape:

```python
class Detector:
    _presidio_analyzer: Any | None = None
    _presidio_failed: bool = False

    def _analyze_with_presidio(self, value: str) -> list[Detection]:
        ...
```

Requirements:

- Import `AnalyzerEngine` inside the helper, not at module import time.
- Analyze at least these entities: `PERSON`, `EMAIL_ADDRESS`, `PHONE_NUMBER`, `IBAN_CODE`, `US_BANK_NUMBER` if available. Map Presidio entity names to local labels, especially `PERSON -> PERSON`, `EMAIL_ADDRESS -> EMAIL`, `PHONE_NUMBER -> PHONE`, and `IBAN_CODE -> IBAN`.
- Catch exceptions during initialization and analysis. On exception, set a failure flag and return `[]`; this keeps the CLI local and robust in environments without Presidio NLP resources.
- Do not log or print input values.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_detectors.py -q` still passes.

### Step 2: Merge Presidio Results With Regex Results

Update `Detector.detect_text()` so it:

1. Collects existing regex detections.
2. Extends with Presidio detections.
3. Returns `_dedupe_detections(detections)`.

Presidio offsets are character offsets; use the same `Detection(start, end, value, confidence, source)` shape. Confidence can come from `result.score`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_detectors.py -q` passes.

### Step 3: Add a Deterministic Fallback for Simple Polish-Like Person Names

Because Presidio may be unavailable in a no-model local environment, add a small conservative fallback only for common two-word names when the column name suggests a contact/person, but not a company/status/amount/date field.

Do this inside classification, not broad global regex, to avoid turning ordinary free text into names too aggressively:

- If `normalized_column` contains one of `contact`, `kontakt`, `owner`, `opiekun`, `person`, `name`, `imie`, `nazwisko`.
- And the value is two or three title-cased words using letters only.
- Then return `Decision("tokenize", "PERSON", reasons=("likely_person_name",))`.

This fallback should make `Contact -> Jan Kowalski` tokenize even if Presidio is disabled. Keep it conservative.

**Verify**: Add `tests/test_detectors.py` coverage for:

- `detector.classify_cell("Contact", "Jan Kowalski").action == "tokenize"`.
- `detector.classify_cell("Status", "PAID").action == "preserve"` remains true.

### Step 4: Add Package-Level Coverage for Generic Contact Names

In `tests/test_package_integration.py`, add a CSV package test:

- Input: `contacts.csv` with headers `Contact,Status` and row `Jan Kowalski,ACTIVE`.
- Run `package_directory()`.
- Assert `Jan Kowalski` is not present in `sanitized/contacts.csv`.
- Assert the sanitized CSV contains `Person 0001`.
- Assert `validate_output(output_dir).passed`.

**Verify**: `UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest tests/test_detectors.py tests/test_package_integration.py -q` passes.

### Step 5: Run Full Gates

Run:

```bash
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m pytest -q
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m ruff check .
UV_CACHE_DIR=/tmp/codex-uv-cache uv run python -m mypy src tests
```

Expected result: all tests pass, ruff reports all checks passed, mypy reports no issues.

## Test Plan

- Unit tests in `tests/test_detectors.py` for generic contact/person classification.
- Integration test in `tests/test_package_integration.py` proving `Contact -> Jan Kowalski` does not survive in sanitized CSV.
- Existing tests for company, email, status enum, and free-text behavior must remain green.

## Done Criteria

- [ ] `Detector.detect_text()` includes Presidio detections when available.
- [ ] The CLI still works without network access or model downloads.
- [ ] Generic contact names are tokenized in supported CSV/XLSX flows.
- [ ] Existing safe status enum behavior is unchanged.
- [ ] `pytest`, `ruff`, and `mypy` gates pass.
- [ ] `plans/README.md` status row for Plan 003 is updated.

## STOP Conditions

Stop and report back if:

- Presidio initialization requires downloading external models during tests or normal CLI execution.
- The fallback starts tokenizing common safe enum/status values.
- The implementation requires changing public CLI behavior or config schema.
- Any in-scope code differs materially from the excerpts above before you begin.

## Maintenance Notes

Reviewers should scrutinize false positives. This tool is allowed to be conservative, but broad name detection can reduce report usefulness if it tokenizes safe business labels. If Presidio remains too heavy or unreliable locally, Plan 005 should remove the Presidio dependencies instead of leaving them declared but unused.
