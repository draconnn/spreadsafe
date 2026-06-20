from pathlib import Path

import pytest

from spreadsafe.detectors import Config, Detection, Detector, load_config


def labels(results: list[Detection]) -> set[str]:
    return {result.label for result in results}


def test_detects_polish_and_generic_identifiers() -> None:
    detector = Detector(load_config(None))

    results = detector.detect_text(
        "Jan Kowalski, PESEL 44051401359, NIP 5252248481, REGON 012345678, "
        "IBAN PL61109010140000071219812874, email jan@example.com, tel +48 600 123 456, "
        "invoice FV/2025/331"
    )

    assert {
        "PESEL",
        "NIP",
        "REGON",
        "IBAN",
        "EMAIL",
        "PHONE",
        "INVOICE_ID",
    }.issubset(labels(results))


def test_presidio_analyzer_uses_configured_locale(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[str] = []

    class StubAnalyzer:
        def analyze(self, **kwargs: object) -> list[object]:
            calls.append(str(kwargs["language"]))
            return []

    monkeypatch.setattr(Detector, "_presidio_analyzer", StubAnalyzer())
    monkeypatch.setattr(Detector, "_presidio_failed", False)

    Detector(Config(locale="pl")).detect_text("Jan Kowalski")

    assert calls == ["pl"]


def test_presidio_locale_failure_does_not_disable_other_locales(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[str] = []

    class StubResult:
        entity_type = "PERSON"
        start = 0
        end = 10
        score = 0.85

    class StubAnalyzer:
        def analyze(self, **kwargs: object) -> list[object]:
            calls.append(str(kwargs["language"]))
            if kwargs["language"] == "pl":
                raise ValueError("unsupported language")
            return [StubResult()]

    monkeypatch.setattr(Detector, "_presidio_analyzer", StubAnalyzer())
    monkeypatch.setattr(Detector, "_presidio_failed", False)

    assert labels(Detector(Config(locale="pl")).detect_text("Jan Kowalski")) == {"PERSON"}
    results = Detector(Config(locale="en")).detect_text("John Smith")

    assert labels(results) == {"PERSON"}
    assert calls == ["pl", "en", "en"]


def test_transient_presidio_failure_does_not_disable_locale(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[str] = []

    class StubResult:
        entity_type = "PERSON"
        start = 0
        end = 10
        score = 0.85

    class StubAnalyzer:
        def analyze(self, **kwargs: object) -> list[object]:
            language = str(kwargs["language"])
            calls.append(language)
            if calls == ["pl"]:
                raise RuntimeError("temporary analyzer failure")
            if language == "pl":
                return [StubResult()]
            return []

    monkeypatch.setattr(Detector, "_presidio_analyzer", StubAnalyzer())
    monkeypatch.setattr(Detector, "_presidio_failed", False)

    assert Detector(Config(locale="pl")).detect_text("Jan Kowalski") == []
    assert labels(Detector(Config(locale="pl")).detect_text("Jan Kowalski")) == {"PERSON"}
    assert calls == ["pl", "en", "pl"]


def test_path_detection_uses_filename_stems_for_presidio() -> None:
    detector = Detector(load_config(None))

    assert detector.detect_path("config.xlsx") == []
    assert labels(detector.detect_path("John Smith.xlsx")) == {"PERSON"}
    assert labels(detector.detect_path("jan@example.com.xlsx")) == {"EMAIL"}


def test_column_overrides_and_safe_enums_drive_classification() -> None:
    config = load_config(None)
    config.sensitive_columns = ["Client Name"]
    config.safe_enum_columns = ["Status"]
    detector = Detector(config)

    assert detector.classify_cell("Client Name", "ACME Sp. z o.o.").action == "tokenize"
    assert detector.classify_cell("Status", "PAID").action == "preserve"
    assert detector.classify_cell("Status", "jan@example.com").action == "tokenize"


def test_free_text_with_multiple_sensitive_hits_is_redacted() -> None:
    detector = Detector(load_config(None))

    decision = detector.classify_cell(
        "Notes",
        "Client ACME called from +48 600 123 456 about invoice FV/2025/331",
    )

    assert decision.action == "redact"
    assert "free_text" in decision.reasons


def test_generic_free_text_with_detections_is_redacted_before_tokenizing() -> None:
    detector = Detector(load_config(None))

    decision = detector.classify_cell(
        "Details",
        "Client jan@example.com called from +48 600 123 456 about contract terms",
    )

    assert decision.action == "redact"
    assert "free_text" in decision.reasons


def test_sensitive_default_columns_win_over_uppercase_enum_shape() -> None:
    detector = Detector(load_config(None))

    assert detector.classify_cell("Company", "ACME SP. Z O.O.").action == "tokenize"
    assert detector.classify_cell("Name", "JAN KOWALSKI").action == "tokenize"


def test_contact_columns_with_likely_person_names_are_tokenized() -> None:
    detector = Detector(load_config(None))

    decision = detector.classify_cell("Contact", "Jan Kowalski")

    assert decision.action == "tokenize"
    assert decision.label == "PERSON"
    uppercase_decision = detector.classify_cell("Contact", "JAN KOWALSKI")
    assert uppercase_decision.action == "tokenize"
    assert uppercase_decision.label == "PERSON"
    assert detector.classify_cell("Status", "PAID").action == "preserve"


def test_load_config_accepts_string_lists(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text(
        "sensitive_columns:\n"
        "  - Email\n",
        encoding="utf-8",
    )

    config = load_config(config_path)

    assert config.sensitive_columns == ["Email"]


def test_configured_sensitive_columns_use_column_specific_labels() -> None:
    config = load_config(None)
    config.sensitive_columns = ["Email", "Phone", "NIP", "Internal ID"]
    detector = Detector(config)

    assert detector.classify_cell("Email", "internal-alias").label == "EMAIL"
    assert detector.classify_cell("Phone", "extension 42").label == "PHONE"
    assert detector.classify_cell("NIP", "vendor-tax-id").label == "NIP"
    assert detector.classify_cell("Internal ID", "customer-17").label == "VALUE"


def test_formula_cells_in_sensitive_columns_are_not_preserved() -> None:
    detector = Detector(load_config(None))

    decision = detector.classify_cell("Email", "=LOWER(A2)")

    assert decision.action == "tokenize"
    assert decision.label == "EMAIL"
    assert "formula" in decision.reasons


def test_load_config_rejects_non_mapping_root(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text("- Email\n", encoding="utf-8")

    with pytest.raises(ValueError, match="spreadsafe.yml must contain a mapping"):
        load_config(config_path)


def test_load_config_rejects_quoted_booleans(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text('redact_free_text: "false"\n', encoding="utf-8")

    with pytest.raises(ValueError, match="redact_free_text must be a boolean"):
        load_config(config_path)


def test_load_config_rejects_scalar_string_lists(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text("sensitive_columns: Email\n", encoding="utf-8")

    with pytest.raises(ValueError, match="sensitive_columns must be a list of strings"):
        load_config(config_path)


def test_load_config_rejects_non_positive_sample_limit(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text("max_sample_rows_per_sheet: 0\n", encoding="utf-8")

    with pytest.raises(ValueError, match="max_sample_rows_per_sheet must be a positive integer"):
        load_config(config_path)


def test_load_config_rejects_non_string_list_items(tmp_path: Path) -> None:
    config_path = tmp_path / "spreadsafe.yml"
    config_path.write_text(
        "deny_columns:\n"
        "  - 123\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="deny_columns must be a list of strings"):
        load_config(config_path)
