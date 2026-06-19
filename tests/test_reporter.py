from spreadsafe.detectors import Detector, load_config
from spreadsafe.reporter import _redact_sensitive_text


def test_report_redaction_handles_overlapping_detections() -> None:
    detector = Detector(load_config(None))

    redacted = _redact_sensitive_text(
        "REGON 012345678 and VAT: PL5252248481",
        detector,
    )

    assert "012345678" not in redacted
    assert "PL5252248481" not in redacted
    assert "_PHONE]" not in redacted
    assert "_NIP]" not in redacted
