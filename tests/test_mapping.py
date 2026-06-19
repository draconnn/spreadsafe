from datetime import date

from spreadsafe.mapping import PseudonymMapper


def test_stable_pseudonyms_are_reused_across_labels() -> None:
    mapper = PseudonymMapper(seed="fixture")

    assert mapper.token("COMPANY", "ACME Sp. z o.o.") == "Company 0001"
    assert mapper.token("COMPANY", "ACME Sp. z o.o.") == "Company 0001"
    assert mapper.token("COMPANY", "Globex") == "Company 0002"
    assert mapper.token("EMAIL", "jan@example.com") == "EMAIL 0001"


def test_dates_shift_by_one_stable_offset() -> None:
    mapper = PseudonymMapper(seed="fixture")

    shifted_a = mapper.shift_date(date(2025, 1, 15))
    shifted_b = mapper.shift_date(date(2025, 1, 16))

    assert (shifted_b - shifted_a).days == 1
    assert shifted_a != date(2025, 1, 15)


def test_amounts_keep_sign_and_change_value() -> None:
    mapper = PseudonymMapper(seed="fixture")

    sanitized = mapper.perturb_amount(14832.72)
    negative = mapper.perturb_amount(-14832.72)

    assert sanitized > 0
    assert sanitized != 14832.72
    assert 10000 < sanitized < 20000
    assert negative < 0
    assert negative != -14832.72
    assert 10000 < abs(negative) < 20000


def test_amount_perturbation_never_uses_noop_multiplier() -> None:
    assert PseudonymMapper(seed="spreadsafe").perturb_amount(17) != 17.0
    assert PseudonymMapper(seed="fixture").perturb_amount(26) != 26.0
