from datetime import date

import pytest

from services.jcom_simulation_models import (
    AgeBracket,
    ResidenceType,
    age_bracket_for,
    calculate_full_age,
    normalize_postal_code,
    parse_birth_date,
)
from ui.main_window_functions import MainWindowFunctions


class ComboStub:
    def __init__(self, value):
        self.value = value

    def currentText(self):
        return self.value


class BirthInputStub:
    _build_birth_date = MainWindowFunctions._build_birth_date

    def __init__(self, era, year, month, day):
        self.era_combo = ComboStub(era)
        self.year_combo = ComboStub(year)
        self.month_combo = ComboStub(month)
        self.day_combo = ComboStub(day)


def test_postal_code_full_width_and_hyphen():
    assert normalize_postal_code("０４１－０８１１") == "0410811"


@pytest.mark.parametrize("value", ["041081", "04108111", "abc-defg"])
def test_postal_code_invalid(value):
    with pytest.raises(ValueError):
        normalize_postal_code(value)


@pytest.mark.parametrize(
    ("birth", "today", "expected"),
    [
        (date(2004, 9, 17), date(2026, 9, 16), 21),
        (date(2004, 9, 16), date(2026, 9, 16), 22),
        (date(2003, 9, 17), date(2026, 9, 16), 22),
        (date(2003, 9, 16), date(2026, 9, 16), 23),
        (date(2000, 9, 17), date(2026, 9, 16), 25),
        (date(2000, 9, 16), date(2026, 9, 16), 26),
        (date(1999, 9, 17), date(2026, 9, 16), 26),
        (date(1999, 9, 16), date(2026, 9, 16), 27),
        (date(2000, 2, 29), date(2026, 2, 28), 25),
        (date(2000, 2, 29), date(2026, 3, 1), 26),
    ],
)
def test_full_age_boundaries_and_leap_day(birth, today, expected):
    assert calculate_full_age(birth, today) == expected


@pytest.mark.parametrize(
    ("age", "expected"),
    [
        (22, AgeBracket.UNDER_22),
        (23, AgeBracket.UNDER_26),
        (26, AgeBracket.UNDER_26),
        (27, AgeBracket.OVER_27),
    ],
)
def test_age_brackets(age, expected):
    assert age_bracket_for(age) is expected


@pytest.mark.parametrize("value", ["", "2024/2/30", "not-a-date"])
def test_birth_date_invalid_or_missing(value):
    with pytest.raises(ValueError):
        parse_birth_date(value)


def test_birth_date_future():
    with pytest.raises(ValueError):
        parse_birth_date("2999/1/1")


def test_existing_japanese_era_input_is_used():
    source = BirthInputStub("平成", "1", "1", "8")
    assert source._build_birth_date() == "1989/1/8"
    assert parse_birth_date(source._build_birth_date()) == date(1989, 1, 8)


def test_existing_birth_input_missing_or_invalid_is_rejected():
    missing = BirthInputStub("平成", "", "1", "1")
    with pytest.raises(ValueError):
        parse_birth_date(missing._build_birth_date())
    invalid = BirthInputStub("平成", "1", "2", "30")
    with pytest.raises(ValueError):
        parse_birth_date(invalid._build_birth_date())
