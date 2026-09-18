from datetime import date, datetime

import pytest

from services.jcom_simulation_models import (
    AddressApproximation,
    AgeBracket,
    JcomSearchCriteria,
    JcomSimulationResult,
    JST,
    ResidenceType,
    SimulationStatus,
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


def test_jcom_search_defaults_missing_birth_date_to_over_27():
    criteria = JcomSearchCriteria.create(
        generation=1,
        postal_code="0410811",
        address="北海道函館市富岡町3丁目27番13号",
        birth_date_text="",
        residence_type=ResidenceType.APARTMENT,
    )

    assert criteria.birth_date is None
    assert criteria.age == 27
    assert criteria.age_bracket is AgeBracket.OVER_27


def test_result_explains_default_age_when_birth_date_was_missing():
    result = JcomSimulationResult(
        request_id="missing-birth-date",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="入力住所",
        confirmed_address="確定住所",
        residence_type=ResidenceType.APARTMENT,
        age=27,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=AgeBracket.OVER_27,
        age_selection_used=True,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        birth_date_provided=False,
        status=SimulationStatus.SUCCESS,
    )

    assert "生年月日未入力のため自動選択: 27歳以上" in result.display_text()
    assert "入力済み生年月日から判定" not in result.display_text()


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


def test_missing_collaboration_course_is_shown_as_undetermined_without_price():
    result = JcomSimulationResult(
        request_id="course-missing",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="入力住所",
        confirmed_address="サイト確定住所",
        residence_type=ResidenceType.DETACHED,
        age=30,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=None,
        age_selection_used=False,
        selected_service="ネットのみ",
        line_type="",
        course="光(N) 1Gコース",
        status=SimulationStatus.COURSE_UNAVAILABLE,
        error_message="判定不能：光(N) 1Gコースを確認できませんでした。",
    )
    text = result.display_text()
    assert text.startswith("判定: 判定不能")
    assert "月額（加入翌月）: 未取得" in text
    assert "サイト確定住所: サイト確定住所" in text


def test_nearest_address_price_is_clearly_not_for_input_address():
    result = JcomSimulationResult(
        request_id="nearest",
        generation=1,
        acquired_at=datetime.now(JST),
        input_address="東京都府中市美好町３丁目３０－３８",
        confirmed_address="東京都府中市美好町３丁目３０番地３６号",
        residence_type=ResidenceType.DETACHED,
        age=30,
        calculated_age_bracket=AgeBracket.OVER_27,
        applied_age_bracket=None,
        age_selection_used=False,
        selected_service="ネットのみ",
        line_type="J:COM NET 光(N)",
        course="光(N) 1Gコース",
        next_month_price_yen=0,
        partial_address=True,
        address_approximations=[AddressApproximation("38", "３６号")],
        status=SimulationStatus.PARTIAL,
    )
    text = result.display_text()
    assert "入力住所の料金ではありません" in text
    assert "指定 38 → 選択 ３６号" in text
    assert "月額（加入翌月）: 0円" in text
    assert "入力住所: 東京都府中市美好町３丁目３０－３８" in text
    assert "サイト確定住所: 東京都府中市美好町３丁目３０番地３６号" in text
