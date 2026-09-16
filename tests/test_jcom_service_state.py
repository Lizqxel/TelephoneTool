from datetime import date
from pathlib import Path
import threading

from services.jcom_simulation_models import (
    AgeBracket,
    JcomSearchCriteria,
    ResidenceType,
    SimulationStatus,
)
from services.jcom_simulation_service import (
    JcomAddressNotFound,
    JcomCancelled,
    JcomSimulationService,
    address_zero_result_is_final,
)


class FakeDriver:
    current_url = "https://onlineshop.jcom.co.jp/Simulation/Simulation05"

    def __init__(self):
        self.quit_called = False

    def quit(self):
        self.quit_called = True


class FakeResultElement:
    text = (
        "J:COM NET 光(N)\n光(N) 1Gコース\n24カ月契約\n"
        "加入翌月の月額\n0円（税込）\n3カ月間\n基本料金 5,258円（税込）"
    )

    def screenshot(self, path):
        return False


def criteria():
    return JcomSearchCriteria(
        request_id="request-1",
        generation=1,
        product="J:COM",
        postal_code="0410811",
        address_original="北海道函館市富岡町３丁目２７番地１３号",
        address_normalized="北海道函館市富岡町3丁目27番地13号",
        birth_date=date(1990, 1, 1),
        age=36,
        age_bracket=AgeBracket.OVER_27,
        residence_type=ResidenceType.APARTMENT,
    )


class SuccessfulService(JcomSimulationService):
    def __init__(self, driver):
        super().__init__(
            threading.Event(), lambda message: None, lambda request: None,
            driver_factory=lambda: driver, screenshot_dir=Path.cwd(),
        )

    def _prune_screenshots(self):
        pass

    def _setup_area(self, criteria):
        pass

    def _select_address(self, criteria):
        return criteria.address_original, False

    def _open_simulation(self):
        pass

    def _apply_age_if_present(self, criteria):
        return True

    def _select_course_and_result(self):
        return FakeResultElement()


class CancelledService(SuccessfulService):
    def _setup_area(self, criteria):
        raise JcomCancelled()


class AddressNotFoundService(SuccessfulService):
    def _select_address(self, criteria):
        raise JcomAddressNotFound()


def test_success_result_and_owned_driver_release():
    driver = FakeDriver()
    result = SuccessfulService(driver).run(criteria())
    assert result.status is SimulationStatus.SUCCESS
    assert result.next_month_price_yen == 0
    assert result.applied_age_bracket is AgeBracket.OVER_27
    assert driver.quit_called


def test_cancelled_search_releases_only_owned_driver():
    driver = FakeDriver()
    result = CancelledService(driver).run(criteria())
    assert result.status is SimulationStatus.CANCELLED
    assert driver.quit_called


def test_address_not_found_reports_selected_residence_type():
    driver = FakeDriver()
    result = AddressNotFoundService(driver).run(criteria())
    assert result.status is SimulationStatus.ADDRESS_NOT_FOUND
    assert "集合住宅" in result.error_message
    assert driver.quit_called


def test_transient_zero_address_result_is_not_treated_as_final():
    loading = (
        "ご入力いただいた郵便番号の検索結果は0件でした。\n"
        "物件情報を取得しています..."
    )
    assert not address_zero_result_is_final(loading)
    assert address_zero_result_is_final(
        "ご入力いただいた郵便番号の検索結果は0件でした。"
    )
