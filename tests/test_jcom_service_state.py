from datetime import date
from pathlib import Path
import threading
import pytest

from services.jcom_simulation_models import (
    AddressCandidate,
    AgeBracket,
    JcomSearchCriteria,
    ResidenceType,
    SimulationStatus,
)
from services.jcom_simulation_service import (
    JcomAddressNotFound,
    JcomCancelled,
    JcomCourseUnavailable,
    JcomSiteError,
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


def test_headless_driver_uses_background_chrome_options(monkeypatch, tmp_path):
    captured = {}

    class CreatedDriver:
        def set_page_load_timeout(self, timeout):
            captured["page_load_timeout"] = timeout

        def implicitly_wait(self, timeout):
            captured["implicit_wait"] = timeout

    def create_chrome(*, service, options):
        captured["arguments"] = options.arguments
        return CreatedDriver()

    monkeypatch.setattr(
        "services.jcom_simulation_service.webdriver.Chrome", create_chrome
    )
    service = JcomSimulationService(
        threading.Event(),
        lambda message: None,
        screenshot_dir=tmp_path,
        headless=True,
    )

    service._create_driver()

    assert "--headless=new" in captured["arguments"]
    assert "--disable-gpu" in captured["arguments"]
    assert "--no-sandbox" in captured["arguments"]
    assert "--disable-dev-shm-usage" in captured["arguments"]
    assert "--disable-software-rasterizer" in captured["arguments"]
    service._profile_dir.cleanup()


def test_price_is_notified_before_screenshot_and_browser_cleanup():
    events = []

    class TimedDriver(FakeDriver):
        def quit(self):
            events.append("quit")
            super().quit()

    class TimedService(SuccessfulService):
        def _capture_result_screenshot(self, element, path, confirmed_address=""):
            events.append("capture")
            return False

    service = TimedService(TimedDriver())
    service.result_ready_callback = lambda result: events.append("result")
    service.screenshot_ready_callback = lambda result: events.append("screenshot")

    service.run(criteria())

    assert events == ["result", "capture", "screenshot", "quit"]


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


def test_site_error_page_is_detected_without_waiting_for_timeout():
    service = JcomSimulationService(threading.Event(), lambda message: None)
    service.driver = FakeDriver()
    service.driver.current_url = (
        "https://onlineshop.jcom.co.jp/Err?errorId=ERGU001&pathName=/"
    )
    with pytest.raises(JcomSiteError, match="ERGU001"):
        service._wait(lambda driver: False, timeout=1)


def test_empty_course_list_is_reported_as_unavailable(monkeypatch):
    class EmptyCourseDriver:
        current_url = "https://onlineshop.jcom.co.jp/Simulation/Simulation05"

        def find_elements(self, by, selector):
            return []

        def find_element(self, by, selector):
            return type("Body", (), {"text": "コースをお選びください"})()

    class ImmediateEvent:
        def is_set(self):
            return False

        def wait(self, timeout):
            return False

    ticks = iter((0.0, 1.0, 2.0, 10.0, 11.0))
    monkeypatch.setattr(
        "services.jcom_simulation_service.time.monotonic",
        lambda: next(ticks),
    )
    service = JcomSimulationService(ImmediateEvent(), lambda message: None)
    service.driver = EmptyCourseDriver()
    with pytest.raises(JcomCourseUnavailable, match="判定不能"):
        service._select_course_and_result()


def test_au_hikari_1g_is_never_selected_as_a_substitute(monkeypatch):
    class AuCourse:
        def get_attribute(self, name):
            return {"value": "1G", "id": "au-course"}.get(name)

        def is_enabled(self):
            return True

        def find_element(self, by, selector):
            return type("Label", (), {"text": "光 1Gコース on auひかり"})()

    class AuOnlyDriver:
        current_url = "https://onlineshop.jcom.co.jp/Simulation/Simulation05"
        clicked = False

        def find_elements(self, by, selector):
            if selector == "input[name='course-net']":
                return [AuCourse()]
            return []

        def find_element(self, by, selector):
            return type("Body", (), {"text": "コースをお選びください"})()

        def execute_script(self, *args):
            self.clicked = True
            raise AssertionError("auひかりをクリックしてはいけません")

    class ImmediateEvent:
        def is_set(self):
            return False

        def wait(self, timeout):
            return False

    ticks = iter((0.0, 1.0, 2.0, 10.0, 11.0))
    monkeypatch.setattr(
        "services.jcom_simulation_service.time.monotonic",
        lambda: next(ticks),
    )
    driver = AuOnlyDriver()
    service = JcomSimulationService(ImmediateEvent(), lambda message: None)
    service.driver = driver
    with pytest.raises(JcomCourseUnavailable, match="判定不能"):
        service._select_course_and_result()
    assert not driver.clicked


def test_identical_town_dom_duplicates_can_be_collapsed_but_distinct_cannot():
    class Element:
        def __init__(self, html):
            self.html = html

        def get_attribute(self, name):
            return self.html if name == "outerHTML" else None

    candidates = [
        AddressCandidate("0:0", "千葉市若葉区若松町"),
        AddressCandidate("0:1", "千葉市若葉区若松町"),
        AddressCandidate("0:2", "該当する住所がない方はこちら", "missing_address"),
    ]
    same = {
        "0:0": Element('<a name="千葉市若葉区">千葉市若葉区若松町</a>'),
        "0:1": Element('<a name="千葉市若葉区">千葉市若葉区若松町</a>'),
    }
    assert JcomSimulationService._identical_duplicate_candidate_id(
        candidates, same
    ) == "0:0"
    different = dict(same)
    different["0:1"] = Element('<a name="異なる経路">千葉市若葉区若松町</a>')
    assert JcomSimulationService._identical_duplicate_candidate_id(
        candidates, different
    ) is None


def test_only_apartment_building_candidates_require_manual_selection():
    from dataclasses import replace

    apartment = criteria()
    detached = replace(apartment, residence_type=ResidenceType.DETACHED)
    buildings = [
        AddressCandidate("2:0", "サニーハイツ"),
        AddressCandidate("2:1", "メゾン・ド・クラ２"),
    ]
    assert JcomSimulationService._is_building_choice(
        apartment, "府中市美好町３丁目|３０番地", buildings
    )
    assert JcomSimulationService._is_building_choice(
        apartment, "富岡町３丁目|２７番地|１３号", buildings
    )
    assert not JcomSimulationService._is_building_choice(
        detached, "府中市美好町３丁目|３０番地", buildings
    )
    assert not JcomSimulationService._is_building_choice(
        apartment,
        "府中市美好町３丁目|３０番地",
        [AddressCandidate("2:0", "３８号")],
    )


def test_only_apartment_room_candidates_after_building_require_manual_selection():
    from dataclasses import replace

    apartment = criteria()
    detached = replace(apartment, residence_type=ResidenceType.DETACHED)
    rooms = [
        AddressCandidate("3:0", "２０５号"),
        AddressCandidate("3:1", "２０６号"),
        AddressCandidate("3:2", "該当する部屋がない方はこちら", "missing_address"),
    ]
    assert JcomSimulationService._is_room_choice(
        apartment, rooms, building_selected=True
    )
    assert not JcomSimulationService._is_room_choice(
        apartment, rooms, building_selected=False
    )
    assert not JcomSimulationService._is_room_choice(
        detached, rooms, building_selected=True
    )


def test_manual_candidates_put_site_next_last_and_exclude_missing_link():
    candidates = [
        AddressCandidate("3:0", "【表示中の住所】で次へ", "next_with_current"),
        AddressCandidate("3:1", "２０５号"),
        AddressCandidate("3:2", "該当する部屋がない方はこちら", "missing_address"),
        AddressCandidate("3:3", "２０６号"),
    ]
    selected = JcomSimulationService._manual_choice_candidates(candidates)
    assert [item.candidate_id for item in selected] == ["3:1", "3:3", "3:0"]


def test_button_text_lookup_uses_one_dom_query_and_clicks_same_element():
    class ButtonDriver:
        current_url = "https://onlineshop.jcom.co.jp/"

        def __init__(self):
            self.button = object()
            self.calls = []

        def execute_script(self, script, *args):
            self.calls.append((script, args))
            if "querySelectorAll" in script:
                return self.button
            assert args == (self.button,)
            return None

    driver = ButtonDriver()
    service = JcomSimulationService(threading.Event(), lambda message: None)
    service.driver = driver
    assert service._click_button_with_text(("エリアを設定する",)) is driver.button
    assert len(driver.calls) == 2
    assert driver.calls[0][1] == (None, ("エリアを設定する",))
