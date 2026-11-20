"""プレディクティブ設定のテスト

プレディクティブのリストでは別システムが発信するため、CTIの状態表示が
「発信中→通話中」に変わらない。設定をオンにすると、CTIの顧客の切り替わりで
提供判定を自動で始める。
"""

import json
import os
import time
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication
from PySide6.QtWidgets import QApplication, QWidget

from services.cti_status_monitor import CTIStatus, CTIStatusMonitor
from ui.main_window import MainWindow
from ui.settings_dialog import SettingsDialog

from test_cti_screen_change_fetch import CTI_ADDRESS, CTI_NAME, _owner

OTHER_NAME = "桃華"
OTHER_ADDRESS = "静岡県掛川市杉谷南１丁目１４−１８"


def app():
    return QApplication.instance() or QApplication([])


def _predictive_owner(settings=None, product="self_collabo"):
    owner = _owner(product)
    owner.settings = {"predictive_mode": True, **(settings or {})}
    owner.scheduled = 0
    owner.cti_screen = (CTI_NAME, CTI_ADDRESS)

    def fetch_cti_data():
        # CTIにいま表示されている顧客を取得する
        owner.fetch_count += 1
        name, address = owner.cti_screen
        data = SimpleNamespace(customer_name=name, address=address, phone="", postal_code="")
        MainWindow.update_form_with_data(owner, data)

    def schedule():
        owner.scheduled += 1

    owner.fetch_cti_data = fetch_cti_data
    owner._schedule_predictive_auto_processing = schedule
    return owner


def _monitor(owner, name=CTI_NAME, address=CTI_ADDRESS):
    if name and address:
        owner.cti_screen = (name, address)
    MainWindow._handle_cti_screen_values(owner, name, address)


def test_customer_switch_starts_area_search_when_predictive_is_on():
    owner = _predictive_owner()
    _monitor(owner)
    assert owner.fetch_count == 1
    # 起動直後に表示されていた顧客では始めない
    assert owner.scheduled == 0

    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.fetch_count == 2
    assert owner.scheduled == 1


def test_customer_switch_only_fetches_when_predictive_is_off():
    owner = _predictive_owner({"predictive_mode": False})
    _monitor(owner)
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.fetch_count == 2
    assert owner.scheduled == 0


def test_first_customer_after_empty_screen_counts_as_a_switch():
    owner = _predictive_owner()
    _monitor(owner, name="", address="")
    assert owner.fetch_count == 0

    _monitor(owner)

    assert owner.fetch_count == 1
    assert owner.scheduled == 1


def test_same_customer_does_not_start_again():
    owner = _predictive_owner()
    _monitor(owner)
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)
    for _ in range(5):
        _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)
    # 画面切替の途中で一瞬空になっても切り替わりとはみなさない
    _monitor(owner, name="", address="")
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.scheduled == 1


def test_refetch_because_inputs_are_empty_is_not_a_switch():
    owner = _predictive_owner()
    _monitor(owner)
    owner.list_name_input.setText("")
    owner.address_input.setText("")

    _monitor(owner)

    assert owner.fetch_count == 2
    assert owner.scheduled == 0


def test_predictive_needs_auto_processing_enabled():
    owner = _predictive_owner({"enable_auto_cti_processing": False})
    _monitor(owner)
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.fetch_count == 2
    assert owner.scheduled == 0


def test_predictive_works_even_when_screen_change_fetch_is_off():
    owner = _predictive_owner({"enable_cti_screen_change_fetch": False})
    _monitor(owner)
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.fetch_count == 2
    assert owner.scheduled == 1


def test_screen_change_fetch_off_without_predictive_does_nothing():
    owner = _predictive_owner({"predictive_mode": False, "enable_cti_screen_change_fetch": False})
    _monitor(owner)
    _monitor(owner, name=OTHER_NAME, address=OTHER_ADDRESS)

    assert owner.fetch_count == 0
    assert owner.scheduled == 0


def _runner_owner(threshold=0):
    owner = SimpleNamespace(
        settings={"predictive_mode": True, "call_duration_threshold": threshold},
        is_auto_processing=False,
        searches=0,
    )

    def auto_search_service_area():
        owner.searches += 1

    owner.auto_search_service_area = auto_search_service_area
    owner._run_predictive_auto_processing = (
        lambda generation: MainWindow._run_predictive_auto_processing(owner, generation)
    )
    return owner


def _process_events_for(seconds):
    deadline = time.time() + seconds
    while time.time() < deadline:
        QCoreApplication.processEvents()
        time.sleep(0.01)


def test_scheduled_search_starts_after_threshold():
    app()
    owner = _runner_owner(threshold=0.2)

    MainWindow._schedule_predictive_auto_processing(owner)
    _process_events_for(0.05)
    assert owner.searches == 0

    _process_events_for(0.4)
    assert owner.searches == 1
    assert owner.is_auto_processing is True


def test_switching_again_while_waiting_cancels_the_earlier_search():
    app()
    owner = _runner_owner(threshold=0.2)

    MainWindow._schedule_predictive_auto_processing(owner)
    MainWindow._schedule_predictive_auto_processing(owner)
    _process_events_for(0.5)

    assert owner.searches == 1


def test_does_not_start_while_previous_auto_search_is_running():
    owner = _runner_owner()
    owner._predictive_generation = 1
    owner.is_auto_processing = True

    MainWindow._run_predictive_auto_processing(owner, 1)

    assert owner.searches == 0


def test_result_button_cancels_a_waiting_search():
    owner = _runner_owner()
    owner._predictive_generation = 3
    owner._predictive_scheduled_at = time.time() - 5

    MainWindow.on_cancel_processing_request(owner, "NG")
    MainWindow._run_predictive_auto_processing(owner, 3)

    assert owner.searches == 0


def test_result_button_keeps_the_search_for_the_customer_who_just_connected():
    # 「NG」のクリック検出が遅れて、次の顧客の予定より後に届いた場合
    owner = _runner_owner()
    owner._predictive_generation = 4
    owner._predictive_scheduled_at = time.time()

    MainWindow.on_cancel_processing_request(owner, "NG")
    MainWindow._run_predictive_auto_processing(owner, 4)

    assert owner.searches == 1


def test_auto_search_with_missing_address_releases_the_auto_flag():
    owner = SimpleNamespace(
        current_product="self_collabo",
        is_auto_processing=True,
        postal_code_input=SimpleNamespace(text=lambda: ""),
        address_input=SimpleNamespace(text=lambda: ""),
    )
    owner._is_unlisted_self_collabo = lambda: False

    MainWindow.auto_search_service_area(owner)

    assert owner.is_auto_processing is False


def _status_monitor(tmp_path, monkeypatch, predictive):
    (tmp_path / "settings.json").write_text(
        json.dumps({"predictive_mode": predictive, "enable_cti_screen_change_fetch": False}),
        encoding="utf-8",
    )
    monkeypatch.chdir(tmp_path)
    calls = []
    monitor = CTIStatusMonitor(
        on_dialing_to_talking_callback=lambda: calls.append("auto"),
        on_screen_changed_callback=lambda *values: calls.append(values),
    )
    return monitor, calls


def test_status_change_does_not_start_auto_processing_when_predictive_is_on(tmp_path, monkeypatch):
    monitor, calls = _status_monitor(tmp_path, monkeypatch, predictive=True)
    monitor._detect_status_change(CTIStatus.WAITING)
    monitor._detect_status_change(CTIStatus.DIALING)
    monitor._detect_status_change(CTIStatus.TALKING)

    assert "auto" not in calls


def test_status_change_still_starts_auto_processing_when_predictive_is_off(tmp_path, monkeypatch):
    monitor, calls = _status_monitor(tmp_path, monkeypatch, predictive=False)
    monitor._detect_status_change(CTIStatus.WAITING)
    monitor._detect_status_change(CTIStatus.DIALING)
    monitor._detect_status_change(CTIStatus.TALKING)

    assert "auto" in calls


def test_screen_is_watched_for_predictive_even_when_screen_change_fetch_is_off(tmp_path, monkeypatch):
    for predictive, expected in ((True, [(CTI_NAME, CTI_ADDRESS)]), (False, [])):
        monitor, calls = _status_monitor(tmp_path, monkeypatch, predictive=predictive)
        monitor.window_handle = 1
        monitor._get_customer_values = lambda: (CTI_NAME, CTI_ADDRESS)
        monkeypatch.setattr("services.cti_status_monitor.win32gui.IsWindow", lambda _hwnd: True)

        monitor._check_screen_change()

        assert calls == expected


def test_settings_dialog_saves_and_restores_predictive_mode(tmp_path):
    app()
    path = tmp_path / "settings.json"
    path.write_text(json.dumps({"mode": "simple", "show_mode_selection": False}), encoding="utf-8")
    parent = QWidget()
    parent.settings_file = str(path)
    parent.current_mode = "simple"

    dialog = SettingsDialog(parent)
    assert dialog.predictive_mode_checkbox.isChecked() is False
    dialog.predictive_mode_checkbox.setChecked(True)
    assert dialog.save_settings()
    dialog.close()

    assert json.loads(path.read_text(encoding="utf-8"))["predictive_mode"] is True

    reopened = SettingsDialog(parent)
    assert reopened.predictive_mode_checkbox.isChecked() is True
    reopened.reset_cti_settings()
    assert reopened.predictive_mode_checkbox.isChecked() is False
    reopened.close()
    parent.close()


def test_predictive_checkbox_follows_auto_processing_checkbox(tmp_path):
    app()
    path = tmp_path / "settings.json"
    path.write_text(json.dumps({"enable_auto_cti_processing": False}), encoding="utf-8")
    parent = QWidget()
    parent.settings_file = str(path)
    parent.current_mode = "simple"

    dialog = SettingsDialog(parent)
    assert dialog.predictive_mode_checkbox.isEnabled() is False
    dialog.cti_auto_processing_checkbox.setChecked(True)
    assert dialog.predictive_mode_checkbox.isEnabled() is True
    dialog.close()
    parent.close()
