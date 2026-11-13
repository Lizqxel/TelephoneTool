import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QDate
from PySide6.QtWidgets import QApplication

from services.call_statistics import StatisticsStore
from ui.call_statistics_dialog import CallStatisticsController, CallStatisticsDialog


def test_dialog_displays_unknowns_and_does_not_collect_when_opened(tmp_path):
    app = QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    dialog = CallStatisticsDialog(controller)
    assert dialog.table.rowCount() == 0
    assert dialog.table.horizontalHeaderItem(4).text() == "対話NG"
    assert controller.worker is None
    dialog.close()


def test_failure_keeps_saved_table_and_changes_status(tmp_path):
    from test_call_statistics import snapshot
    app = QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    controller.store.save(snapshot("2026-11-13T13:10:00", 10, 60))
    dialog = CallStatisticsDialog(controller)
    dialog.day.setDate(QDate(2026, 11, 13))
    controller._completed(None, "CTIがありません")
    assert dialog.table.item(0, 0).text().startswith("13:00")
    assert dialog.table.item(0, 1).text() == "00:01:00"
    assert "保存済み" in dialog.status_label.text()
    dialog.close()


def test_preferences_persist_without_modifying_main_settings(tmp_path):
    app = QApplication.instance() or QApplication([])
    path = tmp_path / "stats.sqlite3"
    controller = CallStatisticsController(store=StatisticsStore(path), autostart=False)
    controller.set_automatic(False)
    dialog = CallStatisticsDialog(controller)
    dialog.duration_unit.setCurrentIndex(1)
    fresh = StatisticsStore(path)
    assert fresh.preference("automatic", "1") == "0"
    assert fresh.preference("duration_unit", "seconds") == "minutes"
    assert not controller.timer.isActive()
    dialog.close()


def test_shutdown_cancels_running_collection_without_destroying_thread(tmp_path, monkeypatch):
    import threading
    from ui import call_statistics_dialog as module
    entered = threading.Event()

    class SlowCollector:
        def __init__(self, cancel):
            self.cancel = cancel

        def collect(self, unit):
            entered.set()
            self.cancel.wait(2)
            raise RuntimeError("cancelled")

    monkeypatch.setattr(module, "CtiStatisticsCollector", SlowCollector)
    app = QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.collect()
    worker = controller.worker
    try:
        assert entered.wait(1)
        assert controller.stop() is False
        assert worker.wait(1500)
        app.processEvents()
        assert controller.worker is None
        assert controller.stop() is True
    finally:
        worker.cancelled.set()
        if controller.worker is not None:
            worker.wait(2500)
            app.processEvents()


def test_dialog_shows_midday_logs_without_early_morning_rows(tmp_path):
    from test_call_statistics import history_row, history_snapshot
    app = QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    controller.store.save(history_snapshot("2026-11-13T16:10:00", [history_row("12:20"), history_row(call_id="call-2")]))
    dialog = CallStatisticsDialog(controller)
    dialog.day.setDate(QDate(2026, 11, 13))
    assert dialog.table.rowCount() == 12
    assert dialog.table.item(0, 0).text().startswith("12:00")
    assert dialog.table.item(1, 0).text().startswith("13:00")
    assert dialog.table.item(1, 2).text() == "1"
    dialog.close()
