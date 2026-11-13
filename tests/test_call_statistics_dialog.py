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
    assert [dialog.table.horizontalHeaderItem(i).text() for i in range(dialog.table.columnCount())] == \
        ["時間帯", "コール", "対話", "トス", "応答時間", "新規見込", "対話NG", "取得状態"]
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
    assert dialog.table.item(0, 4).text() == "00:01:00"
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

        def collect(self, unit, full=True):
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
    assert dialog.table.item(1, 1).text() == "1"
    dialog.close()


def test_approach_ng_with_null_id_appears_in_the_hourly_table(tmp_path):
    from test_call_statistics import history_row, history_snapshot
    app = QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    rows = [history_row(call_id="(Null)"),
            {**history_row("16:34", "NG", "7", "(Null)"), "架電結果詳細": "アプローチNG"}]
    result = history_snapshot("2026-11-13T16:40:00", rows)
    controller.store.save(result)
    dialog = CallStatisticsDialog(controller)
    dialog.day.setDate(QDate(2026, 11, 13))
    controller._completed(result, "")
    index = next(i for i in range(dialog.table.rowCount()) if dialog.table.item(i, 0).text().startswith("16:00"))
    assert dialog.table.item(index, 4).text() == "00:00:07"
    assert dialog.table.item(index, 1).text() == "1"
    assert dialog.table.item(index, 6).text() == "1"
    assert "集計済み" in dialog.status_label.text()
    dialog.close()


def test_hourly_run_sends_saved_hours_and_reports_sheet_result(tmp_path, monkeypatch):
    from test_call_statistics import counter_snapshot, history_row, history_snapshot
    from ui import call_statistics_dialog as module
    app = QApplication.instance() or QApplication([])
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(counter_snapshot("2026-11-13T12:00:02", 1))
    result = history_snapshot("2026-11-13T12:30:02", [history_row("12:20")])
    result = type(result)(result.acquired_at, result.staff_id, result.staff_name, "cti-1",
                          counter_snapshot("2026-11-13T12:30:02", 3).counters, result.talk_seconds, 1, result.history_counters, result.first_history_hour)
    sent = []

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit, full=True):
            return result

    class Sender:
        def send(self, day, staff_id, staff_name, hours):
            sent.append((day, staff_name, hours))
            return {"ok": True, "skipped": [12]}

    monkeypatch.setattr(module, "CtiStatisticsCollector", Collector)
    controller = CallStatisticsController(store=store, autostart=False, sender_factory=Sender)
    controller.set_automatic(False)
    controller.collect()
    assert controller.worker.wait(3000)
    app.processEvents()
    assert sent == [("2026-11-13", "姓　名", [{"hour": 12, "calls": 1, "dialogues": 0, "toss": 0}])]
    assert "シートに枠が無い時間帯：12時" in controller.status


def test_counter_only_run_does_not_send_and_queues_a_full_run(tmp_path, monkeypatch):
    import threading
    from test_call_statistics import counter_snapshot
    from ui import call_statistics_dialog as module
    app = QApplication.instance() or QApplication([])
    release = threading.Event()
    calls = []

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit, full=True):
            calls.append(full)
            if not full:
                release.wait(2)
            return counter_snapshot(f"2026-11-13T12:0{len(calls)}:02", len(calls))

    class Sender:
        def send(self, *args):
            calls.append("sent")
            return {"ok": True}

    monkeypatch.setattr(module, "CtiStatisticsCollector", Collector)
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False, sender_factory=Sender)
    controller.set_automatic(False)
    controller.collect(full=False)
    first = controller.worker
    controller.collect(full=True)  # 実行中に毎時の集計が重なった
    release.set()
    assert first.wait(3000)
    for _ in range(50):
        app.processEvents()
        if controller.worker is not None and controller.worker is not first:
            break
    assert controller.worker is not None and controller.worker.wait(3000)
    app.processEvents()
    assert calls[:2] == [False, True]
