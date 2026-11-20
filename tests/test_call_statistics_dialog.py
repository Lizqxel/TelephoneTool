import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QDate
from PySide6.QtWidgets import QApplication

from services.call_statistics import StatisticsStore
from ui.call_statistics_dialog import CallStatisticsController, CallStatisticsDialog
import pytest


@pytest.fixture(autouse=True)
def no_sheet_sending(monkeypatch):
    # sender_factory を渡さないテストでも、本物の送信用フォームへ送らない。
    from ui import call_statistics_dialog as module
    monkeypatch.setattr(module.SheetSender, "from_settings", classmethod(lambda cls: None))


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

        def collect(self, unit):
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


def test_collection_requested_while_running_is_queued_and_runs_after(tmp_path, monkeypatch):
    import threading
    from test_call_statistics import counter_snapshot
    from ui import call_statistics_dialog as module
    app = QApplication.instance() or QApplication([])
    release = threading.Event()
    calls = []

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit):
            calls.append("collect")
            if len(calls) == 1:
                release.wait(2)
            return counter_snapshot(f"2026-11-13T12:0{len(calls)}:02", len(calls))

        def clear_history(self):
            pass

    monkeypatch.setattr(module, "CtiStatisticsCollector", Collector)
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False, sender_factory=lambda: None)
    controller.set_automatic(False)
    controller.collect()
    first = controller.worker
    assert controller.busy  # 集計を始めた時点で「集計中」になっている
    controller.collect()  # 実行中に重なった
    release.set()
    assert first.wait(3000)
    for _ in range(50):
        app.processEvents()
        if controller.worker is not None and controller.worker is not first:
            break
    assert controller.worker is not None and controller.worker.wait(3000)
    app.processEvents()
    assert calls == ["collect", "collect"]


def test_status_change_is_announced_only_after_the_collection_is_running(tmp_path, monkeypatch):
    import threading
    from ui import call_statistics_dialog as module
    app = QApplication.instance() or QApplication([])
    release = threading.Event()

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit):
            release.wait(2)
            raise RuntimeError("CTIがありません")

    monkeypatch.setattr(module, "CtiStatisticsCollector", Collector)
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False, sender_factory=lambda: None)
    controller.set_automatic(False)
    seen = []
    controller.changed.connect(lambda: seen.append(controller.busy))
    controller.collect()
    worker = controller.worker
    assert seen and seen[0] is True  # 最初の知らせの時点で busy
    release.set()
    assert worker.wait(3000)
    app.processEvents()


@pytest.fixture
def today_is_13th(monkeypatch):
    """テストのデータ（2026-11-13）を「今日」として画面を出す。"""
    from ui import call_statistics_dialog as module

    class FixedDate(QDate):
        @staticmethod
        def currentDate():
            return QDate(2026, 11, 13)

    monkeypatch.setattr(module, "QDate", FixedDate)


def test_hour_of_the_last_collection_is_marked_in_progress(tmp_path, today_is_13th):
    from test_call_statistics import history_row, history_snapshot
    QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    # 13:00に集計した時点では13時台はまだ0件。終わった時間帯の0件と見分けられるようにする。
    controller.store.save(history_snapshot("2026-11-13T13:00:02", [history_row("12:45")]))
    dialog = CallStatisticsDialog(controller)
    dialog.day.setDate(QDate(2026, 11, 13))
    rows = {dialog.table.item(i, 0).text()[:5]: i for i in range(dialog.table.rowCount())}
    assert dialog.table.item(rows["13:00"], 1).text() == "0"
    assert dialog.table.item(rows["13:00"], 7).text().endswith("途中（13:00時点）")
    assert "途中" not in dialog.table.item(rows["12:00"], 7).text()
    dialog.close()


def test_past_day_button_collects_that_day_and_shows_history_only_totals(tmp_path, today_is_13th):
    from dataclasses import replace
    from datetime import datetime
    from services.call_statistics import JST
    from test_call_statistics import history_row, history_snapshot
    QApplication.instance() or QApplication([])
    controller = CallStatisticsController(store=StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    requested = []
    controller.collect = lambda day=None: requested.append(day)
    dialog = CallStatisticsDialog(controller)
    dialog.day.setDate(QDate(2026, 11, 12))
    assert dialog.collect_button.text() == "この日を集計"
    assert "「この日を集計」" in dialog.total_label.text()
    dialog.collect_button.click()
    dialog.day.setDate(QDate(2026, 11, 13))
    assert dialog.collect_button.text() == "今すぐ集計"
    dialog.collect_button.click()
    dialog.day.setDate(QDate(2026, 11, 14))
    assert not dialog.collect_button.isEnabled()  # 先の日付は集計しない
    assert requested == ["2026-11-12", None]
    # 過去日の集計結果（上部カウンター無し）でも表と合計を出し、「途中」は付けない。
    rows = [{**history_row("10:05", call_id="a"), "対応日時": "2026/11/12 10:05"},
            {**history_row("18:40", call_id="b"), "対応日時": "2026/11/12 18:40"}]
    base = history_snapshot("2026-11-13T00:00:00", [])
    from services.call_statistics import hourly_history_totals
    totals = hourly_history_totals(rows, "2026-11-12", base.staff_name)
    past = replace(base, acquired_at=datetime(2026, 11, 12, 23, 59, 59, tzinfo=JST), history_rows=2,
                   counters={key: None for key in base.counters}, talk_seconds=totals["talk_seconds"],
                   history_counters=totals["counters"], first_history_hour=totals["first_hour"], history_only=True)
    controller.store.save(past)
    dialog.day.setDate(QDate(2026, 11, 12))
    assert dialog.staff.currentData() == base.staff_id
    texts = {dialog.table.item(i, 0).text()[:5]: dialog.table.item(i, 7).text() for i in range(dialog.table.rowCount())}
    assert "途中" not in "".join(texts.values())
    assert dialog.total_label.text().startswith("この日の合計　コール 2")
    assert "CTI上部" not in dialog.total_label.text()
    controller._completed(past, "", "")
    assert controller.status.startswith("11/12分を集計済み")
    dialog.close()
