from copy import deepcopy
import threading

import pytest
from PySide6.QtWidgets import QApplication

from services import cti_statistics_service as service
from ui import call_statistics_dialog as ui
from test_call_statistics import history_row, history_snapshot


@pytest.fixture(autouse=True)
def no_sheet_sending(monkeypatch):
    # テスト中に本物の送信用フォームへ送らない。
    monkeypatch.setattr(ui.SheetSender, "from_settings", classmethod(lambda cls: None))


@pytest.mark.parametrize("failure", [None, "collect", "save", "clear"])
def test_clear_runs_only_after_successful_collection_and_save(monkeypatch, failure):
    application = QApplication.instance() or QApplication([])
    events, completed = [], []
    result = history_snapshot("2026-11-13T16:40:00", [history_row()])

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit, full=True):
            events.append("collect")
            if failure == "collect":
                raise RuntimeError("collect failed")
            return result

        def clear_history(self):
            events.append("clear")
            if failure == "clear":
                raise RuntimeError("clear failed")

    class Store:
        def save(self, snapshot):
            assert snapshot is result
            events.append("save")
            if failure == "save":
                raise RuntimeError("save failed")

    monkeypatch.setattr(ui, "CtiStatisticsCollector", Collector)
    worker = ui.StatisticsWorker(Store(), "seconds")
    worker.completed.connect(lambda snapshot, error, notice, full: completed.append((snapshot, error, notice)))
    worker.run()
    if failure == "collect":
        assert events == ["collect"]
        assert completed[0][0] is None
    elif failure == "save":
        assert events == ["collect", "save"]
        assert completed[0][0] is None
    else:
        assert events == ["collect", "save", "clear"]
        assert completed[0][0] is result and completed[0][1] == ""
        assert ("クリア未完了" in completed[0][2]) == (failure == "clear")


def test_clear_warning_keeps_successful_aggregation_visible(tmp_path):
    application = QApplication.instance() or QApplication([])
    controller = ui.CallStatisticsController(store=ui.StatisticsStore(tmp_path / "stats.sqlite3"), autostart=False)
    controller.set_automatic(False)
    result = history_snapshot("2026-11-13T16:40:00", [history_row()])
    controller.store.save(result)
    controller._completed(result, "", "架電履歴のクリア未完了：CTIがありません", True)
    assert "集計済み" in controller.status
    assert "保存済み" in controller.status
    assert "クリア未完了" in controller.status
    assert "集計できません" not in controller.status
    assert controller.store.hours(result.staff_id, "2026-11-13")[13]["total_calls"] == 1


def clear_fixture(monkeypatch, changed=False, recreated=False):
    rows = [{"call": "saved"}]
    staff_rows = [{"スタッフ": "本人", "選択": "True"}]
    nodes = [
        {"hwnd": 10, "parent_hwnd": 0, "text": "CTIメイン", "class_name": "Window", "visible": True},
        {"hwnd": 11, "parent_hwnd": 10, "text": "架電履歴", "class_name": "WindowsForms10.Window.8", "visible": False},
        {"hwnd": 12, "parent_hwnd": 11, "text": "", "class_name": "WindowsForms10.Window.8", "visible": False},
        {"hwnd": 13, "parent_hwnd": 12, "text": "", "class_name": "WindowsForms10.Window.8", "visible": False},
        {"hwnd": 14, "parent_hwnd": 12, "text": "1 件", "class_name": "WindowsForms10.STATIC", "visible": False},
        {"hwnd": 15, "parent_hwnd": 11, "text": "クリア", "class_name": "WindowsForms10.BUTTON", "control_id": 15, "visible": False},
        # A same-named button outside history must never be invoked.
        {"hwnd": 99, "parent_hwnd": 10, "text": "クリア", "class_name": "WindowsForms10.BUTTON", "control_id": 99, "visible": True},
    ]
    calls = []

    class Api:
        def SendMessageTimeoutW(self, parent, message, control_id, hwnd, *rest):
            calls.append((parent, message, control_id, hwnd))
            return 1

    class Windows:
        api = Api()

        def windows(self, title):
            return [{"hwnd": 10, "pid": 7}]

        def snapshot(self, hwnd, limit):
            data = deepcopy(nodes)
            if calls:
                next(n for n in data if n["hwnd"] == 14)["text"] = "" if recreated else "0 件"
                if recreated:
                    next(n for n in data if n["hwnd"] == 13)["hwnd"] = 31
            return {"nodes": data, "errors": []}

    def tables(handles):
        current = [] if calls else ([{"call": "new query"}] if changed else rows)
        return [{"hwnd": 31 if calls and recreated else 13, "headers": ["対応日時", "管理番号"], "rows": current},
                {"hwnd": 20, "headers": ["選択", "スタッフ"], "rows": staff_rows}]

    monkeypatch.setattr(service, "Windows", Windows)
    monkeypatch.setattr(service, "read_msaa", tables)
    monkeypatch.setattr(service, "msaa_tables", lambda value: value)
    collector = service.CtiStatisticsCollector(threading.Event())
    collector._history_context = {"cti": {"hwnd": 10, "pid": 7}, "page_hwnd": 11, "history_hwnd": 13,
                                  "staff_hwnd": 20, "staff_rows": staff_rows, "rows": rows, "conditions": {}}
    return collector, calls


def test_clear_targets_history_only_and_confirms_empty_result(monkeypatch):
    collector, calls = clear_fixture(monkeypatch)
    collector.clear_history()
    assert calls == [(11, 0x111, 15, 15)]
    assert collector.clear_history_info["rows_after"] == 0
    assert collector.clear_history_info["staff_selection_preserved"] is True
    assert collector.clear_history_info["history_visible_after"] is False
    with pytest.raises(RuntimeError, match="クリア対象"):
        collector.clear_history()
    assert len(calls) == 1


def test_changed_history_is_not_cleared(monkeypatch):
    collector, calls = clear_fixture(monkeypatch, changed=True)
    with pytest.raises(RuntimeError, match="変わった"):
        collector.clear_history()
    assert not calls


def test_clear_accepts_recreated_empty_grid_without_result_count_label(monkeypatch):
    collector, calls = clear_fixture(monkeypatch, recreated=True)
    collector.clear_history()
    assert len(calls) == 1
    assert collector.clear_history_info["rows_after"] == 0
    assert collector.clear_history_info["staff_selection_preserved"] is True


def test_cancelled_job_does_not_clear(monkeypatch):
    collector, calls = clear_fixture(monkeypatch)
    collector.cancel.set()
    with pytest.raises(RuntimeError, match="中止"):
        collector.clear_history()
    assert not calls


def test_history_is_not_cleared_when_this_run_did_not_collect_it(monkeypatch):
    application = QApplication.instance() or QApplication([])
    from test_call_statistics import counter_snapshot
    events = []

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit, full=True):
            return counter_snapshot("2026-11-13T16:40:00", 1)  # 条件が合わずスキップした回

        def clear_history(self):
            events.append("clear")

    class Store:
        def save(self, snapshot):
            events.append("save")

    monkeypatch.setattr(ui, "CtiStatisticsCollector", Collector)
    ui.StatisticsWorker(Store(), "seconds").run()
    assert events == ["save"]
