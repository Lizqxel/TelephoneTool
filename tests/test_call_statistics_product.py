"""架電履歴の商材が未選択のとき、検索の間だけ唯一の商材を選び、終わったら元に戻す。

CTIは商材が未選択だと架電結果詳細（現アナNGなど）を空欄で返すため、対話無NGを数えられない。
本物のCTIは操作しないので、メッセージに答える偽の画面で順番と後始末を確かめる。
"""

import ctypes
from datetime import datetime
import threading

import pytest

from services import cti_statistics_service as service
from services.call_statistics import JST, hourly_history_totals
from services.cti_statistics_service import current_list_product, find_labeled_combo, product_to_select


class FastEvent(threading.Event):
    """待ち時間なしで読み直す（取消は通常どおり効く）。"""

    def wait(self, timeout=None):
        return self.is_set()


def fake_history_screen(monkeypatch, items, cursel, search_ok=True, main_product="", dates=None, target_day=None, fail_dates=False,
                        no_calls=False, stuck=False):
    """dates は検索日時の欄（開始・終了）の初期表示。省略時は当日の00:00:00〜23:59:59。

    no_calls: 検索結果が0件（CTIは表も件数も変えず、検索ボタンも押せるまま）。
    stuck: 検索後も表が空のまま、検索ボタンは押せなくなっている（0件とは判断しない）。
    """
    today = datetime.now(JST).date().isoformat()
    day = target_day or today
    state = {"cursel": cursel, "searched_with": None, "events": [], "searched_range": None,
             "dates": list(dates or [f"{today} 00:00:00", f"{today} 23:59:59"]), "search_enabled": True}

    def node(hwnd, parent, cls, text="", rect=(0, 0, 0, 0), control_id=None, visible=False):
        return {"hwnd": hwnd, "parent_hwnd": parent, "class_name": cls, "text": text, "rect": list(rect),
                "control_id": control_id or hwnd, "visible": visible, "enabled": False}

    def nodes():
        return [
            node(10, 0, "Window", "CTIメイン", visible=True),
            # 表示中の架電画面：リストと同じ行に「商材」欄がある。
            node(50, 10, "WindowsForms10.Window.8", "架電画面", visible=True),
            node(51, 50, "WindowsForms10.STATIC", "リスト", (340, 90, 380, 110), visible=True),
            node(52, 50, "WindowsForms10.STATIC", "商材", (530, 90, 565, 110), visible=True),
            node(53, 50, "WindowsForms10.EDIT", main_product, (567, 90, 730, 110), visible=True),
            node(11, 10, "WindowsForms10.Window.8", "架電履歴"),
            node(12, 11, "WindowsForms10.Window.8"),
            # CTIと同じく、開始が上・終了が下。
            node(22, 12, "WindowsForms10.SysDateTimePick32", state["dates"][1], (10, 81, 150, 100)),
            node(21, 12, "WindowsForms10.SysDateTimePick32", state["dates"][0], (10, 60, 150, 80)),
            node(23, 12, "WindowsForms10.STATIC", "電話番号", (300, 40, 380, 58)),
            node(24, 12, "WindowsForms10.EDIT", "", (382, 40, 500, 58)),
            node(25, 12, "WindowsForms10.STATIC", "商材", (0, 20, 100, 38)),
            node(26, 12, "WindowsForms10.COMBOBOX", items[state["cursel"]], (102, 20, 300, 40), control_id=0x1A2B26),
            {**node(27, 12, "WindowsForms10.BUTTON", "検索"), "enabled": state["search_enabled"]},
            node(28, 12, "WindowsForms10.BUTTON", "次へ"),
            node(29, 12, "WindowsForms10.BUTTON", "最初"),
            node(30, 12, "WindowsForms10.STATIC", "検索結果件数", (0, 100, 990, 120)),
            node(31, 12, "WindowsForms10.STATIC", str(len(history_rows())) if history_rows() else "", (88, 102, 173, 118)),
            node(40, 11, "WindowsForms10.Window.8"),
            node(41, 11, "WindowsForms10.Window.8"),
        ]

    def history_rows():
        if state["searched_with"] is None or no_calls or stuck:
            return []
        # 商材を選んで検索したときだけ、CTIは詳細を返す。検索日時に入る日の行だけを返す。
        detail = "現アナNG" if items[state["searched_with"]] != "-" else ""
        start, end = (text[:10] for text in state["searched_range"])
        rows = []
        for index, when in enumerate(sorted({day, today})):
            if start <= when <= end:
                stamp = when.replace("-", "/")
                rows += [{"対応日時": f"{stamp} 10:00:00", "IP電話通話時間": "0", "管理番号": f"A{index}", "架電結果": "NG",
                          "架電結果詳細": detail, "対応スタッフ": "姓　名", "通話ユニークID": f"c1-{index}"},
                         {"対応日時": f"{stamp} 10:05:00", "IP電話通話時間": "0", "管理番号": f"B{index}", "架電結果": "留守",
                          "架電結果詳細": "-", "対応スタッフ": "姓　名", "通話ユニークID": f"c2-{index}"}]
        return rows

    def set_date(api, pid, picker, value):
        assert pid == 7
        if fail_dates:
            raise RuntimeError("日時欄に応答がありません")
        state["dates"][0 if picker["hwnd"] == 21 else 1] = f"{value:%Y-%m-%d %H:%M:%S}"
        state["events"].append(("date", "start" if picker["hwnd"] == 21 else "end", f"{value:%Y-%m-%d %H:%M:%S}"))

    def tables(handles):
        # 0件・クリア後の空の表は、検索結果の列（IP電話通話時間など）を持たない。
        headers = (["対応日時", "管理番号", "会社"] if no_calls or stuck else
                   ["対応日時", "IP電話通話時間", "管理番号", "架電結果", "架電結果詳細", "対応スタッフ", "通話ユニークID"])
        return [{"hwnd": 40, "headers": headers, "rows": history_rows()},
                {"hwnd": 41, "headers": ["選択", "スタッフ"], "rows": [{"選択": "True", "スタッフ": "姓　名"}]}]

    class Api:
        def SendMessageTimeoutW(self, hwnd, message, wparam, lparam, flags, timeout, result):
            value = 0
            if hwnd in (21, 22) and message == service.WM_GETTEXT:
                text = state["dates"][0 if hwnd == 21 else 1]
                (ctypes.c_wchar * (len(text) + 1)).from_address(lparam).value = text
                value = len(text)
            elif hwnd == 26 and message == service.CB_GETCOUNT:
                value = len(items)
            elif hwnd == 26 and message == service.CB_GETCURSEL:
                value = state["cursel"]
            elif hwnd == 26 and message == service.CB_GETLBTEXTLEN:
                value = len(items[wparam])
            elif hwnd == 26 and message == service.CB_GETLBTEXT:
                (ctypes.c_wchar * (len(items[wparam]) + 1)).from_address(lparam).value = items[wparam]
                value = len(items[wparam])
            elif hwnd == 26 and message == service.CB_SETCURSEL:
                state["cursel"] = wparam
                state["events"].append(("set", wparam))
                value = wparam
            elif message == service.WM_COMMAND and lparam == 26:
                assert hwnd == 12 and wparam == (service.CBN_SELCHANGE << 16) | 0x2B26
                state["events"].append(("notify", state["cursel"]))
            elif message == service.WM_COMMAND and lparam == 27:
                state["events"].append(("search", state["cursel"]))
                if not search_ok:
                    return 0
                state["searched_with"] = state["cursel"]
                state["searched_range"] = tuple(state["dates"])
                if history_rows() or stuck:
                    state["search_enabled"] = False  # 結果を表示するとクリアまで検索ボタンは押せない
            elif message == service.WM_COMMAND:
                state["events"].append(("button", lparam))
            result._obj.value = value
            return 1

    class Windows:
        api = Api()

        def snapshot(self, hwnd, limit):
            return {"nodes": nodes(), "errors": []}

    monkeypatch.setattr(service, "read_msaa", tables)
    monkeypatch.setattr(service, "msaa_tables", lambda value: value)
    monkeypatch.setattr(service, "set_date_time", set_date)
    collector = service.CtiStatisticsCollector(FastEvent())
    collector.windows = Windows()
    collector.started_day = today
    cti = {"hwnd": 10, "pid": 7}

    def run():
        return collector.collect_history(cti, collector.windows.snapshot(10, 5000), day, "姓 名")

    return collector, state, run, day


def test_unselected_product_is_selected_only_for_the_search_and_restored(monkeypatch):
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 0)
    rows = run()
    assert state["events"] == [("set", 1), ("notify", 1), ("search", 1), ("set", 0), ("notify", 0)]
    assert state["cursel"] == 0
    assert collector.history_notes == [] and collector.product_note == ""
    counts = hourly_history_totals(rows, day, "姓 名")["counters"][10]
    assert counts["ng_without_dialogue"] == 1 and counts["ng_with_dialogue"] == 0 and counts["dialogues"] == 0
    # クリアの比較に使う検索条件は、元に戻した後の状態。
    assert collector._history_context["conditions"][26] == "-"


def test_product_chosen_by_the_operator_is_left_alone(monkeypatch):
    collector, state, run, _ = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1)
    run()
    assert state["events"] == [("search", 1)]


def test_product_shown_on_the_call_screen_is_selected_among_several(monkeypatch):
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "商材A", "自社コラボ"], 0, main_product="自社コラボ")
    rows = run()
    assert state["events"] == [("set", 2), ("notify", 2), ("search", 2), ("set", 0), ("notify", 0)]
    assert hourly_history_totals(rows, day, "姓 名")["counters"][10]["ng_without_dialogue"] == 1


def test_several_products_are_not_guessed_and_ng_without_detail_stays_unknown(monkeypatch):
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "商材A", "商材B"], 0)
    rows = run()
    assert state["events"] == [("search", 0)]
    assert "2件" in collector.product_note
    totals = hourly_history_totals(rows, day, "姓 名")
    assert totals["counters"][10]["dialogues"] is None
    assert totals["unknown_dialogue_results"] == {"NG（詳細なし）": 1}


def test_product_is_restored_when_the_search_fails(monkeypatch):
    collector, state, run, _ = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 0, search_ok=False)
    with pytest.raises(RuntimeError, match="検索応答"):
        run()
    assert state["events"] == [("set", 1), ("notify", 1), ("search", 1), ("set", 0), ("notify", 0)]
    assert state["cursel"] == 0


def test_product_is_restored_even_when_the_job_is_cancelled(monkeypatch):
    collector, state, run, _ = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 0)
    original = collector.select_choice

    def select_then_cancel(combo, index):
        original(combo, index)
        collector.cancel.set()  # 選んだ直後に取消（検索ボタンは押さない）

    collector.select_choice = select_then_cancel
    with pytest.raises(service.CollectionCancelled):
        run()
    assert ("search", 1) not in state["events"]
    assert state["cursel"] == 0


def test_product_combo_is_found_by_its_label_on_the_same_row():
    nodes = [
        {"hwnd": 1, "parent_hwnd": 9, "class_name": "WindowsForms10.STATIC", "text": "商材", "rect": [-817, 311, -717, 329]},
        {"hwnd": 2, "parent_hwnd": 9, "class_name": "WindowsForms10.COMBOBOX", "text": "-", "rect": [-715, 311, -513, 331]},
        {"hwnd": 3, "parent_hwnd": 9, "class_name": "WindowsForms10.COMBOBOX", "text": "BBC", "rect": [-715, 290, -513, 310]},
        {"hwnd": 4, "parent_hwnd": 9, "class_name": "WindowsForms10.COMBOBOX", "text": "-", "rect": [-765, 332, -715, 352]},
    ]
    assert find_labeled_combo(nodes, "商材")["hwnd"] == 2
    assert find_labeled_combo(nodes[1:], "商材") is None


@pytest.mark.parametrize("items, current, preferred, expected", [
    (["-", "自社コラボ"], 0, "", 1),
    (["-", "自社コラボ"], 1, "", None),  # 担当者が選んだ商材はそのまま
    (["-", "商材A", "商材B"], 0, "", None),  # 架電画面の商材が読めなければ推測しない
    (["-", "商材A", "商材B"], 0, "商材Ｂ", 2),  # 架電画面の商材と同じ名前（全角半角は区別しない）
    (["-", "商材A", "商材B"], 0, "商材C", None),
    (["-", "自社コラボ"], 0, "別の商材", 1),  # 同じ名前が無くても、選べるのが1つならそれ
    (["-", "商材A", "商材B"], 1, "商材B", None),
    (["-"], 0, "", None),
    ([], None, "", None),
])
def test_product_is_chosen_from_the_call_screen_or_the_only_choice(items, current, preferred, expected):
    assert product_to_select(items, current, preferred)[0] == expected


def test_call_screen_product_is_read_from_the_row_of_the_current_list():
    def node(hwnd, parent, cls, text, rect, visible=True):
        return {"hwnd": hwnd, "parent_hwnd": parent, "class_name": cls, "text": text, "rect": list(rect), "visible": visible}

    nodes = [node(1, 9, "WindowsForms10.STATIC", "リスト", (340, 90, 380, 110)),
             node(2, 8, "WindowsForms10.STATIC", "商材", (530, 90, 565, 110)),
             node(3, 8, "WindowsForms10.EDIT", "自社コラボ", (567, 90, 730, 110)),
             # 他のタブ（非表示）や別の行の「商材」は使わない。
             node(4, 7, "WindowsForms10.STATIC", "商材", (530, 90, 565, 110), visible=False),
             node(5, 7, "WindowsForms10.EDIT", "別の商材", (567, 90, 730, 110), visible=False),
             node(6, 8, "WindowsForms10.STATIC", "商材", (10, 300, 60, 320)),
             node(10, 8, "WindowsForms10.EDIT", "顧客の入力", (62, 300, 200, 320))]
    assert current_list_product(nodes) == "自社コラボ"
    assert current_list_product(nodes[1:]) == ""


def test_collect_explains_why_ng_could_not_be_judged(monkeypatch):
    from test_call_statistics import history_row

    class Windows:
        def windows(self, title):
            return [{"hwnd": 10, "pid": 7}]

        def snapshot(self, hwnd, limit):
            return {"nodes": [], "errors": []}

    class Clock(datetime):
        @classmethod
        def now(cls, tz=None):
            return datetime(2026, 11, 13, 14, 0, 2, tzinfo=JST)

    monkeypatch.setattr(service, "Windows", Windows)
    monkeypatch.setattr(service, "read_counters", lambda nodes: {})
    monkeypatch.setattr(service, "counter_context", lambda nodes: "list")
    monkeypatch.setattr(service, "datetime", Clock)
    collector = service.CtiStatisticsCollector()
    monkeypatch.setattr(collector, "read_login_staff", lambda *args: ("姓　名", "bb1"))
    blank_ng = {**history_row(outcome="NG"), "架電結果詳細": ""}

    def history(*args):
        collector.product_note = "架電履歴の商材が未選択で、選べる商材が2件あるため、自動で選んでいません。"
        return [blank_ng]

    monkeypatch.setattr(collector, "collect_history", history)
    result = collector.collect()
    assert "NG（詳細なし） 1件" in result.note
    assert "選べる商材が2件" in result.note
    assert result.history_counters[13]["dialogues"] is None


def test_past_day_is_searched_with_that_day_only_and_the_period_is_restored(monkeypatch):
    from datetime import timedelta
    today = datetime.now(JST).date()
    past = (today - timedelta(days=1)).isoformat()
    wide = ["2026-09-30 00:00:00", f"{today.isoformat()} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 0, dates=wide, target_day=past)
    rows = run()
    assert state["events"] == [
        ("date", "start", f"{past} 00:00:00"), ("date", "end", f"{past} 23:59:59"),
        ("set", 1), ("notify", 1), ("search", 1), ("set", 0), ("notify", 0),
        ("date", "start", wide[0]), ("date", "end", wide[1])]
    assert state["searched_range"] == (f"{past} 00:00:00", f"{past} 23:59:59")
    assert state["dates"] == wide
    assert {r["対応日時"][:10] for r in rows} == {past.replace("-", "/")}
    # クリアの比較に使う検索条件は、元に戻した後の状態。
    assert collector._history_context["conditions"][21] == wide[0]
    assert collector.history_notes == []


def test_wide_period_is_narrowed_to_today_for_a_quick_search(monkeypatch):
    today = datetime.now(JST).date().isoformat()
    wide = ["2026-09-30 00:00:00", f"{today} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, dates=wide)
    run()
    assert state["searched_range"] == (f"{today} 00:00:00", f"{today} 23:59:59")
    assert state["dates"] == wide


def test_period_is_restored_when_the_search_fails(monkeypatch):
    today = datetime.now(JST).date().isoformat()
    wide = ["2026-09-30 00:00:00", f"{today} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, dates=wide, search_ok=False)
    with pytest.raises(RuntimeError, match="検索応答"):
        run()
    assert state["dates"] == wide
    assert state["events"][-2:] == [("date", "start", wide[0]), ("date", "end", wide[1])]


def test_when_the_period_cannot_be_changed_the_current_period_is_used(monkeypatch):
    today = datetime.now(JST).date().isoformat()
    wide = ["2026-09-30 00:00:00", f"{today} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, dates=wide, fail_dates=True)
    rows = run()
    assert state["searched_range"] == tuple(wide)
    assert len(rows) == 2 and "検索日時を変えられなかった" in collector.history_notes[0]


def test_past_day_outside_an_unchangeable_period_is_skipped(monkeypatch):
    from datetime import timedelta
    today = datetime.now(JST).date()
    past = (today - timedelta(days=1)).isoformat()
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, target_day=past, fail_dates=True)
    with pytest.raises(service.CollectionSkipped):
        run()
    assert ("search", 1) not in state["events"]


def test_collect_for_a_past_day_keeps_counters_out_and_marks_history_only(monkeypatch):
    from test_call_statistics import history_row
    from datetime import timedelta

    class Windows:
        def windows(self, title):
            return [{"hwnd": 10, "pid": 7}]

        def snapshot(self, hwnd, limit):
            return {"nodes": [], "errors": []}

    monkeypatch.setattr(service, "Windows", Windows)
    monkeypatch.setattr(service, "read_counters", lambda nodes: pytest.fail("過去日では上部カウンターを読まない"))
    monkeypatch.setattr(service, "counter_context", lambda nodes: "list")
    collector = service.CtiStatisticsCollector()
    monkeypatch.setattr(collector, "read_login_staff", lambda *args: ("姓　名", "bb1"))
    past = (datetime.now(JST).date() - timedelta(days=2)).isoformat()
    seen = []

    def history(cti, before, day, staff_name):
        seen.append(day)
        return [{**history_row(outcome="NG"), "架電結果詳細": "現アナNG", "対応日時": f"{past.replace('-', '/')} 13:10"}]

    monkeypatch.setattr(collector, "collect_history", history)
    result = collector.collect(day=past)
    assert seen == [past]
    assert result.history_only and result.acquired_at.isoformat() == f"{past}T23:59:59+09:00"
    assert all(value is None for value in result.counters.values())
    assert result.history_counters[13]["ng_without_dialogue"] == 1 and result.history_rows == 1
    with pytest.raises(RuntimeError, match="先の日付"):
        collector.collect(day=(datetime.now(JST).date() + timedelta(days=1)).isoformat())


def past_day_snapshot(day="2026-11-12", staff_id="staff-1"):
    from services.call_statistics import StatisticsSnapshot
    from test_call_statistics import history_row
    rows = [{**history_row("10:05", call_id="a"), "対応日時": f"{day.replace('-', '/')} 10:05"},
            {**history_row("18:40", "NG", "30", call_id="b"), "架電結果詳細": "申込意思NG", "対応日時": f"{day.replace('-', '/')} 18:40"}]
    totals = hourly_history_totals(rows, day, "姓　名", ranks={})
    halves = hourly_history_totals(rows, day, "姓　名", ranks={}, interval_minutes=30)
    return StatisticsSnapshot(datetime.fromisoformat(f"{day}T23:59:59+09:00"), staff_id, "姓　名", "s",
                              {key: None for key in ("total_calls", "new_prospects", "ng_with_dialogue", "ng_without_dialogue", "valid_calls")},
                              totals["talk_seconds"], 2, totals["counters"], totals["first_hour"],
                              intervals={m: {**c, "dialogue_seconds": halves["dialogue_seconds"][m]} for m, c in halves["counters"].items()},
                              history_only=True)


def test_past_day_is_saved_from_history_without_touching_counter_baselines(tmp_path):
    from services.call_statistics import StatisticsStore
    from test_call_statistics import counter_snapshot
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(counter_snapshot("2026-11-13T10:00:02", 3))
    store.save(past_day_snapshot())
    store.save(past_day_snapshot())  # 同じ日をもう一度集計しても保存できる
    assert store.latest("staff-1", "2026-11-12") is None  # 上部カウンターの基準は作らない
    assert store.latest("staff-1", "2026-11-13")["valid_calls"] == 3
    hours = {row["hour"]: row for row in store.display_hours("staff-1", "2026-11-12")}
    assert hours[10]["total_calls"] == 1 and hours[18]["dialogues"] == 1 and hours[18]["talk_seconds"] == 30
    assert hours[23]["total_calls"] == 0  # その日の終わりまで確定
    intervals = store.intervals("staff-1", "2026-11-12")
    assert len(intervals) == 48 and {row["observed_at"] for row in intervals} == {"2026-11-12T23:59:59+09:00"}
    assert [person["staff_id"] for person in store.staff()] == ["staff-1"]


def test_past_day_only_staff_appears_in_the_staff_list(tmp_path):
    from services.call_statistics import StatisticsStore
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(past_day_snapshot(staff_id="bb9"))
    assert [(p["staff_id"], p["staff_name"]) for p in store.staff()] == [("bb9", "姓　名")]


def test_worker_sends_a_past_day_with_all_slots_up_to_the_end_of_that_day(tmp_path):
    from PySide6.QtWidgets import QApplication
    from services.call_statistics import StatisticsStore
    from ui import call_statistics_dialog as ui
    QApplication.instance() or QApplication([])
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    asked, sent = [], []

    class Collector:
        def __init__(self, cancel):
            pass

        def collect(self, unit, day=None):
            asked.append(day)
            return past_day_snapshot(day)

        def clear_history(self):
            pass

    class Sender:
        def send(self, day, staff_id, staff_name, hours, intervals=None, observed_at=None):
            sent.append((day, hours, len(intervals), observed_at))
            return {"ok": True}

    original = ui.CtiStatisticsCollector
    ui.CtiStatisticsCollector = Collector
    try:
        ui.StatisticsWorker(store, "seconds", lambda: Sender(), day="2026-11-12").run()
    finally:
        ui.CtiStatisticsCollector = original
    assert asked == ["2026-11-12"]
    day, hours, slots, observed = sent[0]
    assert day == "2026-11-12" and slots == 48 and observed == "2026-11-12T23:59:59+09:00"
    assert {h["hour"]: (h["calls"], h["dialogues"]) for h in hours}[18] == (1, 1)


def test_past_day_payload_is_accepted_by_gas_and_read_back(tmp_path):
    import json
    import subprocess
    from services import team_statistics as team
    from services.call_statistics import StatisticsStore
    from services.call_statistics_sheet import SheetSender, sheet_hours, sheet_intervals
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(past_day_snapshot("2026-11-13"))  # GASのテスト用シートにある日付
    requests = []
    sender = SheetSender({"tokenEntry": "token", "dataEntry": "data", "formResponseUrl": "https://example.invalid"}, "t",
                         post=lambda url, data, timeout: requests.append(data) or type("R", (), {"status_code": 200, "url": url})())
    rows = store.intervals("staff-1", "2026-11-13")
    sender.send("2026-11-13", "staff-1", "姓　名", sheet_hours(store.display_hours("staff-1", "2026-11-13")),
                intervals=sheet_intervals(rows), observed_at=rows[-1]["observed_at"])
    path, stored_path = tmp_path / "payload.json", tmp_path / "stored.json"
    path.write_text(requests[0]["data"], encoding="utf-8")
    checked = subprocess.run(["node", "tests/test_call_statistics_gas.cjs", "--roundtrip", str(path), str(stored_path)],
                             capture_output=True, text=True, encoding="utf-8", timeout=30)
    assert checked.returncode == 0, checked.stdout + checked.stderr
    people = team.parse_intervals(json.loads(stored_path.read_text(encoding="utf-8")), "2026-11-13")
    assert len(people) == 1 and len(people[0].intervals) == 48
    assert people[0].intervals[18 * 60 + 30][:2] == (1, 1)


def test_day_without_calls_is_counted_as_zero_not_an_error(monkeypatch):
    from datetime import timedelta
    today = datetime.now(JST).date()
    past = (today - timedelta(days=1)).isoformat()
    wide = ["2026-09-30 00:00:00", f"{today.isoformat()} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, dates=wide, target_day=past, no_calls=True)
    rows = run()
    assert rows == [] and state["dates"] == wide
    assert hourly_history_totals(rows, past, "姓 名")["counters"][10]["total_calls"] == 0
    pressed = len(state["events"])
    collector.clear_history()  # 0件の一覧はクリアしない（クリアボタンも押せない）
    assert len(state["events"]) == pressed and collector.clear_history_info["skipped"]


def test_empty_table_with_a_locked_search_button_is_still_an_error(monkeypatch):
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 1, stuck=True)
    with pytest.raises(RuntimeError, match="検索後の表"):
        run()


def test_period_and_product_are_restored_when_cancelled_during_the_search(monkeypatch):
    from datetime import timedelta
    today = datetime.now(JST).date()
    past = (today - timedelta(days=1)).isoformat()
    wide = ["2026-09-30 00:00:00", f"{today.isoformat()} 23:59:59"]
    collector, state, run, day = fake_history_screen(monkeypatch, ["-", "自社コラボ"], 0, dates=wide, target_day=past)
    original = collector.press

    def press_then_cancel(button, timeout_ms=15000):
        pressed = original(button, timeout_ms)
        collector.cancel.set()  # 検索ボタンを押した直後に「中止」
        return pressed

    collector.press = press_then_cancel
    with pytest.raises(service.CollectionCancelled):
        run()
    assert state["dates"] == wide and state["cursel"] == 0
    assert state["events"][-4:] == [("set", 0), ("notify", 0), ("date", "start", wide[0]), ("date", "end", wide[1])]
