"""30分集計の境界、保存、送信契約、全員表示を実データ経路に沿って検証する。"""
from dataclasses import replace
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QFontDatabase

from services.call_statistics import JST, StatisticsStore, hourly_history_totals, next_half_hour
from services.call_statistics_sheet import SheetSender, sheet_hours, sheet_intervals
from services import team_statistics as team
from services.team_statistics_report import ranking, report_text, summaries
from ui.team_statistics_analysis import TeamStatisticsAnalysis
from test_call_statistics import history_row, history_snapshot, prospect_row
from test_team_statistics import Response, Session, sheet_grid


def collected(rows, stamp="2026-11-13T13:45:02"):
    base = history_snapshot(stamp, rows)
    halves = hourly_history_totals(rows, "2026-11-13", base.staff_name, ranks={}, interval_minutes=30)
    return replace(base, intervals={minute: {**counts, "dialogue_seconds": halves["dialogue_seconds"][minute]}
                                    for minute, counts in halves["counters"].items()})


def member(staff_id, name, slots, unit="A"):
    return team.IntervalMember(staff_id, name, "部署", unit, slots, "2026-11-13T13:45:02+09:00")


def test_half_hour_boundaries_dialogue_time_zero_calls_and_other_people():
    rows = [history_row("00:00:00", "留守", "5", "a"),
            history_row("10:29:59", "見込", "30.5", "b"),
            history_row("10:30:00", "対話有NG", "60", "c"),
            history_row("10:59:59", "対話無NG", "20", "d"),
            history_row("23:59:59", "継続見込", "0", "e")]
    rows += [rows[1], {**rows[1], "対応スタッフ": "別人"}, {**rows[1], "対応日時": "2026/11/12 10:15"}]
    totals = hourly_history_totals(rows, "2026-11-13", "姓 名", ranks={}, interval_minutes=30)
    assert totals["counters"][600]["total_calls"] == 1
    assert totals["dialogue_seconds"][600] == 30.5
    assert totals["counters"][630]["total_calls"] == 2
    assert totals["counters"][630]["dialogues"] == 1
    assert totals["talk_seconds"][630] == 80 and totals["dialogue_seconds"][630] == 60
    assert totals["dialogue_seconds"][0] == 0 and totals["talk_seconds"][0] == 5
    assert totals["counters"][1410]["dialogues"] == 1
    assert sum(item["total_calls"] for item in totals["counters"].values()) == 5


def test_unknown_outcome_only_invalidates_its_slot_dialogues_and_time():
    rows = [history_row("10:10", "未知", "30", "a"), history_row("10:15", "見込", "60", "b"),
            history_row("10:30", "見込", "20", "c")]
    totals = hourly_history_totals(rows, "2026-11-13", "姓 名", interval_minutes=30)
    assert totals["counters"][600]["total_calls"] == 2
    assert totals["counters"][600]["dialogues"] is None
    assert totals["dialogue_seconds"][600] is None
    assert totals["dialogue_seconds"][630] == 20


def test_toss_stays_once_per_number_in_first_qualifying_half_hour():
    rows = [prospect_row("10:31", "A"), prospect_row("10:01", "A"), prospect_row("11:00", "B", "継続見込")]
    totals = hourly_history_totals(rows, "2026-11-13", "姓 名", ranks={"A": "トス", "B": "トス"}, interval_minutes=30)
    assert totals["counters"][600]["toss"] == 1
    assert totals["counters"][630]["toss"] == totals["counters"][660]["toss"] == 0


@pytest.mark.parametrize("stamp,expected", [("2026-11-13T10:29:59", "2026-11-13T10:30:00"),
                                              ("2026-11-13T10:30:00", "2026-11-13T11:00:00"),
                                              ("2026-11-13T23:59:59", "2026-11-14T00:00:00")])
def test_schedule_uses_next_boundary(stamp, expected):
    assert next_half_hour(datetime.fromisoformat(stamp).replace(tzinfo=JST)) == datetime.fromisoformat(expected).replace(tzinfo=JST)


def test_store_migrates_without_inventing_half_hours_and_preserves_failures(tmp_path):
    path = tmp_path / "stats.sqlite3"
    old = history_snapshot("2026-11-13T13:10:00", [history_row()])
    store = StatisticsStore(path)
    store.save(old)
    assert store.intervals(old.staff_id, "2026-11-13") == []
    result = collected([history_row("13:05", "見込", "90", "a"), history_row("13:40", "留守", "30", "b")])
    store.save(result)
    rows = store.intervals(old.staff_id, "2026-11-13")
    assert len(rows) == 28 and rows[-1]["start_minute"] == 810
    assert rows[-1]["total_calls"] == 1 and rows[-1]["dialogue_seconds"] == 0
    assert rows[-2]["dialogue_seconds"] == 90
    store.save(replace(result, acquired_at=datetime.fromisoformat("2026-11-13T13:50:00+09:00")))
    assert sum(row["total_calls"] for row in store.intervals(old.staff_id, "2026-11-13")) == 2
    saved = store.intervals(old.staff_id, "2026-11-13")
    store.save(replace(result, acquired_at=datetime.fromisoformat("2026-11-13T14:00:02+09:00"), intervals=None))
    assert StatisticsStore(path).intervals(old.staff_id, "2026-11-13") == saved
    with pytest.raises(ValueError, match="古い"):
        store.save(result)
    assert store.intervals(old.staff_id, "2026-11-13") == saved
    assert store.intervals("別人", "2026-11-13") == []
    assert store.intervals(old.staff_id, "2026-11-12") == []


def test_empty_day_is_observed_zero_up_to_capture_not_future(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    result = collected([], "2026-11-13T00:00:02")
    store.save(result)
    rows = store.intervals(result.staff_id, "2026-11-13")
    assert len(rows) == 1 and rows[0]["total_calls"] == rows[0]["dialogue_seconds"] == 0


def test_real_sender_payload_passes_gas_validation(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    result = collected([history_row("13:05", "見込", "90.5", "a")])
    store.save(result)
    requests = []

    def post(url, data, timeout):
        requests.append(data)
        return type("Response", (), {"status_code": 200, "url": url})()

    rows = store.intervals(result.staff_id, "2026-11-13")
    sender = SheetSender({"tokenEntry": "token", "dataEntry": "data", "formResponseUrl": "https://example.invalid"}, "secret-test-token", post=post)
    sender.send("2026-11-13", result.staff_id, result.staff_name, sheet_hours(store.display_hours(result.staff_id, "2026-11-13")),
                intervals=sheet_intervals(rows), observed_at=rows[-1]["observed_at"])
    payload = json.loads(requests[0]["data"])
    assert payload["schemaVersion"] == 2 and "secret-test-token" not in requests[0]["data"]
    assert payload["intervals"][26]["dialogueSeconds"] == 90.5
    path = tmp_path / "payload.json"
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    checked = subprocess.run(["node", "tests/test_call_statistics_gas.cjs", "--payload", str(path)], capture_output=True, text=True, encoding="utf-8", timeout=30)
    assert checked.returncode == 0, checked.stdout + checked.stderr


def test_collector_produces_half_hours_from_the_same_history(monkeypatch):
    from services import cti_statistics_service as service
    rows = [history_row("10:29", "留守", "30", "a"), history_row("10:30", "対話有NG", "60", "b")]

    class Windows:
        def windows(self, title):
            return [{"hwnd": 10, "pid": 7}]

        def snapshot(self, hwnd, limit):
            return {"nodes": [], "errors": []}

    monkeypatch.setattr(service, "Windows", Windows)
    monkeypatch.setattr(service, "read_counters", lambda nodes: {})
    monkeypatch.setattr(service, "counter_context", lambda nodes: "list")
    collector = service.CtiStatisticsCollector()
    monkeypatch.setattr(collector, "read_login_staff", lambda *args: ("姓 名", "bb1"))
    # 時計も固定し、テスト日のCTI履歴を実際のcollect経路へ渡す。
    class Clock(datetime):
        @classmethod
        def now(cls, tz=None):
            return datetime(2026, 11, 13, 10, 45, 2, tzinfo=JST)

    monkeypatch.setattr(service, "datetime", Clock)
    monkeypatch.setattr(collector, "collect_history", lambda *args: rows)
    result = collector.collect()
    assert result.history_rows == 2
    assert result.intervals[600]["dialogue_seconds"] == 0
    assert result.intervals[630]["dialogue_seconds"] == 60
    assert result.history_counters[10]["total_calls"] == 2


def test_worker_sends_saved_half_hours_and_preserves_capture_time_after_history_failure(tmp_path):
    from ui.call_statistics_dialog import StatisticsWorker
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    result = collected([history_row("13:05", "見込", "90", "a")])
    store.save(result)
    failed = replace(result, acquired_at=datetime.fromisoformat("2026-11-13T14:15:02+09:00"), intervals=None)
    store.save(failed)
    sent = []

    class Sender:
        def send(self, day, staff_id, name, hours, **details):
            sent.append(details)
            return {"ok": True}

    worker = StatisticsWorker(store, "seconds", sender_factory=Sender)
    assert "シート送信済み" in worker.send(failed)
    assert sent[0]["observed_at"] == result.acquired_at.isoformat()
    assert len(sent[0]["intervals"]) == 28


class DetailSession(Session):
    def __init__(self, details, status=200):
        super().__init__(sheet_grid())
        self.details, self.detail_status = details, status

    def get(self, url, params, timeout):
        if "通話集計データ_" in params["ranges"]:
            self.ranges.append(params["ranges"])
            assert params["valueRenderOption"] == "UNFORMATTED_VALUE"
            return Response(self.detail_status, {"valueRanges": [{"values": self.details}]} if self.detail_status == 200 else {"error": {"message": "Unable to parse range"}})
        return super().get(url, params, timeout)


def hidden_row(**values):
    row = {"day": "2026-11-13", "staff": "bb1", "name": "岩田悠河", "department": "部署A", "unit": "岩田祐樹", "minute": 600,
           "calls": 3, "dialogues": 2, "toss": 1, "seconds": 90.5, "stamp": "2026-11-13T10:30:02+09:00"}
    row.update(values)
    return list(row.values())


def test_fetches_hidden_sheet_and_keeps_unsent_roster_members_unknown():
    session = DetailSession([team.INTERVAL_HEADERS, hidden_row(), hidden_row(day="2026-11-12")])
    data = team.fetch_day(session, "2026-11-13")
    assert data.interval_error == ""
    assert data.interval_members[0].intervals[600] == (3, 2, 1, 90.5)
    assert len(data.interval_members) == 2 and not data.interval_members[1].intervals
    assert "通話集計データ_2026-11" in session.ranges[0]


def test_missing_hidden_sheet_keeps_legacy_table_with_explicit_message():
    data = team.fetch_day(DetailSession([], 400), "2026-11-13")
    assert len(data.members) == 2 and "30分集計を取得できません" in data.interval_error
    assert all(not member.intervals for member in data.interval_members)


@pytest.mark.parametrize("fail_at", ["A1:D", "rows"])
def test_legacy_timeout_does_not_discard_successful_half_hour_data(fail_at):
    from requests.exceptions import ReadTimeout

    class SlowLegacy(DetailSession):
        def get(self, url, params, timeout):
            assert timeout == (10, 120)
            name, cells = params["ranges"].rsplit("!", 1)
            if "通話集計データ_" not in name and ((fail_at == "A1:D" and cells == "A1:D") or (fail_at == "rows" and cells != "A1:D")):
                raise ReadTimeout()
            return super().get(url, params, timeout)

    session = SlowLegacy([team.INTERVAL_HEADERS, hidden_row()])
    data = team.fetch_day(session, "2026-11-13")
    assert data.interval_members[0].intervals[600] == (3, 2, 1, 90.5)
    assert data.members == [] and "120秒" in data.legacy_error
    assert data.interval_error == ""
    assert "通話集計データ_" in session.ranges[0]


def test_hidden_timeout_keeps_the_legacy_table():
    from requests.exceptions import ReadTimeout

    class SlowHidden(DetailSession):
        def get(self, url, params, timeout):
            if "通話集計データ_" in params["ranges"]:
                raise ReadTimeout()
            return super().get(url, params, timeout)

    data = team.fetch_day(SlowHidden([]), "2026-11-13")
    assert len(data.members) == 2 and data.legacy_error == ""
    assert "120秒" in data.interval_error


def test_timeout_of_both_sources_is_reported_without_zero_results():
    from requests.exceptions import ReadTimeout

    class Offline:
        def get(self, *args, **kwargs):
            raise ReadTimeout()

    with pytest.raises(RuntimeError, match="応答待ち.*120秒"):
        team.fetch_day(Offline(), "2026-11-13")


def test_dialog_can_show_copy_and_recover_half_hour_data_when_legacy_table_failed():
    from test_team_statistics import FakeLogin
    from PySide6.QtCore import QDate
    from ui.team_statistics_dialog import TeamStatisticsDialog
    app = QApplication.instance() or QApplication([])
    person = member("bb1", "岩田悠河", {600: (3, 2, 1, 90)})
    data = team.TeamDay("sheet", "2026-11-13", [], [], [person], legacy_error="応答待ちが時間切れ")
    dialog = TeamStatisticsDialog(login=FakeLogin())
    dialog.day.setDate(QDate(2026, 11, 13))
    dialog._fetched(data, "", data.day)
    assert dialog.analysis.table.rowCount() == 1
    assert "30分集計を取得" in dialog.status.text()
    assert "受信済みの人のみ" in dialog.analysis.message.text()
    assert dialog.row_count() == 0
    dialog.analysis.copy_button.click()
    assert "00:01:30" in QApplication.clipboard().text()
    normal = team.fetch_day(DetailSession([team.INTERVAL_HEADERS, hidden_row()]), data.day)
    dialog._fetched(normal, "", data.day)
    assert dialog.row_count() > 0
    assert "受信済みの人のみ" not in dialog.analysis.message.text()
    dialog.close()


@pytest.mark.parametrize("change", [{"minute": 601}, {"calls": -1}, {"calls": 0.5}, {"seconds": float("nan")},
                                    {"seconds": "bogus"}, {"stamp": "2026-11-13T10:30:00"}, {"staff": ""},
                                    {"calls": True}, {"minute": 660}, {"dialogues": 4}, {"dialogues": ""},
                                    {"stamp": "2026-11-12T10:30:00+09:00"}])
def test_rejects_corrupt_sheet_values(change):
    with pytest.raises(RuntimeError):
        team.parse_intervals([team.INTERVAL_HEADERS, hidden_row(**change)], "2026-11-13")


def test_duplicate_keys_and_wrong_headers_are_reported():
    with pytest.raises(RuntimeError, match="重複"):
        team.parse_intervals([team.INTERVAL_HEADERS, hidden_row(), hidden_row()], "2026-11-13")
    with pytest.raises(RuntimeError, match="見出し"):
        team.parse_intervals([["old header"]], "2026-11-13")


def test_numeric_sort_weighted_average_and_missing_participants():
    people = [member("1", "短い", {600: (10, 10, 0, 100)}), member("2", "長い", {600: (1, 1, 0, 1000)}),
              member("3", "ゼロ", {600: (0, 0, 0, 0)}), member("4", "未送信", {}), member("5", "未判定", {600: (2, None, None, None)})]
    report = ranking(people, 600)
    assert [row.name for row in report.rows][:3] == ["長い", "短い", "ゼロ"]
    assert report.people == 3 and report.calls == 11 and report.dialogues == 11
    assert report.average_dialogue == 100 and report.seconds == 1100
    assert "対象 5人" in summaries(report)[0] and "未取得/未判定 2人" in summaries(report)[0]
    assert ranking(people, 630).people == 0
    assert ranking([member("0", "ゼロ", {600: (0, 0, 0, 0)})], 600).average_dialogue is None


def test_total_duration_does_not_wrap_after_24_hours_and_copy_preserves_order():
    report = ranking([member("1", "A\tB\nC", {600: (1, 1, 0, 100000)})], 600)
    text = report_text(report, "2026-11-13", "全員", 600)
    assert "27:46:40" in text and "A B C\t1\t1\t27:46:40" in text
    assert "10:00～10:30" in text


def test_ui_filters_sorts_copies_and_graph_uses_same_totals(tmp_path):
    app = QApplication.instance() or QApplication([])
    # Windowsのoffscreenプラグインでは通常のフォント探索が行われない。
    # 既存のOSフォントを明示して、日本語を含む画像も確認する。
    font_path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / "meiryo.ttc"
    if font_path.exists():
        font_id = QFontDatabase.addApplicationFont(str(font_path))
        assert font_id >= 0
        app.setFont(QFont(QFontDatabase.applicationFontFamilies(font_id)[0], 10))
    people = [member("1", "一郎", {600: (4, 2, 1, 120), 630: (2, 1, 0, 30)}),
              member("2", "二郎", {600: (2, 1, 0, 300)}, "B")]
    data = team.TeamDay("sheet", "2026-11-13", [], [], people)
    panel = TeamStatisticsAnalysis()
    panel.resize(1100, 650)
    panel.set_data(data, "全員")
    panel.period.setCurrentIndex(panel.period.findData(600))
    assert panel.table.item(0, 0).text() == "二郎" and panel.report.seconds == 420
    panel.copy_button.click()
    copied = QApplication.clipboard().text()
    assert copied.index("二郎\t") < copied.index("一郎\t")
    assert panel.whole.text() in copied and panel.average.text() in copied
    assert panel.chart.points[20][2] == 420 and panel.chart.points[21][2] == 30
    panel.set_data(data, "A")
    assert panel.table.rowCount() == 1 and panel.report.seconds == 120
    assert panel.chart.points[20][1].people == 1
    panel.tabs.setCurrentIndex(1)
    panel.show()
    app.processEvents()
    assert panel.grab().save(str(tmp_path / "trend-preview.png"))
    assert panel.grab().save("logs/team_statistics_trend_preview.png")
    panel.tabs.setCurrentIndex(0)
    assert panel.grab().save(str(tmp_path / "ranking-preview.png"))
    assert panel.grab().save("logs/team_statistics_ranking_preview.png")
    panel.close()


def test_google_error_diagnostics_distinguish_scopes_service_and_invalid_login(caplog):
    for reason, expected in (("SERVICE_DISABLED", "APIが無効"), ("ACCESS_TOKEN_SCOPE_INSUFFICIENT", "許可が不足")):
        response = Response(403, {"error": {"status": "PERMISSION_DENIED", "message": "diagnostic", "details": [{"reason": reason}], "ignored": "secret"}})
        assert expected in str(team._sheet_read_error(response, "sheet", "A1:K"))
        assert reason in caplog.text and "secret" not in caplog.text
    assert "もう一度ログイン" in str(team._sheet_read_error(Response(401, {}), "sheet", "A1:K"))


def test_send_verification_does_not_claim_success_when_only_legacy_table_updated():
    from test_team_statistics import verify_dialog
    app = QApplication.instance() or QApplication([])
    dialog, data = verify_dialog({10: 3}, {10: 3})
    staff_id = dialog.controller.store.staff()[0]["staff_id"]
    dialog.controller.store.intervals = lambda *args: [{"start_minute": 600, "total_calls": 3, "dialogues": 2, "toss": 1, "dialogue_seconds": 90}]
    assert "30分データは未反映" in dialog.verify_own(data)
    data.interval_members = [member(staff_id, "岩田悠河", {600: (3, 2, 1, 90)})]
    assert "反映を確認" in dialog.verify_own(data)
    dialog.close()
