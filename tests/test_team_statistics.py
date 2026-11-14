import base64
import json
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import QDate
from PySide6.QtWidgets import QApplication

from services import team_statistics as team
from ui.team_statistics_dialog import TeamStatisticsDialog


def block_rows(day, members, hours=(10, 11)):
    """シートの1日分（見出し・時台・ラベル・名前の行）をAPIの返す形で作る。"""
    hour_row = ["", "名前", "", "", "日計", "", ""]
    label_row = ["", "", "部署", "GPユニット", "コール", "対話", "トス"]
    for hour in hours:
        hour_row += [f"{hour}時台", "", ""]
        label_row += ["コール", "対話", "トス"]
    rows = [["", f"【11/{day}（日）】"], hour_row, label_row]
    for name, unit, values in members:
        row = ["", name, "岩田祐樹", unit, *[str(v) for v in values[:3]]]
        for k in range(len(hours)):
            chunk = values[3 + 3 * k: 6 + 3 * k]
            row += ["" if v is None else str(v) for v in chunk]
        rows.append(row)
    rows.append([])
    return rows


def sheet_grid():
    grid = [["", "コール集計打ち込み（2026年11月）"], ["", "説明"], []]
    grid += block_rows(1, [("岩田祐樹", "岩田祐樹", [0] * 9)])
    grid += block_rows(10, [("岩田悠河", "岩田祐樹", [9, 9, 9, 9, 9, 9, 9, 9, 9])])
    grid += block_rows(13, [("岩田悠河", "岩田祐樹", [20, 6, 1, 12, 4, 1, 8, 2, 0]),
                            ("三浦甲斐", "杉崎伸明", [5, 1, 0, 5, 1, 0, None, None, None])])
    return grid


class Response:
    def __init__(self, status, body):
        self.status_code, self.body, self.text = status, body, json.dumps(body)

    def json(self):
        return self.body


class Session:
    def __init__(self, grid, status=200):
        self.grid, self.status, self.ranges = grid, status, []

    def get(self, url, params, timeout):
        self.ranges.append(params["ranges"])
        if self.status != 200:
            return Response(self.status, {"error": {"message": "Unable to parse range: x"}})
        name, cells = params["ranges"].rsplit("!", 1)
        if cells == "A1:D":
            values = [row[:4] for row in self.grid]
        else:
            start, end = (int(x) for x in cells.split(":"))
            values = self.grid[start - 1:end]
        while values and not values[-1]:
            values.pop()
        return Response(200, {"valueRanges": [{"range": name, "values": values}]})


def test_reads_the_requested_day_without_mixing_up_11_1_and_11_10():
    session = Session(sheet_grid())
    data = team.fetch_day(session, "2026-11-13")
    assert data.sheet == "コール集計打ち込み用_11月"
    assert data.hours == [10, 11]
    assert [m.name for m in data.members] == ["岩田悠河", "三浦甲斐"]
    assert data.members[0].total == (20, 6, 1) and data.members[0].hours[11] == (8, 2, 0)
    assert data.members[1].unit == "杉崎伸明" and data.members[1].hours[11] == (None, None, None)
    assert session.ranges[0] == "'コール集計打ち込み用_11月'!A1:D"
    assert data_for(session, "2026-11-01").members[0].name == "岩田祐樹"
    assert data_for(session, "2026-11-10").members[0].total == (9, 9, 9)


def data_for(session, day):
    return team.fetch_day(session, day)


def test_sheet_errors_are_reported_in_plain_words():
    with pytest.raises(RuntimeError, match="11/14の表"):
        team.fetch_day(Session(sheet_grid()), "2026-11-14")
    with pytest.raises(RuntimeError, match="2025年11月の表ではありません"):
        team.fetch_day(Session(sheet_grid()), "2025-11-13")
    with pytest.raises(RuntimeError, match="読めません"):
        team.fetch_day(Session(sheet_grid(), 403), "2026-11-13")
    with pytest.raises(RuntimeError, match="がありません"):
        team.fetch_day(Session(sheet_grid(), 400), "2026-12-01")


def test_login_state_is_kept_per_pc_and_logout_removes_it(tmp_path):
    login = team.GoogleLogin(tmp_path / "token.json", client={"installed": {}})
    assert not login.logged_in
    with pytest.raises(RuntimeError, match="ログインしていません"):
        login.credentials()
    (tmp_path / "token.json").write_text(json.dumps({"refresh_token": "r", "email": "a@example.com"}), encoding="utf-8")
    assert login.logged_in and login.email == "a@example.com"
    login.logout()
    assert not login.logged_in


def test_oauth_client_can_come_from_settings_and_email_is_read_for_display(monkeypatch, tmp_path):
    monkeypatch.setattr(team, "_search_dirs", lambda: [tmp_path])
    assert team.client_config([("s", {})]) is None
    config = team.client_config([("s", {"callStatisticsSheet": {"oauthClient": {"client_id": "id", "client_secret": "sec"}}})])
    assert config["installed"]["client_id"] == "id"
    (tmp_path / team.CLIENT_FILE).write_text(json.dumps({"installed": {"client_id": "file"}}), encoding="utf-8")
    assert team.client_config([("s", {})])["installed"]["client_id"] == "file"
    payload = base64.urlsafe_b64encode(json.dumps({"email": "me@bb-connection.com"}).encode()).decode().rstrip("=")
    assert team.email_from_id_token(f"h.{payload}.s") == "me@bb-connection.com"
    assert team.email_from_id_token("broken") == ""


class FakeLogin:
    client = {"installed": {}}
    email = "me@bb-connection.com"

    def __init__(self, logged_in=True):
        self.logged_in = logged_in

    def session(self):
        return Session(sheet_grid())


def test_dialog_looks_like_the_sheet_with_a_total_row_and_own_name_marked():
    app = QApplication.instance() or QApplication([])
    dialog = TeamStatisticsDialog(login=FakeLogin(), own_name="岩田 悠河")
    dialog.day.setDate(QDate(2026, 11, 13))
    data = team.fetch_day(Session(sheet_grid()), "2026-11-13")
    dialog._fetched(data, "", "2026-11-13")
    table = dialog.table
    assert dialog.banner.text() == "【11/13（金）】"
    assert table.rowCount() == 2 + 2 + 1  # 見出し2行＋2人＋全体
    # 見出しはシートと同じ「名前／日計／10時台」と「部署／GPユニット／コール・対話・トス」
    assert [table.item(0, c).text() for c in (0, 3, 6, 9)] == ["名前", "日計", "10時台", "11時台"]
    assert [table.item(1, c).text() for c in (1, 2, 6, 7, 8)] == ["部署", "GPユニット", "コール", "対話", "トス"]
    assert table.rowSpan(0, 0) == 2 and table.columnSpan(0, 3) == 3 and table.columnSpan(0, 6) == 3
    assert table.item(2, 0).text() == "岩田悠河" and table.item(2, 3).text() == "20"
    assert table.item(2, 0).background().color().name() == "#000000"
    assert table.item(2, 0).foreground().color().name() == "#ffd966"  # 自分の名前
    assert table.item(3, 0).foreground().color().name() == "#ffffff"
    assert table.item(2, 1).background().color().name() == "#fff2cc"  # 部署は薄い黄色
    assert table.item(2, 6).background().color().name() == "#ffffff"  # 10時台は白
    assert table.item(2, 9).background().color().name() == "#fff2cc"  # 11時台は薄い黄色
    assert table.item(3, 9).text() == ""  # 未入力は空欄
    assert table.item(4, 0).text() == "全体（2人）" and table.item(4, 3).text() == "25"
    dialog.unit.setCurrentText("杉崎伸明")
    assert table.rowCount() == 4 and table.item(2, 0).text() == "三浦甲斐"
    assert "ログイン中" in dialog.account.text()
    dialog.close()


def test_dialog_keeps_previous_table_when_a_refresh_fails():
    app = QApplication.instance() or QApplication([])
    dialog = TeamStatisticsDialog(login=FakeLogin(), own_name="")
    dialog.day.setDate(QDate(2026, 11, 13))
    dialog._fetched(team.fetch_day(Session(sheet_grid()), "2026-11-13"), "", "2026-11-13")
    dialog._fetched(None, "シートへ接続できませんでした", "2026-11-13")
    assert dialog.table.rowCount() == 5
    assert "前回の表示を残しています" in dialog.status.text()
    dialog.close()


def test_dialog_shows_a_large_login_screen_until_logged_in():
    app = QApplication.instance() or QApplication([])
    dialog = TeamStatisticsDialog(login=FakeLogin(logged_in=False))
    dialog.fetch()
    assert dialog.pages.currentWidget() is dialog.login_page
    assert "Googleにログインしてください" in dialog.login_title.text()
    assert dialog.login_button.text() == "Googleでログイン" and dialog.login_button.isEnabled()
    assert dialog.logout_button.isHidden() and not dialog.refresh_button.isEnabled()
    dialog.close()


def test_missing_oauth_file_explains_where_to_put_it():
    app = QApplication.instance() or QApplication([])
    missing = FakeLogin(logged_in=False)
    missing.client = None
    dialog = TeamStatisticsDialog(login=missing)
    assert dialog.pages.currentWidget() is dialog.login_page
    assert "設定ファイルがありません" in dialog.login_title.text()
    assert "google_oauth_client.json" in dialog.login_detail.text()
    assert not dialog.login_button.isEnabled()
    dialog.close()


def test_logged_in_dialog_shows_the_table_and_expired_login_returns_to_the_login_screen():
    app = QApplication.instance() or QApplication([])
    login = FakeLogin()
    login.logout = lambda: setattr(login, "logged_in", False)
    dialog = TeamStatisticsDialog(login=login)
    assert dialog.pages.currentWidget() is dialog.table_page
    assert not dialog.logout_button.isHidden()
    dialog._fetched(None, "Googleのログインの有効期限が切れました。もう一度ログインしてください。", dialog.day.date().toString("yyyy-MM-dd"))
    assert dialog.pages.currentWidget() is dialog.login_page
    assert "有効期限" in dialog.login_error.text()
    dialog.close()
