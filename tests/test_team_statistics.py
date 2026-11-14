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


class FakeFlow:
    def __init__(self):
        self.redirect_uri = None
        self.fetched = None
        payload = base64.urlsafe_b64encode(json.dumps({"email": "me@bb-connection.com"}).encode()).decode().rstrip("=")
        self.credentials = type("Creds", (), {"refresh_token": "r", "id_token": f"h.{payload}.s",
                                              "to_json": lambda self: json.dumps({"refresh_token": "r"})})()

    def authorization_url(self, **kwargs):
        assert kwargs.get("prompt") == "select_account"
        return "https://accounts.google.com/o/oauth2/auth?client_id=x", None

    def fetch_token(self, authorization_response):
        self.fetched = authorization_response


def fake_login(monkeypatch, tmp_path):
    from google_auth_oauthlib import flow as oauth_flow
    flow = FakeFlow()
    monkeypatch.setattr(oauth_flow.InstalledAppFlow, "from_client_config", classmethod(lambda cls, client, scopes: flow))
    return team.GoogleLogin(tmp_path / "token.json", client={"installed": {}}), flow


def visit_later(url, delay=0.2):
    """ブラウザの代わりに、少し遅れてリダイレクト先を開く（受付終了後の失敗は無視する）。"""
    import threading, time, urllib.request
    def go():
        time.sleep(delay)
        try:
            urllib.request.urlopen(url, timeout=5).read()
        except OSError:
            pass
    threading.Thread(target=go, daemon=True).start()


def test_login_receives_the_redirect_and_saves_the_token(monkeypatch, tmp_path):
    login, flow = fake_login(monkeypatch, tmp_path)
    urls = []
    def on_url(url):
        urls.append(url)
        visit_later(flow.redirect_uri + "favicon.ico", 0.1)  # 関係ないリクエストは無視する
        visit_later(flow.redirect_uri + "?state=s&code=abc", 0.4)
    email = login.login(on_url=on_url, open_browser=False, timeout_seconds=10)
    assert email == "me@bb-connection.com" and login.logged_in
    assert urls == ["https://accounts.google.com/o/oauth2/auth?client_id=x"]
    assert flow.fetched.startswith("https://localhost:") and "code=abc" in flow.fetched


def test_login_can_be_cancelled_while_waiting_for_the_browser(monkeypatch, tmp_path):
    import threading, time
    login, flow = fake_login(monkeypatch, tmp_path)
    cancel = threading.Event()
    threading.Timer(0.3, cancel.set).start()
    started = time.monotonic()
    with pytest.raises(team.LoginCancelled):
        login.login(cancel=cancel, open_browser=False, timeout_seconds=30)
    assert time.monotonic() - started < 3
    assert not login.logged_in


def test_declining_on_the_google_page_is_treated_as_cancel(monkeypatch, tmp_path):
    login, flow = fake_login(monkeypatch, tmp_path)
    with pytest.raises(team.LoginCancelled):
        login.login(on_url=lambda url: visit_later(flow.redirect_uri + "?error=access_denied"), open_browser=False, timeout_seconds=10)
    assert flow.fetched is None and not login.logged_in


class WaitingLogin(FakeLogin):
    def __init__(self):
        super().__init__(logged_in=False)

    def login(self, cancel=None, on_url=None):
        on_url("https://accounts.google.com/o/oauth2/auth?client_id=x")
        cancel.wait(5)
        raise team.LoginCancelled("ログインを中止しました。")


def test_waiting_screen_can_copy_the_url_and_cancel():
    import time
    app = QApplication.instance() or QApplication([])
    dialog = TeamStatisticsDialog(login=WaitingLogin())
    dialog.start_login()
    deadline = time.monotonic() + 3
    while not dialog.login_url and time.monotonic() < deadline:
        app.processEvents()
    assert dialog.login_button.text() == "ログイン待ち…" and not dialog.login_button.isEnabled()
    assert not dialog.cancel_login_button.isHidden() and dialog.copy_url_button.isEnabled()
    dialog.copy_login_url()
    assert QApplication.clipboard().text().startswith("https://accounts.google.com/")
    dialog.cancel_login()
    assert not dialog.logging_in and dialog.login_button.isEnabled()
    assert dialog.cancel_login_button.isHidden() and dialog.copy_url_button.isHidden()
    assert "中止しました" in dialog.login_error.text()
    deadline = time.monotonic() + 1
    while time.monotonic() < deadline:  # 止まった古いログインの結果は無視される
        app.processEvents()
    assert dialog.login_button.isEnabled() and "中止しました" in dialog.login_error.text()
    dialog.close()


from PySide6.QtCore import QObject, Signal


class FakeController(QObject):
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.busy = False
        self.status = ""
        self.collected = 0

    def collect(self):
        self.collected += 1
        self.busy = True
        self.changed.emit()

    def finish(self, message):
        self.busy = False
        self.status = message
        self.changed.emit()


def test_send_button_collects_own_data_and_rereads_the_sheet_after_it_is_written():
    app = QApplication.instance() or QApplication([])
    controller = FakeController()
    dialog = TeamStatisticsDialog(login=FakeLogin(), controller=controller)
    assert dialog.send_button.isEnabled() and dialog.send_button.text() == "自分の集計を送信"
    dialog.send_button.click()
    assert controller.collected == 1
    assert dialog.send_button.text() == "送信中…" and not dialog.send_button.isEnabled()
    assert "シートへ送っています" in dialog.status.text()
    controller.finish("16:31 集計済み：三浦 甲斐／本人の当日履歴 1件／シート送信済み")
    assert dialog.send_button.isEnabled()
    assert "シート送信済み" in dialog.status.text() and "自動で読み直します" in dialog.status.text()
    assert dialog.after_send_timer.isActive()
    dialog.close()


def test_send_failure_is_shown_without_scheduling_a_reread():
    app = QApplication.instance() or QApplication([])
    controller = FakeController()
    dialog = TeamStatisticsDialog(login=FakeLogin(logged_in=False), controller=controller)
    assert dialog.send_button.isEnabled()  # 送信はGoogleログインが無くてもできる
    dialog.send_own()
    controller.finish("集計できませんでした：CTIメインを1つだけ起動してください。 保存済みの値は維持しています。")
    assert "集計できませんでした" in dialog.status.text()
    assert not dialog.after_send_timer.isActive()
    dialog.close()


def test_send_button_is_hidden_without_a_controller():
    app = QApplication.instance() or QApplication([])
    dialog = TeamStatisticsDialog(login=FakeLogin())
    assert dialog.send_button.isHidden()
    dialog.close()
