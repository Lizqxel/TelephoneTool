"""全員分の通話集計（スプレッドシートの日付ごとの表）をTelephoneToolで表示する。

一度Googleにログインすれば、以後は開くたびと5分ごとに読み直す。読み取りだけで、シートは変更しない。
"""

import threading

from PySide6.QtCore import QDate, QObject, QTimer, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (QApplication, QComboBox, QDateEdit, QDialog, QHBoxLayout, QHeaderView, QLabel,
                               QPushButton, QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from services.team_statistics import LABELS, GoogleLogin, LoginCancelled, fetch_day, normalize

REFRESH_MS = 5 * 60 * 1000
AFTER_SEND_MS = 90 * 1000  # シートへの反映（1分ごと）を待ってから読み直す
ALL_UNITS = "全員"
WEEKDAYS = "月火水木金土日"
HEADER_ROWS = 2
# スプレッドシート「コール集計打ち込み用」の配色
BLACK = "#000000"
HEADER_TEXT = "#ffffff"
TITLE_TEXT = "#ff9900"
OWN_TEXT = "#ffd966"
CREAM = "#fff2cc"
WHITE = "#ffffff"
TOTAL = "#d9d9d9"


class _Relay(QObject):
    """バックグラウンドの処理結果を画面のスレッドへ戻す。"""
    login_done = Signal(int, str, str)  # 何回目のログインか、メール、エラー
    login_url = Signal(int, str)
    fetch_done = Signal(object, str, str)


class TeamStatisticsDialog(QDialog):
    def __init__(self, login=None, own_name="", parent=None, fetch=fetch_day, controller=None):
        super().__init__(parent)
        self.login = login or GoogleLogin()
        self.controller = controller  # 自分の集計（CallStatisticsController）。送信ボタンで使う
        self.waiting_send = False
        self.own_name = normalize(own_name)
        self.fetch_function = fetch
        self.data = None
        self.busy = False
        self.logging_in = False
        self.login_attempt = 0
        self.login_cancel = None
        self.login_url = ""
        self.relay = _Relay(self)
        self.relay.login_done.connect(self._login_done)
        self.relay.login_url.connect(self._login_url_ready)
        self.relay.fetch_done.connect(self._fetched)
        self.setWindowTitle("全員の通話集計（スプレッドシート）")
        self.resize(1200, 720)

        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        self.day = QDateEdit(QDate.currentDate())
        self.day.setCalendarPopup(True)
        self.day.setDisplayFormat("yyyy/MM/dd")
        self.unit = QComboBox()
        self.unit.addItem(ALL_UNITS)
        self.refresh_button = QPushButton("更新")
        self.send_button = QPushButton("自分の集計を送信")
        self.send_button.setToolTip("毎時00分の自動送信を待たずに、CTIから自分の当日分を集計してシートへ送ります。")
        self.account = QLabel()
        self.logout_button = QPushButton("ログアウト")
        top.addWidget(QLabel("日付"))
        top.addWidget(self.day)
        top.addWidget(QLabel("GPユニット"))
        top.addWidget(self.unit)
        top.addWidget(self.refresh_button)
        top.addSpacing(16)
        top.addWidget(self.send_button)
        top.addStretch()
        top.addWidget(self.account)
        top.addWidget(self.logout_button)
        layout.addLayout(top)
        self.status = QLabel()
        self.status.setWordWrap(True)
        self.status.setTextFormat(Qt.PlainText)
        layout.addWidget(self.status)

        # 未ログインの間は、表の代わりに大きなログイン案内を出す
        self.pages = QStackedWidget()
        self.login_page = QWidget()
        login_layout = QVBoxLayout(self.login_page)
        login_layout.addStretch(2)
        self.login_title = QLabel()
        self.login_title.setAlignment(Qt.AlignCenter)
        self.login_title.setWordWrap(True)
        self.login_title.setStyleSheet("font-size:22px; font-weight:bold;")
        self.login_detail = QLabel()
        self.login_detail.setAlignment(Qt.AlignCenter)
        self.login_detail.setWordWrap(True)
        self.login_detail.setTextFormat(Qt.PlainText)
        self.login_detail.setStyleSheet("font-size:14px; line-height:150%;")
        self.login_button = QPushButton("Googleでログイン")
        self.login_button.setCursor(Qt.PointingHandCursor)
        self.login_button.setMinimumSize(320, 56)
        self.login_button.setStyleSheet(
            "QPushButton { font-size:18px; font-weight:bold; color:#ffffff; background:#1a73e8; border:none; border-radius:8px; padding:12px 40px; }"
            "QPushButton:hover { background:#1765cc; }"
            "QPushButton:disabled { background:#9aa0a6; color:#f1f3f4; }")
        # ログイン待ちの間だけ出す：違うアカウントのブラウザで開いた場合のやり直し用
        self.cancel_login_button = QPushButton("キャンセル")
        self.copy_url_button = QPushButton("ログイン用のURLをコピー")
        for button in (self.cancel_login_button, self.copy_url_button):
            button.setMinimumHeight(36)
            button.setStyleSheet("QPushButton { font-size:14px; padding:6px 18px; }")
        waiting_buttons = QHBoxLayout()
        waiting_buttons.addStretch()
        waiting_buttons.addWidget(self.cancel_login_button)
        waiting_buttons.addSpacing(12)
        waiting_buttons.addWidget(self.copy_url_button)
        waiting_buttons.addStretch()
        self.login_error = QLabel()
        self.login_error.setAlignment(Qt.AlignCenter)
        self.login_error.setWordWrap(True)
        self.login_error.setTextFormat(Qt.PlainText)
        login_layout.addWidget(self.login_title)
        login_layout.addSpacing(12)
        login_layout.addWidget(self.login_detail)
        login_layout.addSpacing(20)
        login_layout.addWidget(self.login_button, 0, Qt.AlignHCenter)
        login_layout.addSpacing(10)
        login_layout.addLayout(waiting_buttons)
        login_layout.addSpacing(12)
        login_layout.addWidget(self.login_error)
        login_layout.addStretch(3)

        self.table_page = QWidget()
        table_layout = QVBoxLayout(self.table_page)
        table_layout.setContentsMargins(0, 0, 0, 0)
        self.banner = QLabel()
        self.banner.setStyleSheet(f"background:{BLACK}; color:{TITLE_TEXT}; font-weight:bold; font-size:15px; padding:4px 8px;")
        self.banner.hide()
        table_layout.addWidget(self.banner)
        self.table = QTableWidget(0, 0)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.horizontalHeader().hide()
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.setStyleSheet("QTableWidget { gridline-color: #8c8c8c; }")
        table_layout.addWidget(self.table, 1)
        self.pages.addWidget(self.login_page)
        self.pages.addWidget(self.table_page)
        layout.addWidget(self.pages, 1)
        note = QLabel("スプレッドシート「コール集計打ち込み用」の内容を表示しています（読み取りのみ）。"
                      "各自のTelephoneToolが毎時送った値が、1〜2分ほどで反映されます。")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.timer = QTimer(self)
        self.timer.setInterval(REFRESH_MS)
        self.timer.timeout.connect(self.fetch)
        self.after_send_timer = QTimer(self)
        self.after_send_timer.setSingleShot(True)
        self.after_send_timer.setInterval(AFTER_SEND_MS)
        self.after_send_timer.timeout.connect(self.fetch)
        self.send_button.clicked.connect(self.send_own)
        if controller is not None:
            controller.changed.connect(self._controller_changed)
        self._controller_changed()
        self.refresh_button.clicked.connect(self.fetch)
        self.login_button.clicked.connect(self.start_login)
        self.cancel_login_button.clicked.connect(lambda: self.cancel_login())
        self.copy_url_button.clicked.connect(self.copy_login_url)
        self.logout_button.clicked.connect(self.logout)
        self.day.dateChanged.connect(self.fetch)
        self.unit.currentIndexChanged.connect(self.render)
        self.update_login_state()

    # ----- ログイン -----
    def update_login_state(self):
        logged_in = bool(self.login.client) and self.login.logged_in and not self.logging_in
        self.pages.setCurrentWidget(self.table_page if logged_in else self.login_page)
        for widget in (self.day, self.unit):
            widget.setEnabled(logged_in)
        self.refresh_button.setEnabled(logged_in and not self.busy)
        self.logout_button.setVisible(logged_in)
        self.account.setText(f"ログイン中：{self.login.email or 'Googleアカウント'}" if logged_in else "")
        self.cancel_login_button.setVisible(self.logging_in)
        self.copy_url_button.setVisible(self.logging_in)
        self.copy_url_button.setEnabled(bool(self.login_url))
        if logged_in:
            return
        if not self.login.client:
            self.login_title.setText("Googleログインの設定ファイルがありません")
            self.login_detail.setText(
                "配布版のTelephoneTool（exe）には入っているので、そのまま使えます。\n"
                "ソースから動かしている場合は、管理者から google_oauth_client.json を受け取り、\n"
                "main.py と同じフォルダに置いてから、TelephoneToolを起動し直してください。")
            self.login_button.setText("Googleでログイン")
            self.login_button.setEnabled(False)
        elif self.logging_in:
            self.login_title.setText("ブラウザでログインしてください")
            self.login_detail.setText(
                "開いたブラウザで、シートを開ける社内のGoogleアカウントを選んでログインし、「許可」を押してください。\n"
                "「ログインが完了しました」と出たら、この画面に戻ってください（5分以内）。\n\n"
                "違うアカウントのブラウザで開いてしまった場合は、「キャンセル」してやり直すか、\n"
                "「ログイン用のURLをコピー」を押して、社内アカウントでログインしているブラウザのアドレス欄に貼り付けてください。")
            self.login_button.setText("ログイン待ち…")
            self.login_button.setEnabled(False)
        else:
            self.login_title.setText("全員分の集計を見るには、Googleにログインしてください")
            self.login_detail.setText(
                "ログインは最初の1回だけです。\n\n"
                "① 下のボタンを押すと、ブラウザが開きます\n"
                "② シートを開ける社内のGoogleアカウントを選んでログインします\n"
                "③ 「許可」を押します\n"
                "④ 「ログインが完了しました」と出たら、この画面に戻ります")
            self.login_button.setText("Googleでログイン")
            self.login_button.setEnabled(True)

    def show_login_message(self, text, error=True):
        color = "#d93025" if error else "#5f6368"
        self.login_error.setStyleSheet(f"font-size:14px; color:{color}; font-weight:bold;")
        self.login_error.setText(text)

    def start_login(self):
        if self.logging_in or self.login.logged_in or not self.login.client:
            return
        self.login_attempt += 1
        attempt = self.login_attempt
        cancel = threading.Event()
        self.login_cancel = cancel
        self.login_url = ""
        self.logging_in = True
        self.show_login_message("")
        self.status.setText("")
        self.update_login_state()

        def work():
            try:
                email = self.login.login(cancel=cancel, on_url=lambda url: self.relay.login_url.emit(attempt, url))
                self.relay.login_done.emit(attempt, email, "")
            except LoginCancelled as exc:
                self.relay.login_done.emit(attempt, "", f"CANCELLED:{exc}")
            except Exception as exc:
                self.relay.login_done.emit(attempt, "", str(exc) or exc.__class__.__name__)

        threading.Thread(target=work, daemon=True).start()

    def cancel_login(self, silent=False):
        """ログイン待ちをやめる。待っている処理は0.5秒以内に止まり、その結果は無視する。"""
        if not self.logging_in:
            return
        if self.login_cancel is not None:
            self.login_cancel.set()
        self.login_attempt += 1
        self.logging_in = False
        self.login_url = ""
        self.update_login_state()
        if not silent:
            self.show_login_message("ログインを中止しました。もう一度「Googleでログイン」を押すと、アカウントを選び直せます。", error=False)

    def _login_url_ready(self, attempt, url):
        if attempt == self.login_attempt:
            self.login_url = url
            self.update_login_state()

    def copy_login_url(self):
        if self.login_url:
            QApplication.clipboard().setText(self.login_url)
            self.show_login_message("URLをコピーしました。社内アカウントでログインしているブラウザのアドレス欄に貼り付けてください。", error=False)

    def logout(self):
        self.login.logout()
        self.data = None
        self.render()
        self.status.setText("ログアウトしました。")
        self.update_login_state()

    def _login_done(self, attempt, email, error):
        if attempt != self.login_attempt:
            return  # キャンセル済みの古いログインの結果
        self.logging_in = False
        self.login_url = ""
        self.update_login_state()
        if error.startswith("CANCELLED:"):
            self.show_login_message(error.split(":", 1)[1] + "もう一度「Googleでログイン」を押すと、アカウントを選び直せます。", error=False)
            return
        if error:
            self.show_login_message(f"ログインできませんでした：{error}\nもう一度ボタンを押してやり直してください。")
            return
        self.status.setText(f"{email or 'Googleアカウント'} でログインしました。")
        self.fetch()

    # ----- 自分の集計の送信 -----
    def send_own(self):
        if self.controller is None or self.controller.busy:
            return
        self.waiting_send = True
        self.status.setText("自分の集計をCTIから取得して、シートへ送っています…")
        self.controller.collect()
        self._controller_changed()

    def _controller_changed(self):
        controller = self.controller
        busy = controller is not None and controller.busy
        self.send_button.setVisible(controller is not None)
        self.send_button.setEnabled(controller is not None and not busy)
        self.send_button.setText("送信中…" if busy else "自分の集計を送信")
        if not self.waiting_send or busy:
            return
        self.waiting_send = False
        message = controller.status
        if "シート送信済み" in message:
            self.after_send_timer.start()
            message += "　→ 1〜2分後に自動で読み直します。"
        self.status.setText(message)

    # ----- 読み取り -----
    def fetch(self):
        if self.busy or not self.login.logged_in:
            self.update_login_state()  # 未ログインなら大きなログイン案内を表示する
            return
        self.busy = True
        day = self.day.date().toString("yyyy-MM-dd")
        self.status.setText("シートを読み込んでいます…")
        self.update_login_state()

        def work():
            try:
                self.relay.fetch_done.emit(self.fetch_function(self.login.session(), day), "", day)
            except Exception as exc:
                self.relay.fetch_done.emit(None, str(exc) or exc.__class__.__name__, day)

        threading.Thread(target=work, daemon=True).start()

    def _fetched(self, data, error, day):
        self.busy = False
        self.update_login_state()
        if day != self.day.date().toString("yyyy-MM-dd"):
            self.fetch()  # 読み込み中に日付が変わった
            return
        if error and "もう一度ログイン" in error:
            # ログインの期限切れ・取り消し。保存したログイン情報を消して、ログイン案内に戻す。
            self.login.logout()
            self.update_login_state()
            self.show_login_message(error)
            return
        if error:
            self.status.setText(f"シートを読み取れませんでした：{error}" + ("（前回の表示を残しています）" if self.data else ""))
            return
        self.data = data
        units = sorted({m.unit for m in data.members if m.unit})
        current = self.unit.currentText()
        self.unit.blockSignals(True)
        self.unit.clear()
        self.unit.addItems([ALL_UNITS, *units])
        self.unit.setCurrentIndex(max(0, self.unit.findText(current)))
        self.unit.blockSignals(False)
        from datetime import datetime
        self.status.setText(f"{datetime.now():%H:%M} 更新：{data.sheet}　{len(data.members)}人")
        self.render()

    # ----- 表示（スプレッドシートの見た目に合わせる） -----
    def render(self):
        self.table.clear()
        self.table.clearSpans()
        if not self.data:
            self.banner.hide()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return
        date = self.day.date()
        self.banner.setText(f"【{date.month()}/{date.day()}（{WEEKDAYS[date.dayOfWeek() - 1]}）】")
        self.banner.show()
        unit = self.unit.currentText()
        members = [m for m in self.data.members if unit in (ALL_UNITS, "") or m.unit == unit]
        groups = [("日計", None), *[(f"{h}時台", h) for h in self.data.hours]]
        self.table.setColumnCount(3 + 3 * len(groups))
        self.table.setRowCount(HEADER_ROWS + len(members) + 1)

        def values(member, hour):
            return member.total if hour is None else member.hours.get(hour, (None, None, None))

        def put(row, column, text, background, foreground="#000000", bold=False, align=Qt.AlignCenter):
            item = QTableWidgetItem(text)
            item.setTextAlignment(align | Qt.AlignVCenter)
            item.setBackground(QBrush(QColor(background)))
            item.setForeground(QBrush(QColor(foreground)))
            if bold:
                font = QFont(self.table.font())
                font.setBold(True)
                item.setFont(font)
            self.table.setItem(row, column, item)

        def group_color(index):
            # 日計は薄い黄色、時台は白と薄い黄色を交互に（シートと同じ並び）
            return CREAM if index == 0 or index % 2 == 0 else WHITE

        # 見出し2行：1行目「名前／日計／10時台…」、2行目「部署／GPユニット／コール・対話・トス」
        put(0, 0, "名前", BLACK, HEADER_TEXT, True)
        self.table.setSpan(0, 0, 2, 1)
        put(0, 1, "", BLACK)
        put(0, 2, "", BLACK)
        put(1, 1, "部署", BLACK, HEADER_TEXT, True)
        put(1, 2, "GPユニット", BLACK, HEADER_TEXT, True)
        for index, (title, _) in enumerate(groups):
            start = 3 + 3 * index
            put(0, start, title, BLACK, HEADER_TEXT, True)
            self.table.setSpan(0, start, 1, 3)
            for k, label in enumerate(LABELS):
                put(1, start + k, label, BLACK, HEADER_TEXT, True)

        right = Qt.AlignRight
        for offset, member in enumerate(members):
            row = HEADER_ROWS + offset
            own = bool(self.own_name) and normalize(member.name) == self.own_name
            put(row, 0, member.name, BLACK, OWN_TEXT if own else HEADER_TEXT, True, Qt.AlignLeft)
            put(row, 1, member.department, CREAM, bold=True, align=right)
            put(row, 2, member.unit, CREAM, bold=True, align=right)
            for index, (_, hour) in enumerate(groups):
                for k, value in enumerate(values(member, hour)):
                    put(row, 3 + 3 * index + k, "" if value is None else str(value),
                        group_color(index), bold=hour is None or own, align=right)

        # 最終行は表示中の人の合計
        row = HEADER_ROWS + len(members)
        put(row, 0, f"全体（{len(members)}人）", BLACK, HEADER_TEXT, True, Qt.AlignLeft)
        put(row, 1, "", TOTAL)
        put(row, 2, "", TOTAL)
        for index, (_, hour) in enumerate(groups):
            for k in range(3):
                known = [values(m, hour)[k] for m in members if values(m, hour)[k] is not None]
                put(row, 3 + 3 * index + k, str(sum(known)) if known else "", TOTAL, bold=True, align=right)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Fixed)
        header.setDefaultSectionSize(52)
        self.table.setColumnWidth(0, 96)
        self.table.setColumnWidth(1, 84)
        self.table.setColumnWidth(2, 96)

    def showEvent(self, event):
        super().showEvent(event)
        self.update_login_state()
        self.timer.start()
        self.fetch()

    def hideEvent(self, event):
        self.timer.stop()
        self.cancel_login(silent=True)  # 画面を閉じたらログイン待ちもやめる
        super().hideEvent(event)
