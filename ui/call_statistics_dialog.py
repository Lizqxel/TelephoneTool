"""30分ごとの集計・シート送信と、日別・時間帯別の通話集計表。

起動時・毎時00分/30分・手動で、架電履歴と見込検索を裏で検索して集計し、スプレッドシートへ送る。
"""

from datetime import datetime
import os
from pathlib import Path
import threading

from PySide6.QtCore import QDate, QObject, QThread, QTimer, Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDateEdit, QDialog, QHBoxLayout,
                               QHeaderView, QLabel, QPushButton, QTableWidget,
                               QTableWidgetItem, QVBoxLayout)

from services.call_statistics import JST, StatisticsStore, format_duration, next_half_hour
from services.call_statistics_sheet import SheetSender, sheet_hours, sheet_intervals
from services.cti_statistics_service import CtiStatisticsCollector


class StatisticsWorker(QThread):
    completed = Signal(object, str, str)  # 集計結果、エラー、クリア・送信のお知らせ

    def __init__(self, store, numeric_unit, sender_factory=None, parent=None):
        super().__init__(parent)
        self.store = store
        self.numeric_unit = numeric_unit
        self.sender_factory = sender_factory or SheetSender.from_settings
        self.cancelled = threading.Event()

    def run(self):
        try:
            collector = CtiStatisticsCollector(self.cancelled)
            snapshot = collector.collect(self.numeric_unit)
            if self.cancelled.is_set():
                return
            self.store.save(snapshot)
        except Exception as exc:
            self.completed.emit(None, str(exc), "")
            return
        notices = []
        if snapshot.history_counters is not None:
            # 保存に成功してから架電履歴の一覧をクリアする。失敗しても集計は保存済みなので警告だけにする。
            try:
                collector.clear_history()
            except Exception as exc:
                notices.append(f"架電履歴のクリア未完了：{exc}")
        if not self.cancelled.is_set():
            notices.append(self.send(snapshot))
        self.completed.emit(snapshot, "", "／".join(n for n in notices if n))

    def send(self, snapshot):
        """保存済みの当日分をシートへ送る。失敗しても集計結果は保存済みで、次回の送信で回復する。"""
        try:
            sender = self.sender_factory()
            if sender is None:
                return "シートの送信先が未設定のため送信していません。"
            day = snapshot.acquired_at.astimezone(JST).date().isoformat()
            hours = sheet_hours(self.store.display_hours(snapshot.staff_id, day))
            intervals = self.store.intervals(snapshot.staff_id, day)
            if not hours and not intervals:
                return ""
            # 履歴を取れなかった回は、前回保存した30分データの取得日時も維持する。
            details = {"intervals": sheet_intervals(intervals), "observed_at": intervals[-1]["observed_at"]} if intervals else {}
            result = sender.send(day, snapshot.staff_id, snapshot.staff_name, hours, **details)
            skipped = result.get("skipped") or []
            if skipped:
                return "シート送信済み（シートに枠が無い時間帯：" + "・".join(f"{h}時" for h in skipped) + "）"
            # フォームは書き込み結果を返さないので、反映までの時間と確かめ方も添える
            return "シート送信済み（反映は1〜2分後。全員の集計で確認できます）"
        except Exception as exc:
            return f"シート送信失敗：{exc}"


class CallStatisticsController(QObject):
    changed = Signal()
    status_changed = Signal(str)

    def __init__(self, parent=None, store=None, autostart=True, sender_factory=None):
        super().__init__(parent)
        if store is None:
            base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "TelephoneTool"
            store = StatisticsStore(base / "call_statistics.sqlite3")
        self.store = store
        self.sender_factory = sender_factory
        self.worker = None
        self.dialog = None
        self.closing = False
        self.pending = False
        self.status = "初回の集計前です。"
        self.automatic = self.store.preference("automatic", "1") == "1"
        self.timer = QTimer(self)
        # 既定の粗いタイマーは最大5%早く発火し、00分の境界より前に読むと時間帯がずれるため。
        self.timer.setTimerType(Qt.PreciseTimer)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._automatic_tick)
        self.schedule()
        if autostart:
            QTimer.singleShot(2000, self._startup)

    @property
    def busy(self):
        return self.worker is not None and self.worker.isRunning()

    def _startup(self):
        if self.automatic and not self.closing:
            self.collect()

    def schedule(self):
        self.timer.stop()
        now = datetime.now(JST)
        self.next_run = next_half_hour(now)
        if self.automatic and not self.closing:
            # 境界の2秒後に読み、前の時間帯の差分を確定させる（保存側で00分に切り揃える）。
            self.timer.start(max(1, int((self.next_run - now).total_seconds() * 1000)) + 2000)

    def _automatic_tick(self):
        self.collect()
        self.schedule()

    def set_automatic(self, enabled):
        self.automatic = enabled
        self.store.set_preference("automatic", "1" if enabled else "0")
        self.schedule()
        self.set_status("30分ごとの自動集計を有効にしました。" if enabled else "自動集計を停止しました。「今すぐ集計」から手動で取得できます。")

    def set_status(self, message):
        self.status = message
        self.status_changed.emit(message)
        self.changed.emit()

    def collect(self):
        if self.closing:
            return
        if self.busy:
            # 実行中に集計が重なった場合は、終わってから実行する。
            self.pending = True
            return
        worker = StatisticsWorker(self.store, self.store.preference("duration_unit", "seconds"), self.sender_factory, self)
        worker.completed.connect(self._completed)
        worker.finished.connect(self._finished)
        # 「集計中」になってから知らせる（先に知らせると、受け取った画面が集計中でないと誤解する）
        self.worker = worker
        worker.start()
        self.set_status("CTIを集計しています…")

    def _completed(self, snapshot, error, notice=""):
        if error:
            self.set_status(f"集計できませんでした：{error} 保存済みの値は維持しています。")
            return
        # クリアや送信だけ失敗した場合も、集計は保存済みであることを示す。
        saved = "（保存済み）" if any(word in notice for word in ("未完了", "失敗")) else ""
        parts = [f"{snapshot.acquired_at:%H:%M} 集計済み{saved}：{snapshot.staff_name}"]
        if snapshot.history_counters is not None:
            parts.append(f"本人の当日履歴 {snapshot.history_rows}件")
        parts += [text for text in (snapshot.note, notice) if text]
        self.set_status("／".join(parts))

    def _finished(self):
        worker = self.worker
        self.worker = None
        if worker:
            worker.deleteLater()
        self.changed.emit()
        if self.pending and not self.closing:
            self.pending = False
            self.collect()

    def show(self, collect=False):
        if self.dialog is None:
            self.dialog = CallStatisticsDialog(self, self.parent())
        self.dialog.refresh()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()
        if collect:
            self.collect()

    def stop(self):
        self.closing = True
        self.timer.stop()
        if self.busy:
            self.worker.cancelled.set()
            return False
        return True


class CallStatisticsDialog(QDialog):
    HEADERS = ["時間帯", "コール", "対話", "トス", "応答時間", "新規見込", "対話NG", "取得状態"]

    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self._scroll_key = None
        self.setWindowTitle("時間帯別 通話集計")
        self.resize(980, 690)
        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        self.day = QDateEdit(QDate.currentDate())
        self._auto_day = QDate.currentDate()  # 利用者が変えていなければ、日付が変わったら今日へ進める
        self.day.setCalendarPopup(True)
        self.day.setDisplayFormat("yyyy/MM/dd")
        self.staff = QComboBox()
        self.collect_button = QPushButton("今すぐ集計")
        self.collect_button.setToolTip("毎時00分・30分を待たずに、CTIから当日分を取得してシートへ送ります。日付欄は保存済みの表の表示日です。")
        self.automatic = QCheckBox("30分ごとに自動集計")
        self.automatic.setToolTip("起動時と毎時00分・30分に集計してシートへ送ります。")
        self.automatic.setChecked(controller.automatic)
        top.addWidget(QLabel("日付"))
        top.addWidget(self.day)
        top.addWidget(self.staff, 1)
        top.addWidget(self.automatic)
        top.addWidget(self.collect_button)
        layout.addLayout(top)
        options = QHBoxLayout()
        options.addWidget(QLabel("通話時間が数値だけの場合の単位"))
        self.duration_unit = QComboBox()
        self.duration_unit.addItem("秒", "seconds")
        self.duration_unit.addItem("分", "minutes")
        self.duration_unit.setCurrentIndex(max(0, self.duration_unit.findData(controller.store.preference("duration_unit", "seconds"))))
        options.addWidget(self.duration_unit)
        options.addStretch()
        self.next_label = QLabel()
        options.addWidget(self.next_label)
        layout.addLayout(options)
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        self.status_label.setTextFormat(Qt.PlainText)
        layout.addWidget(self.status_label)
        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(len(self.HEADERS) - 1, QHeaderView.ResizeToContents)
        self.table.setToolTip("コール・対話・トスは架電履歴から、集計開始前の分も含めて当日分を集計します。")
        layout.addWidget(self.table, 1)
        self.total_label = QLabel()
        self.total_label.setTextFormat(Qt.PlainText)
        self.total_label.setWordWrap(True)
        layout.addWidget(self.total_label)
        note = QLabel("コール＝架電履歴の件数（通話0秒・留守も含む）。対話＝見込＋対話有NG（NGは現アナNG・多重留守NG以外を対話有）。\n"
                      "トス＝当日の見込のうち、見込検索のランクが「トス」のもの。応答時間＝IP電話通話時間。対話NG＝対話有NG。")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.collect_button.clicked.connect(lambda: controller.collect())
        self.automatic.toggled.connect(controller.set_automatic)
        self.duration_unit.currentIndexChanged.connect(self._unit_changed)
        self.day.dateChanged.connect(self.refresh)
        self.staff.currentIndexChanged.connect(self.refresh)
        controller.changed.connect(self.refresh)
        self.refresh()

    def _unit_changed(self):
        self.controller.store.set_preference("duration_unit", self.duration_unit.currentData())
        self.controller.set_status("通話時間の単位を変更しました。次の集計で当日の応答時間を再計算します。")

    def refresh(self):
        today = QDate.currentDate()
        if self._auto_day != today and self.day.date() == self._auto_day:
            self._auto_day = today
            self.day.setDate(today)  # dateChanged から refresh が呼ばれる
            return
        controller = self.controller
        people = controller.store.staff()
        selected = self.staff.currentData()
        self.staff.blockSignals(True)
        self.staff.clear()
        for person in people:
            self.staff.addItem(person["staff_name"], person["staff_id"])
        if selected:
            self.staff.setCurrentIndex(max(0, self.staff.findData(selected)))
        self.staff.blockSignals(False)
        self.collect_button.setEnabled(not controller.busy and not controller.closing)
        self.collect_button.setText("集計中…" if controller.busy else "今すぐ集計")
        self.duration_unit.setEnabled(not controller.busy)
        self.status_label.setText(controller.status)
        self.next_label.setText(f"次回の集計・送信 {controller.next_run:%H:%M}" if controller.automatic else "自動集計 停止中")
        day = self.day.date().toString("yyyy-MM-dd")
        staff_id = self.staff.currentData() or ""
        rows = controller.store.display_hours(staff_id, day)
        self.table.setRowCount(len(rows))
        row_indexes = {}

        def number(row, key):
            return str(row[key]) if row.get(key) is not None else "未取得"

        for index, row in enumerate(rows):
            hour = row["hour"]
            row_indexes[hour] = index
            values = [f"{hour:02}:00–{hour + 1:02}:00",
                      *[number(row, k) for k in ("total_calls", "dialogues", "toss")],
                      format_duration(row["talk_seconds"]),
                      *[number(row, k) for k in ("new_prospects", "ng_with_dialogue")],
                      row["counter_status"]]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(index, column, item)
        latest = controller.store.latest(staff_id, day)
        if latest:
            def total(key):
                known = [r[key] for r in rows if r.get(key) is not None]
                return str(sum(known)) if known else "未取得"
            seconds = [r["talk_seconds"] for r in rows if r["talk_seconds"] is not None]
            self.total_label.setText(f"当日合計　コール {total('total_calls')}／対話 {total('dialogues')}／トス {total('toss')}／応答時間 {format_duration(sum(seconds))}\n"
                                     f"CTI上部の累計：総コール数 {latest['total_calls']}／有効コール {latest.get('valid_calls')}／新規見込 {latest['new_prospects']}"
                                     f"　最終取得 {datetime.fromisoformat(latest['observed_at']):%Y/%m/%d %H:%M:%S}")
            key = (staff_id, day)
            if self._scroll_key != key:
                self._scroll_key = key
                hour = datetime.fromisoformat(latest["observed_at"]).hour
                index = row_indexes.get(hour, 0)
                self.table.scrollToItem(self.table.item(index, 0), QAbstractItemView.PositionAtCenter)
        else:
            self.total_label.setText("この日の保存済み集計はありません。「今すぐ集計」で現在のログインスタッフの当日分を取得します。")
