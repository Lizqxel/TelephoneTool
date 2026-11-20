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


CANCELLED = "集計を中止しました。"


class StatisticsWorker(QThread):
    completed = Signal(object, str, str)  # 集計結果、エラー（中止は CANCELLED）、クリア・送信のお知らせ

    def __init__(self, store, numeric_unit, sender_factory=None, parent=None, day=None):
        super().__init__(parent)
        self.store = store
        self.numeric_unit = numeric_unit
        self.sender_factory = sender_factory or SheetSender.from_settings
        self.day = day  # None は当日。過去日（yyyy-mm-dd）はその日の架電履歴を集計する
        self.cancelled = threading.Event()

    def run(self):
        try:
            collector = CtiStatisticsCollector(self.cancelled)
            snapshot = (collector.collect(self.numeric_unit) if self.day is None
                        else collector.collect(self.numeric_unit, day=self.day))
            if self.cancelled.is_set():
                self.completed.emit(None, CANCELLED, "")
                return
            self.store.save(snapshot)
        except Exception as exc:
            # 中止した回は、途中で出た例外の文言ではなく「中止しました」と知らせる（CTIの検索条件は集計側で元に戻す）。
            self.completed.emit(None, CANCELLED if self.cancelled.is_set() else str(exc), "")
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
        else:
            notices.append("中止したため、シートへは送っていません。")
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
            return "シート送信済み（GASの次回自動更新後に反映。全員の集計で確認できます）"
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
        self.pending_day = None
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

    def collect(self, day=None):
        """当日分を集計する。day に過去日（yyyy-mm-dd）を渡すと、その日の架電履歴を集計する。"""
        if self.closing:
            return
        if self.busy:
            # 実行中に集計が重なった場合は、終わってから実行する（過去日の指定も覚えておく）。
            self.pending = True
            self.pending_day = day if day is not None else self.pending_day
            return
        worker = StatisticsWorker(self.store, self.store.preference("duration_unit", "seconds"), self.sender_factory, self, day=day)
        worker.completed.connect(self._completed)
        worker.finished.connect(self._finished)
        # 「集計中」になってから知らせる（先に知らせると、受け取った画面が集計中でないと誤解する）
        self.worker = worker
        worker.start()
        self.set_status("CTIを集計しています…")

    def cancel(self):
        """実行中の集計を中止する。CTIの検索日時・商材は集計側で元に戻し、保存済みの値は残す。"""
        if not self.busy or self.worker.cancelled.is_set():
            return
        self.pending = False  # 中止した直後に、待っていた集計を始めない（自動集計は次の00分・30分に動く）
        self.pending_day = None
        self.worker.cancelled.set()
        self.set_status("集計を中止しています…（CTIの検索日時・商材を元に戻してから止まります）")

    @property
    def cancelling(self):
        return self.busy and self.worker.cancelled.is_set()

    def _completed(self, snapshot, error, notice=""):
        if error == CANCELLED:
            self.set_status("集計を中止しました。保存済みの値は維持しています。")
            return
        if error:
            self.set_status(f"集計できませんでした：{error} 保存済みの値は維持しています。")
            return
        # クリアや送信だけ失敗した場合も、集計は保存済みであることを示す。
        saved = "（保存済み）" if any(word in notice for word in ("未完了", "失敗")) else ""
        if snapshot.history_only:
            parts = [f"{snapshot.acquired_at:%m/%d}分を集計済み{saved}：{snapshot.staff_name}"]
            if snapshot.history_counters is not None:
                parts.append(f"本人の{snapshot.acquired_at:%m/%d}の履歴 {snapshot.history_rows}件")
        else:
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
            day, self.pending_day = self.pending_day, None
            self.collect(day)

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
        self.cancel_button = QPushButton("中止")
        self.cancel_button.setToolTip("集計を中止します。CTIの検索日時・商材は元に戻し、保存済みの値は残します。")
        self.cancel_button.hide()
        self.automatic = QCheckBox("30分ごとに自動集計")
        self.automatic.setToolTip("起動時と毎時00分・30分に集計してシートへ送ります。")
        self.automatic.setChecked(controller.automatic)
        top.addWidget(QLabel("日付"))
        top.addWidget(self.day)
        top.addWidget(self.staff, 1)
        top.addWidget(self.automatic)
        top.addWidget(self.collect_button)
        top.addWidget(self.cancel_button)
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
        note = QLabel("コール＝架電履歴の件数（通話0秒・留守も含む）。対話＝見込＋対話有NG（NGは現アナNG・多重留守NG以外を対話有。"
                      "架電禁止NG・非勧奨登録NGは禁止）。\n"
                      "トス＝当日の見込のうち、見込検索のランクが「トス」のもの。応答時間＝IP電話通話時間。対話NG＝対話有NG。"
                      "「途中」の行は、その時刻までの件数です。")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.collect_button.clicked.connect(self._collect_clicked)
        self.cancel_button.clicked.connect(controller.cancel)
        self.automatic.toggled.connect(controller.set_automatic)
        self.duration_unit.currentIndexChanged.connect(self._unit_changed)
        self.day.dateChanged.connect(self.refresh)
        self.staff.currentIndexChanged.connect(self.refresh)
        controller.changed.connect(self.refresh)
        self.refresh()

    def selected_past_day(self):
        """日付欄が今日より前なら yyyy-mm-dd、今日なら None。"""
        selected = self.day.date()
        return selected.toString("yyyy-MM-dd") if selected < QDate.currentDate() else None

    def _collect_clicked(self):
        if self.day.date() > QDate.currentDate():
            return
        self.controller.collect(self.selected_past_day())

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
        past = self.selected_past_day()
        self.collect_button.setEnabled(not controller.busy and not controller.closing and self.day.date() <= today)
        self.collect_button.setText("集計中…" if controller.busy else "この日を集計" if past else "今すぐ集計")
        # 集計中は日付を固定し、終わったときに集計した日の表がそのまま見えるようにする。
        self.day.setEnabled(not controller.busy)
        self.cancel_button.setVisible(controller.busy and not controller.closing)
        self.cancel_button.setEnabled(controller.busy and not controller.cancelling)
        self.cancel_button.setText("中止中…" if controller.cancelling else "中止")
        self.collect_button.setToolTip(
            "CTIの架電履歴で、検索日時をこの日だけにして検索し、この日の分を集計してシートへ送ります。検索日時は元に戻します。"
            if past else "毎時00分・30分を待たずに、CTIから当日分を取得してシートへ送ります。日付欄は保存済みの表の表示日です。")
        self.duration_unit.setEnabled(not controller.busy)
        self.status_label.setText(controller.status)
        self.next_label.setText(f"次回の集計・送信 {controller.next_run:%H:%M}" if controller.automatic else "自動集計 停止中")
        day = self.day.date().toString("yyyy-MM-dd")
        staff_id = self.staff.currentData() or ""
        rows = controller.store.display_hours(staff_id, day)
        self.table.setRowCount(len(rows))
        row_indexes = {}
        latest = controller.store.latest(staff_id, day)
        observed = datetime.fromisoformat(latest["observed_at"]) if latest else None

        def number(row, key):
            return str(row[key]) if row.get(key) is not None else "未取得"

        for index, row in enumerate(rows):
            hour = row["hour"]
            row_indexes[hour] = index
            status = row["counter_status"]
            if observed and hour == observed.hour and past is None:
                # 今日の最後に集計した時刻の時間帯は、その時刻までの件数（終わった時間帯の0件と区別する）
                status = f"{status}・途中（{observed:%H:%M}時点）"
            values = [f"{hour:02}:00–{hour + 1:02}:00",
                      *[number(row, k) for k in ("total_calls", "dialogues", "toss")],
                      format_duration(row["talk_seconds"]),
                      *[number(row, k) for k in ("new_prospects", "ng_with_dialogue")],
                      status]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(index, column, item)
        if rows:
            def total(key):
                known = [r[key] for r in rows if r.get(key) is not None]
                return str(sum(known)) if known else "未取得"
            seconds = [r["talk_seconds"] for r in rows if r["talk_seconds"] is not None]
            text = f"{'当日' if past is None else 'この日の'}合計　コール {total('total_calls')}／対話 {total('dialogues')}／トス {total('toss')}／応答時間 {format_duration(sum(seconds))}"
            if latest:
                text += (f"\nCTI上部の累計：総コール数 {latest['total_calls']}／有効コール {latest.get('valid_calls')}／新規見込 {latest['new_prospects']}"
                         f"　最終取得 {observed:%Y/%m/%d %H:%M:%S}")
            self.total_label.setText(text)
            key = (staff_id, day)
            if self._scroll_key != key:
                self._scroll_key = key
                index = row_indexes.get(observed.hour, 0) if observed else 0
                self.table.scrollToItem(self.table.item(index, 0), QAbstractItemView.PositionAtCenter)
        elif past:
            self.total_label.setText("この日の保存済み集計はありません。「この日を集計」で、CTIの架電履歴からこの日の分を集計します。")
        else:
            self.total_label.setText("この日の保存済み集計はありません。「今すぐ集計」で現在のログインスタッフの当日分を取得します。")
