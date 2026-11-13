"""30分周期の集計処理と、日別・時間帯別の通話集計表。"""

from datetime import datetime
import os
from pathlib import Path
import threading

from PySide6.QtCore import QDate, QObject, QThread, QTimer, Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDateEdit, QDialog, QHBoxLayout,
                               QHeaderView, QLabel, QPushButton, QTableWidget,
                               QTableWidgetItem, QVBoxLayout)

from services.call_statistics import JST, StatisticsStore, format_duration, next_half_hour
from services.cti_statistics_service import CtiStatisticsCollector


class StatisticsWorker(QThread):
    completed = Signal(object, str)

    def __init__(self, store, numeric_unit, parent=None):
        super().__init__(parent)
        self.store = store
        self.numeric_unit = numeric_unit
        self.cancelled = threading.Event()

    def run(self):
        try:
            snapshot = CtiStatisticsCollector(self.cancelled).collect(self.numeric_unit)
            if self.cancelled.is_set():
                return
            self.store.save(snapshot)
            self.completed.emit(snapshot, "")
        except Exception as exc:
            self.completed.emit(None, str(exc))


class CallStatisticsController(QObject):
    changed = Signal()
    status_changed = Signal(str)

    def __init__(self, parent=None, store=None, autostart=True):
        super().__init__(parent)
        if store is None:
            base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "TelephoneTool"
            store = StatisticsStore(base / "call_statistics.sqlite3")
        self.store = store
        self.worker = None
        self.dialog = None
        self.closing = False
        self.status = "初回の集計前です。上部カウンターは初回を基準に差分を記録します。"
        self.automatic = self.store.preference("automatic", "1") == "1"
        self.timer = QTimer(self)
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
            self.timer.start(max(1, int((self.next_run - now).total_seconds() * 1000)))

    def _automatic_tick(self):
        self.collect()
        self.schedule()

    def set_automatic(self, enabled):
        self.automatic = enabled
        self.store.set_preference("automatic", "1" if enabled else "0")
        self.schedule()
        self.set_status("30分ごとの自動集計を有効にしました。" if enabled else "自動集計を停止しました。集計ボタンから手動で取得できます。")

    def set_status(self, message):
        self.status = message
        self.status_changed.emit(message)
        self.changed.emit()

    def collect(self):
        if self.busy or self.closing:
            return
        self.set_status("CTIを集計しています…")
        worker = StatisticsWorker(self.store, self.store.preference("duration_unit", "seconds"), self)
        self.worker = worker
        worker.completed.connect(self._completed)
        worker.finished.connect(self._finished)
        worker.start()
        self.changed.emit()

    def _completed(self, snapshot, error):
        if error:
            self.set_status(f"集計できませんでした：{error} 保存済みの値は維持しています。")
        else:
            self.set_status(f"{snapshot.acquired_at:%H:%M} 集計済み：{snapshot.staff_name}／履歴 {snapshot.history_rows}件")

    def _finished(self):
        worker = self.worker
        self.worker = None
        if worker:
            worker.deleteLater()
        self.changed.emit()

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
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self._scroll_key = None
        self.setWindowTitle("時間帯別 通話集計")
        self.resize(920, 690)
        layout = QVBoxLayout(self)
        top = QHBoxLayout()
        self.day = QDateEdit(QDate.currentDate())
        self.day.setCalendarPopup(True)
        self.day.setDisplayFormat("yyyy/MM/dd")
        self.staff = QComboBox()
        self.collect_button = QPushButton("集計")
        self.collect_button.setToolTip("現在のログインスタッフの当日分を取得します。日付欄は保存済みの表の表示日です。")
        self.automatic = QCheckBox("30分ごとに自動集計")
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
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["時間帯", "応答時間", "総コール数", "新規見込", "対話NG", "取得状態"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.table.setToolTip("過去のコール数も履歴から集計します。新規見込・対話有NGは履歴で明確に判定できる分を優先し、それ以外は取得済みの累計差分を表示します。")
        layout.addWidget(self.table, 1)
        self.total_label = QLabel()
        self.total_label.setTextFormat(Qt.PlainText)
        self.total_label.setWordWrap(True)
        layout.addWidget(self.total_label)
        note = QLabel("集計開始前のコールも履歴から集計し、最初のログの時間帯から表示します。通話0秒のコールも件数に含めます。\n"
                      "応答時間＝IP電話通話時間。対話NG＝対話有NG。履歴で判別できない結果は取得済みの差分を使い、復元できない値は未取得です。")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.collect_button.clicked.connect(controller.collect)
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
        self.collect_button.setText("集計中…" if controller.busy else "集計")
        self.duration_unit.setEnabled(not controller.busy)
        self.status_label.setText(controller.status)
        self.next_label.setText(f"次回 {controller.next_run:%H:%M}" if controller.automatic else "自動集計 停止中")
        day = self.day.date().toString("yyyy-MM-dd")
        staff_id = self.staff.currentData() or ""
        rows = controller.store.display_hours(staff_id, day)
        self.table.setRowCount(len(rows))
        row_indexes = {}
        for index, row in enumerate(rows):
            hour = row["hour"]
            row_indexes[hour] = index
            values = [f"{hour:02}:00–{hour + 1:02}:00", format_duration(row["talk_seconds"]),
                      *[str(row[k]) if row.get(k) is not None else "未取得" for k in ("total_calls", "new_prospects", "ng_with_dialogue")],
                      row["counter_status"]]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(index, column, item)
        seconds = [r["talk_seconds"] for r in rows if r["talk_seconds"] is not None]
        latest = controller.store.latest(staff_id, day)
        if latest:
            self.total_label.setText(f"当日の応答時間 {format_duration(sum(seconds))}　｜　直近のCTI累計：総コール数 {latest['total_calls']}／新規見込 {latest['new_prospects']}／対話NG {latest['ng_with_dialogue']}\n"
                                     f"最終取得 {datetime.fromisoformat(latest['observed_at']):%Y/%m/%d %H:%M:%S}")
            key = (staff_id, day)
            if self._scroll_key != key:
                self._scroll_key = key
                hour = datetime.fromisoformat(latest["observed_at"]).hour
                index = row_indexes.get(hour, 0)
                self.table.scrollToItem(self.table.item(index, 0), QAbstractItemView.PositionAtCenter)
        else:
            self.total_label.setText("この日の保存済み集計はありません。集計ボタンは現在のログインスタッフの当日分を取得します。")
