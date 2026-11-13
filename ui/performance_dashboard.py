"""Native PySide6 dashboard for daily CTI performance."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


def format_duration(seconds: int, long: bool = False) -> str:
    seconds = max(0, int(seconds or 0))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if long and hours:
        return f"{hours}時間{minutes:02d}分{secs:02d}秒"
    if hours:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"


def percentage(numerator: int, denominator: int) -> str:
    return f"{(numerator / denominator * 100):.1f}%" if denominator else "0.0%"


class MetricCard(QGroupBox):
    def __init__(self, title: str, parent=None) -> None:
        super().__init__(title, parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        self.value_label = QLabel("0")
        self.value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.value_label.setStyleSheet("font-size: 18pt; font-weight: 600; color: #1F4E78;")
        layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(str(value))


class StaffDetailDialog(QDialog):
    def __init__(self, staff: Dict[str, Any], business_date: str, parent=None) -> None:
        super().__init__(parent)
        self.staff = staff
        self.business_date = business_date
        self.setWindowTitle(f"個人実績 - {staff.get('staff_name', '')}")
        self.resize(900, 650)
        layout = QVBoxLayout(self)

        heading = QLabel(f"{staff.get('staff_name', '')}　{business_date}")
        heading.setStyleSheet("font-size: 15pt; font-weight: 600;")
        layout.addWidget(heading)

        metrics = QGridLayout()
        cards = [
            ("総コール", staff.get("calls", 0)),
            ("対話（有効コール）", staff.get("dialogues", 0)),
            ("トス（新規見込）", staff.get("tosses", 0)),
            ("対話時間", format_duration(staff.get("talk_seconds", 0))),
        ]
        for column, (title, value) in enumerate(cards):
            card = MetricCard(title)
            card.set_value(str(value))
            metrics.addWidget(card, 0, column)
        layout.addLayout(metrics)

        layout.addWidget(QLabel("時間帯別"))
        self.hourly_table = QTableWidget(0, 7)
        self.hourly_table.setHorizontalHeaderLabels(
            ["時間帯", "コール", "対話", "対話率", "トス", "トス率", "対話時間"]
        )
        self.hourly_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.hourly_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.hourly_table.verticalHeader().setVisible(False)
        self.hourly_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.hourly_table, 2)

        layout.addWidget(QLabel("架電明細（顧客情報は保存しません）"))
        self.detail_table = QTableWidget(0, 5)
        self.detail_table.setHorizontalHeaderLabels(
            ["架電時刻", "架電結果", "対話", "トス", "IP電話通話時間"]
        )
        self.detail_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.detail_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.detail_table.verticalHeader().setVisible(False)
        self.detail_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.detail_table, 3)

        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.accept)
        buttons = QHBoxLayout()
        buttons.addStretch()
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self._populate()

    def _populate(self) -> None:
        hours = sorted(self.staff.get("hours", []), key=lambda item: int(item.get("hour", 0)))
        self.hourly_table.setRowCount(len(hours))
        for row_index, item in enumerate(hours):
            calls = int(item.get("calls", 0))
            dialogues = int(item.get("dialogues", 0))
            tosses = int(item.get("tosses", 0))
            values = [
                f"{int(item.get('hour', 0))}時台", calls, dialogues,
                percentage(dialogues, calls), tosses, percentage(tosses, dialogues),
                format_duration(item.get("talk_seconds", 0)),
            ]
            self._set_row(self.hourly_table, row_index, values)

        details = sorted(self.staff.get("details", []), key=lambda item: item.get("called_at", ""))
        self.detail_table.setRowCount(len(details))
        for row_index, item in enumerate(details):
            called_at = str(item.get("called_at", ""))
            try:
                time_text = datetime.fromisoformat(called_at).strftime("%H:%M:%S")
            except ValueError:
                time_text = called_at
            values = [
                time_text,
                item.get("result") or "－",
                "有効" if item.get("is_dialogue") else "－",
                "新規見込" if item.get("is_toss") else "－",
                format_duration(item.get("talk_seconds", 0)),
            ]
            self._set_row(self.detail_table, row_index, values)

    @staticmethod
    def _set_row(table: QTableWidget, row: int, values: List[Any]) -> None:
        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value))
            if column > 0:
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row, column, item)


class PerformanceDashboardDialog(QDialog):
    def __init__(self, manager, parent=None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.current_staff_rows: List[Dict[str, Any]] = []
        self.setWindowTitle("本日の実績")
        self.resize(1080, 680)
        self.setMinimumSize(820, 520)
        self._build_ui()

        self.manager.updated.connect(self.reload)
        self.manager.status_changed.connect(self.sync_status.setText)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.setInterval(30000)
        self.refresh_timer.timeout.connect(self.refresh_from_server)
        self.refresh_timer.start()
        self.reload()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        controls.addWidget(QLabel("日付"))
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy/MM/dd")
        self.date_edit.dateChanged.connect(self._date_changed)
        controls.addWidget(self.date_edit)

        controls.addWidget(QLabel("時間帯"))
        self.hour_combo = QComboBox()
        self.hour_combo.addItem("全日", None)
        for hour in range(24):
            self.hour_combo.addItem(f"{hour}時台", hour)
        self.hour_combo.currentIndexChanged.connect(self.reload)
        controls.addWidget(self.hour_combo)

        self.refresh_button = QPushButton("最新に更新")
        self.refresh_button.clicked.connect(self.refresh_from_server)
        controls.addWidget(self.refresh_button)

        self.import_button = QPushButton("CTI履歴CSVを取り込む")
        self.import_button.setToolTip("架電履歴のCSVを使って当日実績を再構築します")
        self.import_button.clicked.connect(self.import_csv)
        controls.addWidget(self.import_button)

        controls.addStretch()
        self.sync_status = QLabel("ローカル保存中")
        self.sync_status.setStyleSheet("color: #287A45;")
        controls.addWidget(self.sync_status)
        layout.addLayout(controls)

        metrics = QGridLayout()
        self.calls_card = MetricCard("総コール")
        self.dialogues_card = MetricCard("対話（有効コール）")
        self.tosses_card = MetricCard("トス（新規見込）")
        self.talk_card = MetricCard("対話時間")
        for index, card in enumerate(
            [self.calls_card, self.dialogues_card, self.tosses_card, self.talk_card]
        ):
            metrics.addWidget(card, 0, index)
        layout.addLayout(metrics)

        self.staff_table = QTableWidget(0, 8)
        self.staff_table.setHorizontalHeaderLabels(
            ["担当者", "コール", "対話", "対話率", "トス", "トス率", "対話時間", "平均対話"]
        )
        self.staff_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.staff_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.staff_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.staff_table.verticalHeader().setVisible(False)
        header = self.staff_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, 8):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        self.staff_table.cellClicked.connect(self._staff_clicked)
        layout.addWidget(self.staff_table, 1)

        hint = QLabel("担当者名をクリックすると、時間帯別実績と架電時刻の明細を確認できます。")
        hint.setStyleSheet("color: #667280;")
        layout.addWidget(hint)

    def _business_date(self) -> str:
        return self.date_edit.date().toString("yyyy-MM-dd")

    def _date_changed(self) -> None:
        self.reload()
        self.refresh_from_server()

    def refresh_from_server(self) -> None:
        self.manager.sync_now(self._business_date())

    def import_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "CTI架電履歴CSVを選択", "", "CSVファイル (*.csv);;すべてのファイル (*)"
        )
        if not path:
            return
        try:
            count = self.manager.import_history_csv(path)
            QMessageBox.information(self, "取り込み完了", f"{count}件の架電履歴を取り込みました。")
        except Exception as exc:
            QMessageBox.critical(self, "取り込みエラー", f"CSVの取り込みに失敗しました。\n{exc}")

    def reload(self) -> None:
        dashboard = self.manager.dashboard(self._business_date())
        selected_hour: Optional[int] = self.hour_combo.currentData()
        staff_rows: List[Dict[str, Any]] = []
        totals = {"calls": 0, "dialogues": 0, "tosses": 0, "talk_seconds": 0}
        for staff in dashboard.get("staff", []):
            if selected_hour is None:
                row = dict(staff)
            else:
                hour_row = next(
                    (item for item in staff.get("hours", []) if int(item.get("hour", -1)) == selected_hour),
                    {},
                )
                row = dict(staff)
                for key in totals:
                    row[key] = int(hour_row.get(key, 0))
                row["details"] = [
                    item for item in staff.get("details", []) if int(item.get("hour", -1)) == selected_hour
                ]
            if any(int(row.get(key, 0)) for key in totals):
                staff_rows.append(row)
                for key in totals:
                    totals[key] += int(row.get(key, 0))

        self.current_staff_rows = staff_rows
        self.calls_card.set_value(str(totals["calls"]))
        self.dialogues_card.set_value(str(totals["dialogues"]))
        self.tosses_card.set_value(str(totals["tosses"]))
        self.talk_card.set_value(format_duration(totals["talk_seconds"]))

        self.staff_table.setRowCount(len(staff_rows))
        for row_index, staff in enumerate(staff_rows):
            calls = int(staff.get("calls", 0))
            dialogues = int(staff.get("dialogues", 0))
            tosses = int(staff.get("tosses", 0))
            talk_seconds = int(staff.get("talk_seconds", 0))
            name = str(staff.get("staff_name", ""))
            if staff.get("corrected"):
                name += " 〔修正あり〕"
            values = [
                name, calls, dialogues, percentage(dialogues, calls), tosses,
                percentage(tosses, dialogues), format_duration(talk_seconds),
                format_duration(talk_seconds // dialogues if dialogues else 0),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setForeground(QColor("#1769AA"))
                    item.setToolTip("クリックして個人詳細を表示")
                else:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.staff_table.setItem(row_index, column, item)

    def _staff_clicked(self, row: int, column: int) -> None:
        if column != 0 or not (0 <= row < len(self.current_staff_rows)):
            return
        StaffDetailDialog(self.current_staff_rows[row], self._business_date(), self).exec()

