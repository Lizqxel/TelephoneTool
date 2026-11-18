"""30分別ランキング・コピーと、Qt標準描画による推移グラフ。"""
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import (QApplication, QComboBox, QHeaderView, QHBoxLayout, QLabel, QPushButton,
                               QScrollArea, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from services.call_statistics import format_duration
from services.team_statistics_report import period_label, ranking, report_text, row_text, slot_label, summaries


class TrendChart(QWidget):
    def __init__(self):
        super().__init__()
        self.points = []
        self.metric_key = "seconds"
        self.setMinimumSize(900, 300)
        self.setMouseTracking(True)

    def set_points(self, reports, metric):
        self.metric_key = metric
        self.points = [(minute, report, (getattr(report, metric) if report.people else None)) for minute, report in reports]
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))
        area = QRectF(84, 24, max(1, self.width() - 104), max(1, self.height() - 80))
        maximum = max((value or 0 for _, _, value in self.points), default=0) or 1
        painter.setPen(QColor("#64748b"))
        for step in range(5):
            y = area.bottom() - area.height() * step / 4
            painter.setPen(QPen(QColor("#e2e8f0")))
            painter.drawLine(QPointF(area.left(), y), QPointF(area.right(), y))
            painter.setPen(QColor("#475569"))
            value = maximum * step / 4
            label = format_duration(value) if self.metric_key == "seconds" else f"{value:g}"
            painter.drawText(QRectF(0, y - 10, 78, 20), Qt.AlignRight | Qt.AlignVCenter, label)
        width = area.width() / 48
        for index, (minute, report, value) in enumerate(self.points):
            x = area.left() + index * width
            if value is not None:
                height = area.height() * value / maximum
                painter.fillRect(QRectF(x + 1, area.bottom() - max(1, height), max(1, width - 2), max(1, height)), QColor("#2563eb"))
            if minute % 120 == 0:
                painter.setPen(QColor("#475569"))
                painter.drawText(QRectF(x - 15, area.bottom() + 6, 44, 20), Qt.AlignCenter, slot_label(minute))
        painter.drawText(QRectF(84, self.height() - 24, self.width() - 104, 20), Qt.AlignCenter,
                         "時間帯（30分）　空欄＝未取得　／　マウスを合わせると実績と取得人数を表示")
        painter.end()

    def mouseMoveEvent(self, event):
        index = int((event.position().x() - 84) / max(1, self.width() - 104) * 48)
        if event.position().x() < 84 or not 0 <= index < len(self.points):
            self.setToolTip("")
            return
        minute, report, value = self.points[index]
        text = "未取得" if value is None else format_duration(value) if self.metric_key == "seconds" else str(value)
        self.setToolTip(f"{period_label(minute)}\n{text}\n取得済み {report.people}人 / 対象 {len(report.rows)}人")
        super().mouseMoveEvent(event)


class TeamStatisticsAnalysis(QWidget):
    def __init__(self):
        super().__init__()
        self.data = None
        self.unit_name = "全員"
        self.members = []
        self.report = ranking([])
        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        self.period = QComboBox()
        self.period.addItem("取得済み全日", None)
        for minute in range(0, 1440, 30):
            self.period.addItem(period_label(minute), minute)
        self.copy_button = QPushButton("表示内容をコピー")
        self.copy_notice = QLabel()
        controls.addWidget(QLabel("時間帯"))
        controls.addWidget(self.period)
        controls.addStretch()
        controls.addWidget(self.copy_notice)
        controls.addWidget(self.copy_button)
        layout.addLayout(controls)
        self.message = QLabel()
        self.message.setWordWrap(True)
        self.message.setTextFormat(Qt.PlainText)
        layout.addWidget(self.message)
        self.whole = QLabel()
        self.average = QLabel()
        for label in (self.whole, self.average):
            label.setWordWrap(True)
            label.setTextFormat(Qt.PlainText)
            layout.addWidget(label)
        self.tabs = QTabWidget()
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["氏名", "コール数", "対話数", "対話時間"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().hide()
        self.tabs.addTab(self.table, "対話時間ランキング")
        graph = QWidget()
        graph_layout = QVBoxLayout(graph)
        self.metric_selector = QComboBox()
        for text, key in (("対話時間", "seconds"), ("対話数", "dialogues"), ("コール数", "calls")):
            self.metric_selector.addItem(text, key)
        graph_layout.addWidget(self.metric_selector)
        self.chart = TrendChart()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.chart)
        graph_layout.addWidget(scroll)
        self.tabs.addTab(graph, "30分ごとの推移")
        layout.addWidget(self.tabs, 1)
        note = QLabel("対話時間＝見込＋対話有NGの通話時間。全体・平均は3項目すべて取得済みの人を対象にします。\n"
                      "取得途中の時間帯は途中実績です。全日では各人の最終取得時点までを集計します。グラフは選択日の全時間帯を表示します。")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.period.currentIndexChanged.connect(self.refresh)
        self.metric_selector.currentIndexChanged.connect(self.refresh_chart)
        self.copy_button.clicked.connect(self.copy)
        self.refresh()

    def set_data(self, data, unit):
        self.data, self.unit_name = data, unit
        self.members = [member for member in (data.interval_members if data else []) if unit in ("全員", "") or member.unit == unit]
        self.refresh()

    def refresh(self):
        self.copy_notice.clear()
        self.report = ranking(self.members, self.period.currentData())
        self.table.setRowCount(len(self.report.rows))
        for index, row in enumerate(self.report.rows):
            for column, value in enumerate(row_text(row)):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignLeft if column == 0 else Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(index, column, item)
        self.whole.setText(summaries(self.report)[0])
        self.average.setText(summaries(self.report)[1])
        stamps = [member.observed_at for member in self.members if member.observed_at]
        message = self.data.interval_error if self.data else "シートを読み込むと表示します。"
        if self.data and self.data.legacy_error:
            message += "　従来表・名簿は未取得のため、対象人数は30分データ受信済みの人のみです。"
        if stamps and self.data:
            message += f"　対象日 {self.data.day}　取得日時 {min(stamps)} ～ {max(stamps)}"
        self.message.setText(message)
        self.copy_button.setEnabled(bool(self.data and any(member.intervals for member in self.members)))
        self.refresh_chart()

    def refresh_chart(self):
        self.chart.set_points([(minute, ranking(self.members, minute)) for minute in range(0, 1440, 30)], self.metric_selector.currentData())

    def copy(self):
        if not self.data or not self.members:
            return
        QApplication.clipboard().setText(report_text(self.report, self.data.day, self.unit_name, self.period.currentData()))
        self.copy_notice.setText("コピーしました")
