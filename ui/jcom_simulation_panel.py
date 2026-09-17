"""J:COM検索条件・進捗・料金カードを表示する専用パネル。"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from services.jcom_simulation_models import JcomSimulationResult, ResidenceType


class JcomSimulationPanel(QGroupBox):
    search_requested = Signal()
    cancel_requested = Signal()
    residence_changed = Signal()

    def __init__(self, parent=None):
        super().__init__("J:COM 料金シミュレーション", parent)
        layout = QVBoxLayout(self)

        residence_row = QHBoxLayout()
        residence_row.addWidget(QLabel("住居タイプ"))
        self.detached_radio = QRadioButton("戸建住宅")
        self.apartment_radio = QRadioButton("集合住宅")
        self.residence_group = QButtonGroup(self)
        self.residence_group.addButton(self.detached_radio)
        self.residence_group.addButton(self.apartment_radio)
        self.apartment_radio.setChecked(True)
        residence_row.addWidget(self.detached_radio)
        residence_row.addWidget(self.apartment_radio)
        residence_row.addStretch(1)
        layout.addLayout(residence_row)

        self.condition_label = QLabel(
            "入力済みの郵便番号・住所・生年月日を使用します。\n"
            "ネットのみ／J:COM未利用／光(N) 1Gコース"
        )
        self.condition_label.setWordWrap(True)
        self.condition_label.setStyleSheet("color: #455A64; padding: 4px;")
        layout.addWidget(self.condition_label)

        button_row = QHBoxLayout()
        self.search_button = QPushButton("料金シミュレーション検索")
        self.search_button.setStyleSheet(
            "QPushButton { background:#0068B7; color:white; padding:8px; border-radius:4px; }"
            "QPushButton:hover { background:#00579A; }"
        )
        self.cancel_button = QPushButton("キャンセル")
        self.cancel_button.setEnabled(False)
        button_row.addWidget(self.search_button)
        button_row.addWidget(self.cancel_button)
        layout.addLayout(button_row)

        self.progress_label = QLabel("未検索")
        self.progress_label.setWordWrap(True)
        layout.addWidget(self.progress_label)

        self.tabs = QTabWidget()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setPlaceholderText("検索結果はここに表示されます")
        text_page = QWidget()
        text_layout = QVBoxLayout(text_page)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.addWidget(self.result_text)
        self.copy_button = QPushButton("料金テキストをコピー")
        self.copy_button.setEnabled(False)
        text_layout.addWidget(self.copy_button)
        self.tabs.addTab(text_page, "テキスト")

        self.raw_text = QTextEdit()
        self.raw_text.setReadOnly(True)
        self.raw_text.setPlaceholderText("J:COM料金結果領域の取得原文")
        self.tabs.addTab(self.raw_text, "取得原文")

        self.image_label = QLabel("画像未取得")
        self.image_label.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.image_label.setWordWrap(True)

        image_page = QWidget()
        image_layout = QVBoxLayout(image_page)
        image_layout.setContentsMargins(0, 0, 0, 0)
        image_tools = QHBoxLayout()
        self.fit_image_button = QPushButton("幅に合わせる")
        self.actual_image_button = QPushButton("100%")
        self.zoom_out_button = QPushButton("−")
        self.zoom_in_button = QPushButton("＋")
        self.open_image_button = QPushButton("別画面で開く")
        self.image_zoom_label = QLabel("100%")
        for button in (
            self.fit_image_button,
            self.actual_image_button,
            self.zoom_out_button,
            self.zoom_in_button,
            self.open_image_button,
        ):
            button.setEnabled(False)
        image_tools.addWidget(self.fit_image_button)
        image_tools.addWidget(self.actual_image_button)
        image_tools.addWidget(self.zoom_out_button)
        image_tools.addWidget(self.zoom_in_button)
        image_tools.addWidget(self.image_zoom_label)
        image_tools.addStretch(1)
        image_tools.addWidget(self.open_image_button)
        image_layout.addLayout(image_tools)

        self.image_scroll = QScrollArea()
        self.image_scroll.setWidget(self.image_label)
        self.image_scroll.setWidgetResizable(False)
        image_layout.addWidget(self.image_scroll)
        self.tabs.addTab(image_page, "サイト画面")
        self._source_pixmap = QPixmap()
        self._image_scale = 1.0
        self._fit_image = True
        self._screenshot_path = ""
        self.tabs.setVisible(False)
        layout.addWidget(self.tabs)

        self.search_button.clicked.connect(self.search_requested)
        self.cancel_button.clicked.connect(self.cancel_requested)
        self.detached_radio.toggled.connect(self._on_residence_changed)
        self.apartment_radio.toggled.connect(self._on_residence_changed)
        self.copy_button.clicked.connect(self._copy_result)
        self.fit_image_button.clicked.connect(self._fit_image_to_width)
        self.actual_image_button.clicked.connect(lambda: self._set_image_scale(1.0))
        self.zoom_out_button.clicked.connect(lambda: self._set_image_scale(self._image_scale / 1.25))
        self.zoom_in_button.clicked.connect(lambda: self._set_image_scale(self._image_scale * 1.25))
        self.open_image_button.clicked.connect(self._open_image_external)
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _on_residence_changed(self, checked):
        if checked:
            self.invalidate_result("住居タイプが変更されました。再検索してください。")
            self.residence_changed.emit()

    def residence_type(self) -> ResidenceType:
        return ResidenceType.DETACHED if self.detached_radio.isChecked() else ResidenceType.APARTMENT

    def set_age_condition(self, text: str):
        self.condition_label.setText(
            f"{text}\nネットのみ／J:COM未利用／光(N) 1Gコース"
        )

    def set_running(self, running: bool):
        self.search_button.setEnabled(not running)
        self.cancel_button.setEnabled(running)
        if running:
            self.progress_label.setText("検索を開始しています…")

    def set_progress(self, message: str):
        self.progress_label.setText(message)

    def show_result(self, result: JcomSimulationResult):
        self.set_running(False)
        if result.successful:
            if result.address_approximations:
                progress = "近い住所で取得：入力住所の料金ではありません"
            elif result.partial_address:
                progress = "住所の一部未確認：料金はサイトが確定した住所に対する結果です"
            else:
                progress = "取得完了"
        else:
            progress = result.error_message or result.status.value
        self.progress_label.setText(progress)
        text = result.display_text()
        self.result_text.setPlainText(text)
        self.raw_text.setPlainText(result.raw_text or "原文未取得")
        self.copy_button.setEnabled(bool(text))
        self.tabs.setVisible(True)
        self.tabs.setCurrentIndex(0)
        if result.screenshot_path and Path(result.screenshot_path).is_file():
            self._screenshot_path = result.screenshot_path
            self._source_pixmap = QPixmap(result.screenshot_path)
            self.image_label.setToolTip(result.screenshot_path)
            self._set_image_controls_enabled(not self._source_pixmap.isNull())
            self._fit_image = True
            QTimer.singleShot(0, self._fit_image_to_width)
        else:
            self._screenshot_path = ""
            self._source_pixmap = QPixmap()
            self._set_image_controls_enabled(False)
            self.image_label.setPixmap(QPixmap())
            self.image_label.setText("画像未取得（料金テキストは利用できます）")
            self.image_label.adjustSize()

    def invalidate_result(self, message: str = "入力条件が変更されました。"):
        self.result_text.clear()
        self.raw_text.clear()
        self.copy_button.setEnabled(False)
        self.tabs.setVisible(False)
        self._screenshot_path = ""
        self._source_pixmap = QPixmap()
        self._set_image_controls_enabled(False)
        self.image_label.setPixmap(QPixmap())
        self.image_label.setText("画像未取得")
        self.progress_label.setText(message)

    def _copy_result(self):
        text = self.result_text.toPlainText()
        if text:
            QApplication.clipboard().setText(text)
            self.progress_label.setText("料金テキストをコピーしました")

    def _set_image_controls_enabled(self, enabled: bool):
        for button in (
            self.fit_image_button,
            self.actual_image_button,
            self.zoom_out_button,
            self.zoom_in_button,
            self.open_image_button,
        ):
            button.setEnabled(enabled)

    def _apply_image_scale(self):
        if self._source_pixmap.isNull():
            return
        width = max(1, round(self._source_pixmap.width() * self._image_scale))
        scaled = self._source_pixmap.scaledToWidth(width, Qt.SmoothTransformation)
        self.image_label.setText("")
        self.image_label.setPixmap(scaled)
        self.image_label.resize(scaled.size())
        self.image_zoom_label.setText(f"{round(self._image_scale * 100)}%")

    def _set_image_scale(self, scale: float):
        self._fit_image = False
        self._image_scale = min(3.0, max(0.25, scale))
        self._apply_image_scale()

    def _fit_image_to_width(self):
        if self._source_pixmap.isNull():
            return
        viewport_width = max(1, self.image_scroll.viewport().width() - 4)
        self._fit_image = True
        self._image_scale = min(1.0, viewport_width / self._source_pixmap.width())
        self._apply_image_scale()

    def _on_tab_changed(self, index: int):
        if index == 2 and self._fit_image and not self._source_pixmap.isNull():
            QTimer.singleShot(0, self._fit_image_to_width)

    def _open_image_external(self):
        if self._screenshot_path and Path(self._screenshot_path).is_file():
            QDesktopServices.openUrl(QUrl.fromLocalFile(self._screenshot_path))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._fit_image and not self._source_pixmap.isNull():
            QTimer.singleShot(0, self._fit_image_to_width)
