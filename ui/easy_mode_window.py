"""
イージーモードのメインウィンドウを提供するモジュール

このモジュールは、イージーモード（ガイド付きモード）のメインウィンドウを実装します。
主な機能：
- ステップバイステップの操作ガイド
- 地域検索の簡易化
- 電話番号の入力支援
- 検索結果の視覚的な表示
- 設定の簡易化

制限事項：
- 地域検索は最大50件まで表示
- 電話番号は10桁まで入力可能
- 検索結果は最大50件まで表示
"""

import os
import sys
import json
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QMessageBox, QMenuBar,
    QMenu, QStatusBar, QProgressBar, QApplication, QFileDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QDialog, QFormLayout, QSpinBox, QCheckBox, QDialogButtonBox,
    QGroupBox, QRadioButton, QButtonGroup, QFrame
)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
from PySide6.QtGui import QAction, QIcon, QFont, QPixmap, QPalette, QColor

from .base_window import BaseWindow
from .custom_widgets import CustomComboBox
from ..services.area_search_service import AreaSearchService
from ..services.cti_service import CTIService
from ..utils.logger import setup_logger
from utils.settings import settings
from version import check_version

# 高齢者向けのスタイル設定
LARGE_FONT = QFont("MS Gothic", 16)
MEDIUM_FONT = QFont("MS Gothic", 14)
SMALL_FONT = QFont("MS Gothic", 12)

BUTTON_STYLE = """
    QPushButton {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 15px 30px;
        font-size: 16px;
        border-radius: 8px;
        min-width: 200px;
    }
    QPushButton:hover {
        background-color: #45a049;
    }
    QPushButton:pressed {
        background-color: #3e8e41;
    }
    QPushButton:disabled {
        background-color: #cccccc;
    }
"""

GROUP_BOX_STYLE = """
    QGroupBox {
        font-size: 16px;
        font-weight: bold;
        border: 2px solid #4CAF50;
        border-radius: 8px;
        margin-top: 1em;
        padding-top: 1em;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
    }
"""

LABEL_STYLE = """
    QLabel {
        font-size: 16px;
        color: #333333;
        padding: 5px;
    }
"""

class EasyModeWindow(BaseWindow):
    """イージーモードのメインウィンドウクラス"""
    
    def __init__(self):
        """ウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("電話番号検索ツール - イージーモード")
        self.current_step = 1
        self.setup_ui()
        self.setup_connections()
        self.load_settings()
        
    def setup_ui(self):
        """UIの初期設定"""
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(20)  # ウィジェット間の間隔を広げる
        layout.setContentsMargins(20, 20, 20, 20)  # マージンを設定
        
        # ガイド表示エリア
        guide_group = self.create_guide_group()
        layout.addWidget(guide_group)
        
        # 操作エリア
        operation_group = self.create_operation_group()
        layout.addWidget(operation_group)
        
        # 検索結果表示エリア
        result_group = self.create_result_group()
        layout.addWidget(result_group)
        
        # ステータスバーの設定
        self.statusBar().showMessage("準備完了")
        
    def create_guide_group(self) -> QWidget:
        """ガイド表示グループの作成"""
        group = QGroupBox("操作ガイド")
        group.setStyleSheet(GROUP_BOX_STYLE)
        layout = QVBoxLayout(group)
        layout.setSpacing(15)
        
        # ステップ表示
        self.step_label = QLabel("ステップ1: 検索方法を選択してください")
        self.step_label.setFont(LARGE_FONT)
        self.step_label.setStyleSheet("color: #2E7D32; font-weight: bold;")
        layout.addWidget(self.step_label)
        
        # ガイド説明
        self.guide_text = QLabel("")
        self.guide_text.setFont(MEDIUM_FONT)
        self.guide_text.setWordWrap(True)
        self.guide_text.setStyleSheet(LABEL_STYLE)
        layout.addWidget(self.guide_text)
        
        return group
        
    def create_operation_group(self) -> QWidget:
        """操作グループの作成"""
        group = QGroupBox("操作パネル")
        group.setStyleSheet(GROUP_BOX_STYLE)
        layout = QVBoxLayout(group)
        layout.setSpacing(20)
        
        # 検索方法選択
        search_method_group = QButtonGroup()
        self.area_radio = QRadioButton("地域から検索")
        self.phone_radio = QRadioButton("電話番号から検索")
        self.area_radio.setFont(MEDIUM_FONT)
        self.phone_radio.setFont(MEDIUM_FONT)
        search_method_group.addButton(self.area_radio)
        search_method_group.addButton(self.phone_radio)
        
        method_layout = QHBoxLayout()
        method_layout.setSpacing(30)
        method_layout.addWidget(self.area_radio)
        method_layout.addWidget(self.phone_radio)
        layout.addLayout(method_layout)
        
        # 地域検索エリア
        self.area_search_widget = self.create_area_search_widget()
        layout.addWidget(self.area_search_widget)
        
        # 電話番号検索エリア
        self.phone_search_widget = self.create_phone_search_widget()
        layout.addWidget(self.phone_search_widget)
        
        # 次へボタン
        self.next_button = QPushButton("次へ")
        self.next_button.setFont(LARGE_FONT)
        self.next_button.setStyleSheet(BUTTON_STYLE)
        layout.addWidget(self.next_button, alignment=Qt.AlignmentFlag.AlignCenter)
        
        return group
        
    def create_area_search_widget(self) -> QWidget:
        """地域検索ウィジェットの作成"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setSpacing(20)
        
        # 都道府県選択
        self.prefecture_combo = CustomComboBox()
        self.prefecture_combo.setFont(MEDIUM_FONT)
        self.prefecture_combo.addItem("都道府県を選択", "")
        self.prefecture_combo.setMinimumHeight(40)
        layout.addWidget(self.prefecture_combo)
        
        # 市区町村選択
        self.city_combo = CustomComboBox()
        self.city_combo.setFont(MEDIUM_FONT)
        self.city_combo.addItem("市区町村を選択", "")
        self.city_combo.setMinimumHeight(40)
        layout.addWidget(self.city_combo)
        
        return widget
        
    def create_phone_search_widget(self) -> QWidget:
        """電話番号検索ウィジェットの作成"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setSpacing(20)
        
        # 電話番号入力
        self.phone_input = QLineEdit()
        self.phone_input.setFont(MEDIUM_FONT)
        self.phone_input.setPlaceholderText("電話番号を入力（例：03-1234-5678）")
        self.phone_input.setMaxLength(10)
        self.phone_input.setMinimumHeight(40)
        layout.addWidget(self.phone_input)
        
        return widget
        
    def create_result_group(self) -> QWidget:
        """検索結果表示グループの作成"""
        group = QGroupBox("検索結果")
        group.setStyleSheet(GROUP_BOX_STYLE)
        layout = QVBoxLayout(group)
        
        # 結果テーブル
        self.result_table = QTableWidget()
        self.result_table.setFont(MEDIUM_FONT)
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels([
            "電話番号", "地域", "事業者", "備考"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_table.setMinimumHeight(300)
        self.result_table.setStyleSheet("""
            QTableWidget {
                font-size: 14px;
                gridline-color: #dddddd;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                padding: 8px;
                font-size: 14px;
                font-weight: bold;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QTableWidget::item:selected {
                background-color: #e8f5e9;
                color: black;
            }
        """)
        layout.addWidget(self.result_table)
        
        return group
        
    def setup_connections(self):
        """シグナルとスロットの接続"""
        # 検索方法選択
        self.area_radio.toggled.connect(self.on_search_method_changed)
        self.phone_radio.toggled.connect(self.on_search_method_changed)
        
        # 地域検索関連
        self.prefecture_combo.currentIndexChanged.connect(self.on_prefecture_changed)
        
        # 次へボタン
        self.next_button.clicked.connect(self.on_next_clicked)
        
        # 結果テーブル関連
        self.result_table.itemDoubleClicked.connect(self.on_result_double_clicked)
        
    def on_search_method_changed(self, checked: bool):
        """検索方法が変更されたときの処理"""
        if not checked:
            return
            
        if self.area_radio.isChecked():
            self.area_search_widget.setVisible(True)
            self.phone_search_widget.setVisible(False)
            self.update_guide_text("地域を選択して検索を開始します。")
        else:
            self.area_search_widget.setVisible(False)
            self.phone_search_widget.setVisible(True)
            self.update_guide_text("電話番号を入力して検索を開始します。")
            
    def on_prefecture_changed(self, index: int):
        """都道府県が変更されたときの処理"""
        prefecture = self.prefecture_combo.currentData()
        if prefecture:
            self.update_city_combo(prefecture)
            
    def on_next_clicked(self):
        """次へボタンがクリックされたときの処理"""
        if self.area_radio.isChecked():
            self.search_by_area()
        else:
            self.search_by_phone()
            
    def on_result_double_clicked(self, item: QTableWidgetItem):
        """検索結果がダブルクリックされたときの処理"""
        row = item.row()
        phone_number = self.result_table.item(row, 0).text()
        self.copy_to_clipboard(phone_number)
        
    def update_guide_text(self, text: str):
        """ガイドテキストの更新"""
        self.guide_text.setText(text)
        
    def update_city_combo(self, prefecture: str):
        """市区町村コンボボックスの更新"""
        self.city_combo.clear()
        self.city_combo.addItem("市区町村を選択", "")
        
        # TODO: 市区町村データの取得と設定
        
    def search_by_area(self):
        """地域による検索"""
        try:
            prefecture = self.prefecture_combo.currentData()
            city = self.city_combo.currentData()
            
            if not prefecture or not city:
                QMessageBox.warning(self, "警告", "都道府県と市区町村を選択してください。")
                return
                
            # TODO: 地域検索の実装
        except Exception as e:
            logging.error(f"地域検索中にエラーが発生しました: {str(e)}")
            QMessageBox.critical(self, "エラー", "検索中にエラーが発生しました。")
            
    def search_by_phone(self):
        """電話番号による検索"""
        try:
            phone_number = self.phone_input.text().strip()
            if not phone_number:
                QMessageBox.warning(self, "警告", "電話番号を入力してください。")
                return
                
            # TODO: 電話番号検索の実装
        except Exception as e:
            logging.error(f"電話番号検索中にエラーが発生しました: {str(e)}")
            QMessageBox.critical(self, "エラー", "検索中にエラーが発生しました。")
            
    def copy_to_clipboard(self, text: str):
        """クリップボードにコピー"""
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage(f"クリップボードにコピーしました: {text}")
        
    def load_settings(self):
        """設定の読み込み"""
        try:
            # 設定の適用
            self.settings = settings
        except Exception as e:
            logging.error(f"設定の読み込み中にエラーが発生しました: {str(e)}")
            QMessageBox.warning(self, "警告", "設定の読み込みに失敗しました。") 