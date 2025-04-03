"""
シンプルモードのメインウィンドウを提供するモジュール

このモジュールは、シンプルモード（通常モード）のメインウィンドウを実装します。
主な機能：
- 地域検索
- 電話番号の入力と検索
- 検索結果の表示
- 設定の管理
- ログ出力

制限事項：
- 地域検索は最大100件まで表示
- 電話番号は10桁まで入力可能
- 検索結果は最大100件まで表示
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
    QDialog, QFormLayout, QSpinBox, QCheckBox, QDialogButtonBox
)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
from PySide6.QtGui import QAction, QIcon, QFont

from .base_window import BaseWindow
from .custom_widgets import CustomComboBox
from ..services.area_search_service import AreaSearchService
from ..services.cti_service import CTIService
from ..utils.logger import setup_logger
from ..utils.settings import load_settings, save_settings
from version import check_version

class SimpleModeWindow(BaseWindow):
    """シンプルモードのメインウィンドウクラス"""
    
    def __init__(self):
        """ウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("電話番号検索ツール - シンプルモード")
        self.setup_ui()
        self.setup_connections()
        self.load_settings()
        
    def setup_ui(self):
        """UIの初期設定"""
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # 地域検索エリア
        area_search_group = self.create_area_search_group()
        layout.addWidget(area_search_group)
        
        # 電話番号検索エリア
        phone_search_group = self.create_phone_search_group()
        layout.addWidget(phone_search_group)
        
        # 検索結果表示エリア
        result_group = self.create_result_group()
        layout.addWidget(result_group)
        
        # ステータスバーの設定
        self.statusBar().showMessage("準備完了")
        
    def create_area_search_group(self) -> QWidget:
        """地域検索グループの作成"""
        group = QWidget()
        layout = QHBoxLayout(group)
        
        # 都道府県選択
        self.prefecture_combo = CustomComboBox()
        self.prefecture_combo.addItem("都道府県を選択", "")
        layout.addWidget(self.prefecture_combo)
        
        # 市区町村選択
        self.city_combo = CustomComboBox()
        self.city_combo.addItem("市区町村を選択", "")
        layout.addWidget(self.city_combo)
        
        # 検索ボタン
        self.area_search_button = QPushButton("地域検索")
        layout.addWidget(self.area_search_button)
        
        return group
        
    def create_phone_search_group(self) -> QWidget:
        """電話番号検索グループの作成"""
        group = QWidget()
        layout = QHBoxLayout(group)
        
        # 電話番号入力
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("電話番号を入力")
        self.phone_input.setMaxLength(10)
        layout.addWidget(self.phone_input)
        
        # 検索ボタン
        self.phone_search_button = QPushButton("電話番号検索")
        layout.addWidget(self.phone_search_button)
        
        return group
        
    def create_result_group(self) -> QWidget:
        """検索結果表示グループの作成"""
        group = QWidget()
        layout = QVBoxLayout(group)
        
        # 結果テーブル
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels([
            "電話番号", "地域", "事業者", "備考"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.result_table)
        
        return group
        
    def setup_connections(self):
        """シグナルとスロットの接続"""
        # 地域検索関連
        self.prefecture_combo.currentIndexChanged.connect(self.on_prefecture_changed)
        self.area_search_button.clicked.connect(self.on_area_search_clicked)
        
        # 電話番号検索関連
        self.phone_search_button.clicked.connect(self.on_phone_search_clicked)
        
        # 結果テーブル関連
        self.result_table.itemDoubleClicked.connect(self.on_result_double_clicked)
        
    def on_prefecture_changed(self, index: int):
        """都道府県が変更されたときの処理"""
        prefecture = self.prefecture_combo.currentData()
        if prefecture:
            self.update_city_combo(prefecture)
            
    def on_area_search_clicked(self):
        """地域検索ボタンがクリックされたときの処理"""
        prefecture = self.prefecture_combo.currentData()
        city = self.city_combo.currentData()
        
        if not prefecture or not city:
            QMessageBox.warning(self, "警告", "都道府県と市区町村を選択してください。")
            return
            
        self.search_by_area(prefecture, city)
        
    def on_phone_search_clicked(self):
        """電話番号検索ボタンがクリックされたときの処理"""
        phone_number = self.phone_input.text().strip()
        if not phone_number:
            QMessageBox.warning(self, "警告", "電話番号を入力してください。")
            return
            
        self.search_by_phone(phone_number)
        
    def on_result_double_clicked(self, item: QTableWidgetItem):
        """検索結果がダブルクリックされたときの処理"""
        row = item.row()
        phone_number = self.result_table.item(row, 0).text()
        self.copy_to_clipboard(phone_number)
        
    def update_city_combo(self, prefecture: str):
        """市区町村コンボボックスの更新"""
        self.city_combo.clear()
        self.city_combo.addItem("市区町村を選択", "")
        
        # TODO: 市区町村データの取得と設定
        
    def search_by_area(self, prefecture: str, city: str):
        """地域による検索"""
        try:
            # TODO: 地域検索の実装
            pass
        except Exception as e:
            logging.error(f"地域検索中にエラーが発生しました: {str(e)}")
            QMessageBox.critical(self, "エラー", "検索中にエラーが発生しました。")
            
    def search_by_phone(self, phone_number: str):
        """電話番号による検索"""
        try:
            # TODO: 電話番号検索の実装
            pass
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
            settings = load_settings()
            # TODO: 設定の適用
        except Exception as e:
            logging.error(f"設定の読み込み中にエラーが発生しました: {str(e)}")
            QMessageBox.warning(self, "警告", "設定の読み込みに失敗しました。") 