"""
メインウィンドウ

このモジュールは、アプリケーションのメインウィンドウを
提供します。
"""

import datetime
import logging
import json
import os
import re
import time
import requests
from urllib.parse import quote
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QComboBox, QPushButton,
                              QTextEdit, QGroupBox, QMessageBox, QScrollArea,
                              QApplication, QToolTip, QSplitter, QMenuBar, QMenu,
                              QSizePolicy, QProgressBar, QFrame)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread, QSize
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices

from version import VERSION, GITHUB_OWNER, GITHUB_REPO, APP_NAME

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from ui.main_window_functions import MainWindowFunctions
from utils.string_utils import validate_name, validate_furigana, convert_to_half_width_except_space
from utils.furigana_utils import convert_to_furigana
from services.oneclick import OneClickService
from services.phone_button_monitor import PhoneButtonMonitor
from .update_dialog import UpdateDialog


class CustomComboBox(QComboBox):
    """スクロールでの値変更を防止するカスタムコンボボックス"""
    def wheelEvent(self, event):
        """ホイールイベントを無視"""
        event.ignore()


class MainWindow(QMainWindow, MainWindowFunctions):
    """メインウィンドウクラス"""
    
    def __init__(self):
        """メインウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(600, 400)
        
        # 現在の日付を取得して受付日に設定
        current_date = datetime.datetime.now().strftime("%Y/%m/%d")
        
        # フォントの設定
        self.setStyleSheet("""
            * {
                font-family: 'Roboto', 'Noto Sans JP', sans-serif;
            }
            QMainWindow {
                background-color: #F5F5F5;
            }
            QLineEdit {
                padding: 4px 8px;
                border: 1px solid #E0E0E0;
                border-radius: 2px;
                background-color: white;
                font-size: 12px;
                min-height: 24px;
            }
            QLineEdit:focus {
                border: 2px solid #1565C0;
                background-color: #FFFFFF;
            }
            QComboBox {
                padding: 4px 8px;
                border: 1px solid #E0E0E0;
                border-radius: 2px;
                background-color: white;
                font-size: 12px;
                min-height: 24px;
            }
            QComboBox:focus {
                border: 2px solid #1565C0;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(icons/arrow_down.png);
                width: 12px;
                height: 12px;
            }
            QGroupBox {
                border: none;
                background-color: white;
                border-radius: 4px;
                margin-top: 1em;
                font-size: 13px;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
                color: #1565C0;
                font-weight: bold;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #F5F5F5;
                width: 8px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background: #BDBDBD;
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
            QLabel {
                font-size: 12px;
                color: #424242;
                margin-bottom: 2px;
            }
            QWidget#mainWidget {
                background-color: #F5F5F5;
            }
        """)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(16, 8, 16, 16)
        main_layout.setSpacing(8)
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # フォーマットテンプレートの読み込み
        self.load_settings()
        
        # トップバーの作成
        self.create_top_bar(main_layout)
        
        # スプリッターの作成
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #cccccc;
                width: 2px;
            }
            QSplitter::handle:hover {
                background-color: #999999;
            }
            QSplitter::handle:pressed {
                background-color: #666666;
            }
        """)
        splitter.setHandleWidth(2)  # スプリッターハンドルの幅を設定
        
        # 入力フォームエリア（左側）をスクロール可能に
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        self.create_input_form(form_layout)
        
        # スクロールエリアの作成
        scroll_area = QScrollArea()
        scroll_area.setWidget(form_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # スクロールエリアのスタイルシートを更新
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #F0F0F0;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #CCCCCC;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
            }
        """)
        
        # プレビューエリア（右側）
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.create_preview_area(preview_layout)
        
        # スプリッターにウィジェットを追加
        splitter.addWidget(scroll_area)
        splitter.addWidget(preview_widget)
        
        # 初期のサイズ比率を設定（7:3）
        splitter.setSizes([700, 300])
        
        # スプリッターをメインレイアウトに追加
        main_layout.addWidget(splitter)
        
        # シグナルの設定
        self.setup_signals()
        
        # Google Sheetsの設定
        self.setup_google_sheets()
        
        # フォントサイズの適用
        self.apply_font_size()
        
        # CTI連携サービスの初期化
        self.cti_service = OneClickService()
        
        # 電話ボタン監視の初期化と開始
        self.phone_monitor = PhoneButtonMonitor(self.fetch_cti_data)
        self.phone_monitor.start_monitoring()
        
        # カウントダウン表示用のラベル
        self.countdown_label = QLabel()
        self.countdown_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #E74C3C;
                padding: 5px;
                border: 1px solid #E74C3C;
                border-radius: 4px;
                background-color: #FFEBEE;
            }
        """)
        self.countdown_label.hide()
        main_layout.addWidget(self.countdown_label)
        
        # カウントダウン更新用のタイマー
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        
        self.init_menu()
        
        # 起動時にアップデートをチェック
        QTimer.singleShot(0, self.check_for_updates)
    
    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        top_bar = QWidget()
        top_bar.setFixedHeight(48)
        top_bar.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
            }
        """)
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(8, 0, 8, 0)
        top_bar_layout.setSpacing(8)
        
        # ボタンのベーススタイル
        button_style = """
            QPushButton {
                background-color: #1565C0;
                color: white;
                border: none;
                border-radius: 2px;
                padding: 4px 12px;
                font-size: 12px;
                font-weight: 500;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #0D47A1;
            }
            QPushButton:pressed {
                background-color: #0A3D91;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """
        
        # 各ボタンの作成と設定
        self.oneclick_btn = QPushButton("顧客情報取得")
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("営コメ作成")
        self.screenshot_btn = QPushButton("提供判定のスクリーンショット確認")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
        self.settings_btn = QPushButton("設定")
        
        # ボタンのスタイル設定
        buttons = [self.oneclick_btn, self.clear_btn, self.cti_copy_btn,
                  self.screenshot_btn, self.spreadsheet_btn, self.settings_btn]
        
        for btn in buttons:
            btn.setStyleSheet(button_style)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            top_bar_layout.addWidget(btn)
        
        parent_layout.addWidget(top_bar)
        
        # ボタンの接続
        self.oneclick_btn.clicked.connect(self.fetch_cti_data)
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.copy_cti_to_clipboard)
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        self.spreadsheet_btn.setEnabled(False)  # 未実装
        self.settings_btn.clicked.connect(self.show_settings)
    
    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        form_group = QGroupBox("受注者入力項目")
        form_group.setStyleSheet("""
            QGroupBox {
                background-color: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                margin-top: 1.5em;
                font-size: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
                color: #424242;
                background-color: white;
            }
        """)
        
        form_layout = QVBoxLayout(form_group)
        form_layout.setSpacing(4)
        form_layout.setContentsMargins(12, 16, 12, 12)
        
        # 共通のスタイル設定
        input_style = """
            QLineEdit, QComboBox {
                min-height: 30px;
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 2px;
                background-color: white;
                font-size: 12px;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 1px solid #1565C0;
            }
            QLineEdit:hover, QComboBox:hover {
                border: 1px solid #1565C0;
            }
            QLineEdit:disabled, QComboBox:disabled {
                background-color: #f5f5f5;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(down_arrow.png);
                width: 12px;
                height: 12px;
            }
        """
        
        # グループボックスのスタイル
        group_style = """
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                margin-top: 1em;
                padding-top: 1em;
                font-weight: bold;
                background-color: #f8f8f8;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
                color: #333;
            }
        """
        
        # ボタンのスタイル
        button_style = """
            QPushButton {
                min-height: 30px;
                padding: 5px 15px;
                border: 1px solid #1565C0;
                border-radius: 2px;
                background-color: #1565C0;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0D47A1;
            }
            QPushButton:pressed {
                background-color: #0A3D91;
            }
            QPushButton:disabled {
                background-color: #ccc;
                border-color: #ccc;
            }
        """
        
        # 緑色のボタン（地図を開くボタン）のスタイル
        green_button_style = """
            QPushButton {
                min-height: 30px;
                padding: 5px 15px;
                border: 1px solid #2E7D32;
                border-radius: 2px;
                background-color: #2E7D32;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1B5E20;
            }
            QPushButton:pressed {
                background-color: #0A280E;
            }
            QPushButton:disabled {
                background-color: #ccc;
                border-color: #ccc;
            }
        """
        
        # プレビューエリアのスタイル
        preview_style = """
            QTextEdit {
                min-height: 200px;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 2px;
                background-color: white;
                font-family: "MS Gothic", "Yu Gothic", sans-serif;
                font-size: 12px;
                line-height: 1.5;
            }
            QTextEdit:focus {
                border: 1px solid #1565C0;
            }
        """
        
        # スクロールバーのスタイル
        scrollbar_style = """
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 10px;
                margin: 0px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical {
                background: #ccc;
                min-height: 20px;
                border-radius: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background: #999;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """
        
        # 各入力フィールドの作成
        self.operator_input = QLineEdit()  # 対応者名
        self.available_time_input = QLineEdit()  # 出やすい時間帯
        self.contractor_input = QLineEdit()  # 契約者名
        self.furigana_mode_combo = CustomComboBox()  # カスタムコンボボックスを使用
        self.furigana_mode_combo.addItems(["自動", "手動"])
        self.furigana_mode_combo.setStyleSheet(input_style)
        self.furigana_input = QLineEdit()
        self.postal_code_input = QLineEdit()  # 郵便番号
        self.address_input = QLineEdit()  # 住所
        self.address_furigana_input = QLineEdit()  # 住所フリガナ
        self.list_name_input = QLineEdit()  # リスト名
        self.list_furigana_input = QLineEdit()  # リストフリガナ
        self.list_phone_input = QLineEdit()  # リスト電話番号
        self.list_postal_code_input = QLineEdit()  # リスト郵便番号
        self.list_address_input = QLineEdit()  # リスト住所
        self.list_address_furigana_input = QLineEdit()  # リスト住所フリガナ
        self.order_person_input = QLineEdit()  # 受注者名
        self.fee_input = QLineEdit("2500円")  # 料金認識（初期値を2500円に設定）
        self.other_number_input = QLineEdit("なし")  # 他番号
        self.phone_device_input = QLineEdit("プッシュホン")  # 電話機
        self.forbidden_line_input = QLineEdit("なし")  # 禁止回線
        self.relationship_input = QLineEdit()  # 続柄
        self.nd_input = QLineEdit()  # ND
        self.order_date_input = QLineEdit()  # 受付日
        self.order_date_input.setText(datetime.datetime.now().strftime("%Y/%m/%d"))  # 現在の日付を設定
        
        # コンボボックスの作成（すべてカスタムコンボボックスを使用）
        self.current_line_combo = CustomComboBox()
        self.current_line_combo.addItems(["なし", "あり"])
        
        self.judgment_combo = CustomComboBox()
        self.judgment_combo.addItems(["", "○", "×"])
        
        self.net_usage_combo = CustomComboBox()
        self.net_usage_combo.addItems(["", "利用", "未利用"])
        
        self.family_approval_combo = CustomComboBox()
        self.family_approval_combo.addItems(["ok", "ng"])
        
        # プログレスバーの作成
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        # エリア検索ボタンの作成
        self.area_search_btn = QPushButton("提供エリア検索")
        self.area_search_btn.setStyleSheet(button_style)
        
        # 地図ボタンの作成
        self.map_btn = QPushButton("地図を開く")
        self.map_btn.setStyleSheet(green_button_style)
        
        # エリア結果ラベルの作成
        self.area_result_label = QLabel("提供エリア: 未検索")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #95a5a6;
                border-radius: 4px;
                background-color: #f8f9fa;
                color: #95a5a6;
            }
        """)
        
        # フィールドの配置（元の順番に戻す）
        fields = [
            # 基本情報
            ("対応者名", self.operator_input),
            ("出やすい時間帯", self.available_time_input),
            ("契約者名", self.contractor_input),
            ("フリガナ", self.furigana_input),
            ("生年月日", self.create_date_input()),
            
            # 区切り線1
            ("", QFrame()),
            
            # 住所情報
            ("郵便番号", self.postal_code_input),
            ("住所", self.address_input),
            ("住所フリガナ", self.address_furigana_input),
            
            # 区切り線2
            ("", QFrame()),
            
            # リスト情報
            ("リスト名", self.list_name_input),
            ("リストフリガナ", self.list_furigana_input),
            ("リスト電話番号", self.list_phone_input),
            ("リスト郵便番号", self.list_postal_code_input),
            ("リスト住所", self.list_address_input),
            ("リスト住所フリガナ", self.list_address_furigana_input),
            
            # エリア検索関連のウィジェット
            ("", self.create_area_search_widget()),
            
            # 区切り線3
            ("", QFrame()),
            
            # 受注情報
            ("受注者名", self.order_person_input),
            ("料金認識", self.fee_input),
            ("他番号", self.other_number_input),
            ("電話機", self.phone_device_input),
            ("禁止回線", self.forbidden_line_input),
            ("続柄", self.relationship_input),
            ("ND", self.nd_input),
            ("受付日", self.order_date_input),
            
            # 区切り線4
            ("", QFrame()),
            
            # 判定情報
            ("現回線", self.current_line_combo),
            ("判定", self.judgment_combo),
            ("ネット利用", self.net_usage_combo),
            ("家族了承", self.family_approval_combo)
        ]
        
        for label_text, widget in fields:
            if isinstance(widget, QFrame):
                # 区切り線のスタイル設定
                widget.setFrameShape(QFrame.Shape.HLine)
                widget.setFrameShadow(QFrame.Shadow.Sunken)
                widget.setStyleSheet("""
                    QFrame {
                        background-color: #E0E0E0;
                        margin: 16px 0;
                        height: 2px;
                    }
                """)
                form_layout.addWidget(widget)
            else:
                field_container = QWidget()
                field_layout = QVBoxLayout(field_container)
                field_layout.setSpacing(2)
                field_layout.setContentsMargins(0, 0, 0, 4)
                
                if label_text:  # ラベルテキストがある場合のみラベルを追加
                    label = QLabel(label_text)
                    field_layout.addWidget(label)
                
                if isinstance(widget, QLineEdit):
                    widget.setStyleSheet(input_style)
                elif isinstance(widget, QComboBox):
                    widget.setStyleSheet(input_style)
                
                field_layout.addWidget(widget)
                form_layout.addWidget(field_container)
        
        parent_layout.addWidget(form_group)
    
    def create_date_input(self):
        """生年月日入力用のウィジェットを作成"""
        date_widget = QWidget()
        date_layout = QHBoxLayout(date_widget)
        date_layout.setSpacing(8)
        date_layout.setContentsMargins(0, 0, 0, 0)
        
        # スタイル設定
        input_style = """
            QLineEdit {
                padding: 4px 8px;
                border: 1px solid #E0E0E0;
                border-radius: 2px;
                background-color: white;
                font-size: 12px;
                min-height: 24px;
                min-width: 60px;
            }
            QLineEdit:focus {
                border: 2px solid #1565C0;
            }
        """
        
        combo_style = """
            QComboBox {
                padding: 4px 8px;
                border: 1px solid #E0E0E0;
                border-radius: 2px;
                background-color: white;
                font-size: 12px;
                min-height: 24px;
                min-width: 80px;
            }
            QComboBox:focus {
                border: 2px solid #1565C0;
            }
        """
        
        # 年号選択（カスタムコンボボックスを使用）
        self.era_combo = CustomComboBox()
        self.era_combo.addItems(["昭和", "平成", "令和"])
        self.era_combo.setStyleSheet(combo_style)
        
        # 年入力（手動入力可能なQLineEdit）
        self.year_input = QLineEdit()
        self.year_input.setStyleSheet(input_style)
        self.year_input.setMaxLength(4)  # 最大4文字まで
        self.year_input.setPlaceholderText("年")
        
        # 月入力（手動入力可能なQLineEdit）
        self.month_input = QLineEdit()
        self.month_input.setStyleSheet(input_style)
        self.month_input.setMaxLength(2)  # 最大2文字まで
        self.month_input.setPlaceholderText("月")
        
        # 日入力（手動入力可能なQLineEdit）
        self.day_input = QLineEdit()
        self.day_input.setStyleSheet(input_style)
        self.day_input.setMaxLength(2)  # 最大2文字まで
        self.day_input.setPlaceholderText("日")
        
        # ラベルのスタイル
        label_style = "QLabel { color: #424242; font-size: 12px; }"
        
        # 各要素をレイアウトに追加
        date_layout.addWidget(self.era_combo)
        date_layout.addWidget(QLabel("年", styleSheet=label_style))
        date_layout.addWidget(self.year_input)
        date_layout.addWidget(QLabel("月", styleSheet=label_style))
        date_layout.addWidget(self.month_input)
        date_layout.addWidget(QLabel("日", styleSheet=label_style))
        date_layout.addWidget(self.day_input)
        
        # 数字のみ入力可能なバリデーターを設定
        validator = QIntValidator(1, 9999)  # 年のバリデーターを4桁に変更
        self.year_input.setValidator(validator)
        self.month_input.setValidator(QIntValidator(1, 12))
        self.day_input.setValidator(QIntValidator(1, 31))
        
        return date_widget
    
    def create_preview_area(self, parent_layout):
        """プレビューエリアを作成"""
        preview_widget = QWidget()
        preview_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
            }
        """)
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setSpacing(8)
        preview_layout.setContentsMargins(16, 16, 16, 16)
        
        # プレビューエリアのスタイル
        preview_style = """
            QTextEdit {
                min-height: 200px;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 2px;
                background-color: white;
                font-family: "MS Gothic", "Yu Gothic", sans-serif;
                font-size: 12px;
                line-height: 1.5;
            }
            QTextEdit:focus {
                border: 1px solid #1565C0;
            }
        """
        
        # CTIフォーマットプレビュー
        cti_label = QLabel("CTIフォーマットプレビュー")
        cti_label.setStyleSheet("""
            QLabel {
                color: #1976D2;
                font-size: 13px;
                font-weight: bold;
                padding: 4px 0;
            }
        """)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet(preview_style)
        
        preview_layout.addWidget(cti_label)
        preview_layout.addWidget(self.preview_text)
        parent_layout.addWidget(preview_widget)
    
    def setup_signals(self):
        """シグナルの設定"""
        # 自動フォーマット用のシグナル
        self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
        self.postal_code_input.textChanged.connect(self.format_postal_code)
        self.postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.list_postal_code_input.textChanged.connect(self.format_postal_code)
        self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.address_input.textChanged.connect(self.convert_to_half_width)
        self.address_input.textChanged.connect(self.auto_generate_address_furigana)
        self.list_address_input.textChanged.connect(self.convert_to_half_width)
        self.list_address_input.textChanged.connect(self.auto_generate_list_address_furigana)
        self.era_combo.currentTextChanged.connect(self.update_year_combo)
        
        # フリガナモードの切り替えシグナル
        self.furigana_mode_combo.currentTextChanged.connect(self.on_furigana_mode_changed)
        
        # 名前とフリガナのバリデーション用のシグナル
        self.contractor_input.textChanged.connect(self.validate_contractor_name)
        self.furigana_input.textChanged.connect(self.validate_furigana_input)
        self.list_name_input.textChanged.connect(self.validate_list_name)
        self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
        
        # フリガナ自動変換のシグナル（初期状態で接続）
        self.contractor_input.textChanged.connect(self.auto_generate_furigana)
        self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
        self.address_input.textChanged.connect(self.auto_generate_address_furigana)
        self.list_address_input.textChanged.connect(self.auto_generate_list_address_furigana)
        
        # 入力時に背景色をリセットするシグナル
        self.operator_input.textChanged.connect(self.reset_background_color)
        self.available_time_input.textChanged.connect(self.reset_background_color)
        self.contractor_input.textChanged.connect(self.reset_background_color)
        self.furigana_input.textChanged.connect(self.reset_background_color)
        self.postal_code_input.textChanged.connect(self.reset_background_color)
        self.address_input.textChanged.connect(self.reset_background_color)
        self.list_name_input.textChanged.connect(self.reset_background_color)
        self.list_furigana_input.textChanged.connect(self.reset_background_color)
        self.list_phone_input.textChanged.connect(self.reset_background_color)
        self.list_postal_code_input.textChanged.connect(self.reset_background_color)
        self.list_address_input.textChanged.connect(self.reset_background_color)
        self.order_person_input.textChanged.connect(self.reset_background_color)
        self.fee_input.textChanged.connect(self.reset_background_color)
        self.relationship_input.textChanged.connect(self.reset_background_color)
        self.nd_input.textChanged.connect(self.reset_background_color)
        self.order_date_input.textChanged.connect(self.reset_background_color)
        
        # ボタンのシグナル接続
        self.area_search_btn.clicked.connect(self.search_service_area)
        self.map_btn.clicked.connect(self.open_street_view)
        
        # プレビュー更新のシグナル
        self.contractor_input.textChanged.connect(self.update_preview)
        self.furigana_input.textChanged.connect(self.update_preview)
        self.postal_code_input.textChanged.connect(self.update_preview)
        self.address_input.textChanged.connect(self.update_preview)
        self.list_name_input.textChanged.connect(self.update_preview)
        self.list_furigana_input.textChanged.connect(self.update_preview)
        self.list_phone_input.textChanged.connect(self.update_preview)
        self.list_postal_code_input.textChanged.connect(self.update_preview)
        self.list_address_input.textChanged.connect(self.update_preview)
        self.order_person_input.textChanged.connect(self.update_preview)
        self.fee_input.textChanged.connect(self.update_preview)
        self.other_number_input.textChanged.connect(self.update_preview)
        self.phone_device_input.textChanged.connect(self.update_preview)
        self.forbidden_line_input.textChanged.connect(self.update_preview)
        self.relationship_input.textChanged.connect(self.update_preview)
        self.nd_input.textChanged.connect(self.update_preview)
        self.order_date_input.textChanged.connect(self.update_preview)
        self.current_line_combo.currentTextChanged.connect(self.update_preview)
        self.judgment_combo.currentTextChanged.connect(self.update_preview)
        self.net_usage_combo.currentTextChanged.connect(self.update_preview)
        self.family_approval_combo.currentTextChanged.connect(self.update_preview)
    
    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            try:
                # ダイアログがOKで閉じられた場合、設定を再読み込み
                self.load_settings()
                # フォントサイズを適用
                self.apply_font_size()
                # 電話ボタン監視の設定を更新
                self.phone_monitor.update_settings()
                # ウィンドウ全体を更新
                self.update()
                logging.info("設定を更新しました")
            except Exception as e:
                logging.error(f"設定の更新中にエラーが発生しました: {e}")
                QMessageBox.critical(self, "エラー", f"設定の更新中にエラーが発生しました: {e}")
            
    def update_countdown(self):
        """カウントダウン表示を更新"""
        try:
            if hasattr(self.phone_monitor, 'is_counting_down') and self.phone_monitor.is_counting_down:
                remaining_time = self.phone_monitor.delay_seconds - (time.time() - self.phone_monitor.countdown_start_time)
                if remaining_time > 0:
                    self.countdown_label.setText(f"情報取得まで: {int(remaining_time)}秒")
                    self.countdown_label.show()
                else:
                    self.countdown_label.hide()
                    self.countdown_timer.stop()
            else:
                self.countdown_label.hide()
                self.countdown_timer.stop()
        except Exception as e:
            logging.error(f"カウントダウン表示の更新中にエラー: {e}")
            self.countdown_label.hide()
            self.countdown_timer.stop()
            
    def update_form_with_data(self, data):
        """
        CTIデータをフォームに反映します
        
        Args:
            data: CTIから取得したデータ
        """
        try:
            # 顧客名
            if data.customer_name:
                # 半角スペースを全角スペースに変換
                converted_customer_name = data.customer_name.replace(' ', '　')
                converted_customer_name = convert_to_half_width_except_space(converted_customer_name)
                self.list_name_input.setText(converted_customer_name)
                self.contractor_input.setText(converted_customer_name)
            
            # 住所
            if data.address:
                converted_address = convert_to_half_width_except_space(data.address)
                self.address_input.setText(converted_address)
                self.list_address_input.setText(converted_address)
            
            # 電話番号
            if data.phone:
                converted_phone = convert_to_half_width_except_space(data.phone)
                self.list_phone_input.setText(converted_phone)
            
            # 郵便番号
            if data.postal_code:
                converted_postal_code = convert_to_half_width_except_space(data.postal_code)
                self.postal_code_input.setText(converted_postal_code)
                self.list_postal_code_input.setText(converted_postal_code)
                
            # プレビューを更新しない（営業コメントを自動作成しない）
            # self.update_preview()
            
            # 成功メッセージ
            self.statusBar().showMessage("データを取得しました", 5000)
            
        except Exception as e:
            logging.error(f"フォーム更新中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"フォームの更新中にエラーが発生しました: {e}")
            
    def fetch_cti_data(self):
        """CTIデータを取得"""
        try:
            # カウントダウン表示を非表示
            self.countdown_label.hide()
            self.countdown_timer.stop()
            
            # CTIデータの取得処理
            data = self.cti_service.get_all_fields_data()
            if data:
                # メインスレッドでUIを更新
                QApplication.instance().postEvent(self, QEvent(QEvent.User))
                self.update_form_with_data(data)
                logging.info("CTIデータの取得に成功しました")
            else:
                logging.warning("CTIデータの取得に失敗しました")
        except Exception as e:
            logging.error(f"CTIデータの取得中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"CTIデータの取得中にエラーが発生しました: {e}")
            
    def event(self, event):
        """イベントハンドラ"""
        if event.type() == QEvent.User:
            # メインスレッドでUIを更新
            self.update_form_with_data(self.cti_service.get_all_fields_data())
            return True
        elif event.type() == QEvent.User + 1:
            # メインスレッドでプレビューを更新
            try:
                preview_text = self.generate_preview_text()
                if preview_text:
                    self.preview_text.setText(preview_text)
            except Exception as e:
                logging.error(f"プレビュー更新中にエラー: {e}")
                self.preview_text.setText("プレビューの更新に失敗しました")
            return True
        return super().event(event)

    def validate_contractor_name(self, text):
        """
        契約者名の入力を検証します。
        全角文字のみを許可し、半角文字が含まれている場合は警告を表示します。
        
        Args:
            text (str): 入力されたテキスト
        """
        import unicodedata
        
        # 空文字列の場合は検証をスキップ
        if not text:
            return
        
        # 半角文字が含まれているかチェック
        has_half_width = any(unicodedata.east_asian_width(char) in ['Na', 'H'] for char in text)
        
        if has_half_width:
            self.statusBar().showMessage("契約者名は全角文字で入力してください", 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_furigana_input(self, text):
        """
        フリガナの入力を検証します。
        カタカナと長音記号のみを許可し、それ以外の文字が含まれている場合は警告を表示します。
        
        Args:
            text (str): 入力されたテキスト
        """
        import re
        
        # 空文字列の場合は検証をスキップ
        if not text:
            return
        
        # カタカナと長音記号のみを許可する正規表現パターン
        katakana_pattern = r'^[ァ-ヶーヽヾ]+$'
        
        if not re.match(katakana_pattern, text):
            self.statusBar().showMessage("フリガナは全角カタカナで入力してください", 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_list_name(self, text):
        """
        リスト名の入力を検証します。
        半角英数字とハイフンのみを許可し、それ以外の文字が含まれている場合は警告を表示します。
        
        Args:
            text (str): 入力されたテキスト
        """
        import re
        
        # 空文字列の場合は検証をスキップ
        if not text:
            return
        
        # 半角英数字とハイフンのみを許可する正規表現パターン
        pattern = r'^[A-Za-z0-9\-_]+$'
        
        if not re.match(pattern, text):
            self.statusBar().showMessage("リスト名は半角英数字とハイフンのみ使用できます", 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_list_furigana(self):
        """リストフリガナのバリデーション"""
        text = self.list_furigana_input.text()
        if not validate_furigana(text):
            # 背景色変更を削除
            QToolTip.showText(
                self.list_furigana_input.mapToGlobal(QPoint(0, 0)),
                "フリガナに数字や不適切な文字を含めることはできません",
                self.list_furigana_input
            )
        else:
            # 背景色変更を削除
            QToolTip.hideText()

    def reset_background_color(self):
        """
        フィールドの背景色をリセットする
        
        入力の有無に関わらず、対応する未入力警告の背景色をリセットします。
        """
        sender = self.sender()
        if sender:
            sender.setStyleSheet("")

    def closeEvent(self, event):
        """ウィンドウを閉じる際の処理"""
        # 電話ボタン監視を停止
        if hasattr(self, 'phone_monitor'):
            self.phone_monitor.stop_monitoring()
        event.accept()

    def update_preview(self):
        """プレビューを更新"""
        try:
            # メインスレッドでプレビューを更新
            QApplication.instance().postEvent(self, QEvent(QEvent.User + 1))
        except Exception as e:
            logging.error(f"プレビュー更新中にエラー: {e}")

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        self.operator_input.clear()
        self.available_time_input.clear()  # 出やすい時間帯をクリア
        self.contractor_input.clear()
        self.furigana_input.clear()
        self.postal_code_input.clear()
        self.address_input.clear()
        self.address_furigana_input.clear()  # 住所フリガナをクリア
        self.list_name_input.clear()
        self.list_furigana_input.clear()
        self.list_phone_input.clear()
        self.list_postal_code_input.clear()
        self.list_address_input.clear()
        self.list_address_furigana_input.clear()  # リスト住所フリガナをクリア
        # 受注者名はクリアしない（保持する）
        # self.order_person_input.clear()
        # 料金認識はクリアしない（保持する）
        # self.fee_input.clear()
        
        # 他番号、電話機、禁止回線には初期値を設定
        self.other_number_input.setText("なし")
        self.phone_device_input.setText("プッシュホン")
        self.forbidden_line_input.setText("なし")
        
        self.relationship_input.clear()
        # コンボボックスをデフォルト値に
        self.era_combo.setCurrentIndex(0)
        self.year_input.clear()
        self.month_input.clear()
        self.day_input.clear()
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)  # okがインデックス0になる
        # 結果ラベルをクリア
        self.area_result_label.setText("提供エリア: 未検索")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #95a5a6;
                border-radius: 4px;
                background-color: #f8f9fa;
                color: #95a5a6;
            }
        """)
        # スクリーンショットボタンをクリア
        self.update_screenshot_button()
        # プレビューもクリア
        self.preview_text.clear()

    def init_menu(self):
        """メニューバーの初期化"""
        menubar = self.menuBar()
        
        # ファイルメニュー
        file_menu = menubar.addMenu("ファイル")
        
        # 終了
        exit_action = file_menu.addAction("終了")
        exit_action.triggered.connect(self.close)
        
        # ヘルプメニュー
        help_menu = menubar.addMenu("ヘルプ")
        
        # アップデートの確認
        update_action = help_menu.addAction("アップデートの確認")
        update_action.triggered.connect(self.show_update_dialog)
        
        # バージョン情報
        about_action = help_menu.addAction("バージョン情報")
        about_action.triggered.connect(self.show_about_dialog)
        
        # バージョン表示ラベル
        version_label = QLabel(f"v{VERSION}")
        version_label.setStyleSheet("""
            QLabel {
                color: #95A5A6;
                font-size: 12px;
                padding: 2px 6px;
                margin-right: 5px;
            }
        """)
        menubar.setCornerWidget(version_label, Qt.TopRightCorner)
        
    def show_update_dialog(self):
        """アップデート設定ダイアログを表示する"""
        dialog = UpdateDialog(self)
        dialog.settings_file = self.settings_file  # 設定ファイルのパスを渡す
        dialog.exec()
        
    def show_about_dialog(self):
        """バージョン情報ダイアログを表示する"""
        msg = f"{APP_NAME} v{VERSION}\n\n"
        msg += "ライセンス: MIT License"
        QMessageBox.information(self, "バージョン情報", msg)

    def check_for_updates(self):
        """アップデートをチェック"""
        try:
            # GitHubのAPIを使用して最新リリースを取得
            url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            response = requests.get(url)
            response.raise_for_status()
            latest_release = response.json()
            
            latest_version = latest_release["tag_name"].lstrip("v")
            current_version = VERSION
            
            if latest_version > current_version:
                # 新しいバージョンが利用可能
                msg = f"新しいバージョン v{latest_version} が利用可能です。\n"
                msg += f"現在のバージョン: v{current_version}\n\n"
                msg += "更新しますか？"
                
                reply = QMessageBox.question(self, "アップデート", msg,
                                          QMessageBox.StandardButton.Yes |
                                          QMessageBox.StandardButton.No)
                
                if reply == QMessageBox.StandardButton.Yes:
                    # アップデートダイアログを作成して更新を実行
                    dialog = UpdateDialog(self)
                    dialog.settings_file = self.settings_file
                    dialog.download_and_apply_update(latest_release)
        except Exception as e:
            logging.error(f"アップデートチェック中にエラー: {e}")

    def show_screenshot(self):
        """スクリーンショットを表示する"""
        try:
            if hasattr(self, 'screenshot_path') and self.screenshot_path:
                screenshot_path = self.screenshot_path
            else:
                screenshot_path = "debug_screenshot.png"
            
            if not os.path.exists(screenshot_path):
                QMessageBox.warning(
                    self,
                    "エラー",
                    "スクリーンショットファイルが見つかりません。"
                )
                return
            
            # QPixmapを使用して画像を表示
            from PySide6.QtGui import QPixmap
            from PySide6.QtWidgets import QLabel, QDialog, QVBoxLayout, QScrollArea
            from PySide6.QtCore import Qt
            
            dialog = QDialog(self)
            dialog.setWindowTitle("スクリーンショット - 提供判定結果")
            dialog.setMinimumSize(800, 600)
            layout = QVBoxLayout(dialog)
            
            # スクロールエリアを作成
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            
            # ラベルを作成してピクスマップを設定
            label = QLabel()
            pixmap = QPixmap(screenshot_path)
            
            # 画像のアスペクト比を維持しながらスケーリング
            scaled_pixmap = pixmap.scaled(
                800,  # 最大幅
                4000,  # 十分な高さ（スクロール可能）
                Qt.AspectRatioMode.KeepAspectRatio,  # アスペクト比を維持
                Qt.TransformationMode.SmoothTransformation  # スムーズな変換
            )
            
            label.setPixmap(scaled_pixmap)
            
            # スクロールエリアにラベルを設定
            scroll_area.setWidget(label)
            layout.addWidget(scroll_area)
            
            dialog.setLayout(layout)
            dialog.exec()
            
        except Exception as e:
            logging.error(f"スクリーンショット表示エラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"スクリーンショットの表示中にエラーが発生しました: {str(e)}"
            )

    def search_service_area(self):
        """提供エリア検索を開始"""
        postal_code = self.postal_code_input.text().strip()
        address = self.address_input.text().strip()
        
        if not postal_code or not address:
            QMessageBox.warning(self, "入力エラー", "郵便番号と住所を入力してください。")
            return
        
        try:
            # 既存のスレッドとワーカーをクリーンアップ
            self.cleanup_thread()
            
            # 検索ボタンをキャンセルボタンに変更
            self.area_search_btn.setEnabled(True)
            self.area_search_btn.setText("キャンセル")
            self.area_search_btn.setStyleSheet("""
                QPushButton {
                    background-color: #E74C3C;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    text-align: center;
                    font-size: 14px;
                    margin: 4px 2px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #C0392B;
                }
                QPushButton:pressed {
                    background-color: #A93226;
                }
            """)
            self.area_search_btn.clicked.disconnect()
            self.area_search_btn.clicked.connect(self.cancel_search)
            
            # プログレスバーを表示
            self.progress_bar.setVisible(True)
            
            # 検索ステータスを更新
            self.area_result_label.setText("提供エリア: 検索を開始します...")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #3498DB;
                    border-radius: 4px;
                    background-color: #E3F2FD;
                    color: #2980B9;
                }
            """)
            
            # ワーカーを作成
            self.worker = ServiceAreaSearchWorker(postal_code, address)
            self.worker.finished.connect(self.on_search_completed)
            self.worker.progress.connect(self.update_search_progress)
            
            # スレッドを作成して検索を開始
            self.thread = QThread()
            self.worker.moveToThread(self.thread)
            self.thread.started.connect(self.worker.run)
            self.thread.finished.connect(self.thread.deleteLater)
            self.thread.start()
            
        except Exception as e:
            logging.error(f"検索の開始に失敗: {str(e)}")
            self.reset_search_button()
            QMessageBox.critical(self, "エラー", f"検索の開始に失敗しました: {str(e)}")

    def cancel_search(self):
        """提供エリア検索をキャンセルする"""
        # キャンセル中の状態をUIに即時反映
        self.area_search_btn.setEnabled(False)
        self.area_search_btn.setText("キャンセル中...")
        self.area_result_label.setText("提供エリア: キャンセル中...")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #F39C12;
                border-radius: 4px;
                background-color: #FFF3E0;
                color: #F39C12;
            }
        """)

        # バックエンド処理のキャンセル
        if hasattr(self, 'worker'):
            self.worker.cancel()
            # キャンセル完了を待つため、ボタンとプログレスバーはそのまま維持

    def reset_search_button(self):
        """検索ボタンを初期状態に戻す"""
        self.area_search_btn.setText("提供エリア検索")
        self.area_search_btn.setEnabled(True)
        self.area_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498DB;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2980B9;
            }
            QPushButton:pressed {
                background-color: #2471A3;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        # 検索ボタンのクリックイベントを元に戻す
        self.area_search_btn.clicked.disconnect()
        self.area_search_btn.clicked.connect(self.search_service_area)

    def on_search_completed(self, result):
        """検索完了時の処理"""
        # プログレスバーを非表示
        self.progress_bar.setVisible(False)
        
        status = result.get("status", "failure")
        
        if status == "cancelled":
            self.area_result_label.setText("提供エリア: 検索がキャンセルされました")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #F39C12;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #F39C12;
                }
            """)
            # キャンセル完了後に検索ボタンを初期状態に戻す
            self.reset_search_button()
            return
        
        # キャンセル以外の完了時の処理
        self.reset_search_button()
        
        if status == "available":
            self.area_result_label.setText("提供エリア: 提供可能")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #27AE60;
                    border-radius: 4px;
                    background-color: #E8F5E9;
                    color: #27AE60;
                }
            """)
            self.judgment_combo.setCurrentText("○")
        elif status == "unavailable":
            self.area_result_label.setText("提供エリア: 未提供")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #E74C3C;
                    border-radius: 4px;
                    background-color: #FFEBEE;
                    color: #E74C3C;
                }
            """)
            self.judgment_combo.setCurrentText("×")
        else:
            self.area_result_label.setText("提供エリア: 判定失敗")
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #F39C12;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #F39C12;
                }
            """)
            self.judgment_combo.setCurrentText("")

        # スクリーンショットの更新
        if "screenshot" in result:
            self.update_screenshot_button(result["screenshot"])

        # 詳細情報の表示
        if "details" in result and result.get("show_popup", True):
            details = result["details"]
            details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
            QMessageBox.information(self, "検索結果", details_text)

    def cleanup_thread(self):
        """
        スレッドのクリーンアップを行う
        """
        try:
            if self.thread and isinstance(self.thread, QThread):
                if self.thread.isRunning():
                    self.thread.quit()
                    self.thread.wait()
                self.thread.deleteLater()
                self.thread = None
        except Exception as e:
            logging.error(f"スレッドのクリーンアップ中にエラー: {str(e)}")
            # エラーが発生しても、スレッドをNoneに設定して続行
            self.thread = None

    def on_furigana_mode_changed(self, mode):
        """
        フリガナモードが変更された時の処理
        
        Args:
            mode (str): 選択されたモード（"自動" or "手動"）
        """
        if mode == "自動":
            # 自動モードの場合、現在の契約者名とリスト名からフリガナを生成
            self.auto_generate_furigana(self.contractor_input.text())
            self.auto_generate_list_furigana(self.list_name_input.text())
            self.auto_generate_address_furigana(self.address_input.text())
            self.auto_generate_list_address_furigana(self.list_address_input.text())
            # 自動生成を有効化
            self.contractor_input.textChanged.connect(self.auto_generate_furigana)
            self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
            self.address_input.textChanged.connect(self.auto_generate_address_furigana)
            self.list_address_input.textChanged.connect(self.auto_generate_list_address_furigana)
        else:
            # 手動モードの場合、自動生成を無効化
            try:
                self.contractor_input.textChanged.disconnect(self.auto_generate_furigana)
            except:
                pass
            try:
                self.list_name_input.textChanged.disconnect(self.auto_generate_list_furigana)
            except:
                pass
            try:
                self.address_input.textChanged.disconnect(self.auto_generate_address_furigana)
            except:
                pass
            try:
                self.list_address_input.textChanged.disconnect(self.auto_generate_list_address_furigana)
            except:
                pass

    def auto_generate_furigana(self, text):
        """
        契約者名からフリガナを自動生成
        
        Args:
            text (str): 契約者名
        """
        if self.furigana_mode_combo.currentText() == "自動":
            try:
                furigana = convert_to_furigana(text)
                if furigana:
                    self.furigana_input.setText(furigana)
            except Exception as e:
                logging.error(f"フリガナ自動生成エラー: {e}")

    def auto_generate_list_furigana(self, text):
        """
        リスト名からフリガナを自動生成
        
        Args:
            text (str): リスト名
        """
        if self.furigana_mode_combo.currentText() == "自動":
            try:
                furigana = convert_to_furigana(text)
                if furigana:
                    self.list_furigana_input.setText(furigana)
            except Exception as e:
                logging.error(f"リストフリガナ自動生成エラー: {e}")

    def auto_generate_address_furigana(self, text):
        """
        住所からフリガナを自動生成
        
        Args:
            text (str): 住所
        """
        if self.furigana_mode_combo.currentText() == "自動":
            try:
                furigana = convert_to_furigana(text)
                if furigana:
                    self.address_furigana_input.setText(furigana)
            except Exception as e:
                logging.error(f"住所フリガナ自動生成エラー: {e}")

    def create_area_search_widget(self):
        """エリア検索関連のウィジェットを作成"""
        search_container = QWidget()
        search_layout = QVBoxLayout(search_container)
        search_layout.setSpacing(8)
        search_layout.setContentsMargins(0, 8, 0, 0)
        
        # ボタンのスタイル
        button_style = """
            QPushButton {
                min-height: 30px;
                padding: 5px 15px;
                border: 1px solid #1565C0;
                border-radius: 2px;
                background-color: #1565C0;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0D47A1;
            }
            QPushButton:pressed {
                background-color: #0A3D91;
            }
            QPushButton:disabled {
                background-color: #ccc;
                border-color: #ccc;
            }
        """
        
        # 緑色のボタン（地図を開くボタン）のスタイル
        green_button_style = """
            QPushButton {
                min-height: 30px;
                padding: 5px 15px;
                border: 1px solid #2E7D32;
                border-radius: 2px;
                background-color: #2E7D32;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1B5E20;
            }
            QPushButton:pressed {
                background-color: #0A280E;
            }
            QPushButton:disabled {
                background-color: #ccc;
                border-color: #ccc;
            }
        """
        
        # エリア検索ボタンの作成
        self.area_search_btn = QPushButton("提供エリア検索")
        self.area_search_btn.setStyleSheet(button_style)
        
        # 地図ボタンの作成
        self.map_btn = QPushButton("地図を開く")
        self.map_btn.setStyleSheet(green_button_style)
        
        # エリア結果ラベルの作成
        self.area_result_label = QLabel("提供エリア: 未検索")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #95a5a6;
                border-radius: 4px;
                background-color: #f8f9fa;
                color: #95a5a6;
            }
        """)
        
        # プログレスバーの作成
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        button_container = QHBoxLayout()
        button_container.addWidget(self.area_search_btn)
        button_container.addWidget(self.map_btn)
        
        search_layout.addLayout(button_container)
        search_layout.addWidget(self.progress_bar)
        search_layout.addWidget(self.area_result_label)
        
        return search_container


class ServiceAreaSearchWorker(QObject):
    """
    提供エリア検索を実行するワーカークラス
    """
    finished = Signal(dict)  # 検索結果を通知するシグナル
    progress = Signal(str)   # 進捗状況を通知するシグナル
    
    def __init__(self, postal_code, address):
        super().__init__()
        self.postal_code = postal_code
        self.address = address
        self._is_cancelled = False
    
    def cancel(self):
        """
        検索をキャンセルする
        """
        self._is_cancelled = True
    
    def run(self):
        """
        提供エリア検索を実行し、結果をシグナルで通知する
        """
        try:
            # 進捗状況を通知するコールバック関数を定義
            def progress_callback(message):
                if self._is_cancelled:
                    raise CancellationError("検索がキャンセルされました")
                self.progress.emit(message)

            # 検索を実行
            result = search_service_area(
                self.postal_code,
                self.address,
                progress_callback=progress_callback
            )
            if self._is_cancelled:
                raise CancellationError("検索がキャンセルされました")
            self.finished.emit(result)
        except CancellationError as e:
            logging.info("検索がキャンセルされました")
            self.progress.emit("検索がキャンセルされました")
            self.finished.emit({
                "status": "cancelled",
                "message": "検索がキャンセルされました"
            })
        except Exception as e:
            logging.error(f"検索処理中にエラーが発生: {str(e)}")
            self.progress.emit("エラーが発生しました")
            self.finished.emit({
                "status": "error",
                "message": f"検索処理中にエラーが発生: {str(e)}"
            })

class CancellationError(Exception):
    """検索キャンセル時に発生する例外"""
    pass

