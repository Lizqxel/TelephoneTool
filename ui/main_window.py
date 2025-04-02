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
                              QSizePolicy, QProgressBar)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread
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
        self.setMinimumSize(600, 400)  # 最小サイズを600x400に変更
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        
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
        preview_group = QGroupBox("プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        self.create_preview_area(preview_layout)
        
        # スプリッターにウィジェットを追加
        splitter.addWidget(scroll_area)
        splitter.addWidget(preview_group)
        
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
        top_bar.setFixedHeight(32)
        top_bar.setStyleSheet("""
            QWidget {
                background-color: #2C3E50;
                color: white;
            }
        """)
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(5, 2, 5, 2)
        top_bar_layout.setSpacing(4)
        
        # ワンクリック取得ボタン（名称変更：顧客情報取得）
        self.oneclick_btn = QPushButton("顧客情報取得")
        self.oneclick_btn.setStyleSheet("""
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 2px 6px;
                border-radius: 2px;
                background-color: #27AE60;
                min-height: 18px;
                max-height: 22px;
            }
            QPushButton:hover {
                background-color: #2ECC71;
            }
            QPushButton:pressed {
                background-color: #27AE60;
            }
        """)
        self.oneclick_btn.clicked.connect(self.fetch_cti_data)
        self.oneclick_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        top_bar_layout.addWidget(self.oneclick_btn)
        
        # 既存のボタン
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("営コメ作成")
        self.screenshot_btn = QPushButton("提供判定のスクリーンショット確認")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記（未実装）")
        self.settings_btn = QPushButton("設定")
        
        # ボタンのスタイル設定
        button_style = """
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 2px 6px;
                border-radius: 2px;
                min-height: 18px;
                max-height: 22px;
                font-size: 9pt;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
            QPushButton:pressed {
                background-color: #2C3E50;
            }
        """
        
        # 各ボタンのサイズポリシーを設定
        buttons = [self.clear_btn, self.cti_copy_btn, 
                  self.screenshot_btn, self.spreadsheet_btn, self.settings_btn]
        
        for btn in buttons:
            btn.setStyleSheet(button_style)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        # ボタンの接続
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.copy_cti_to_clipboard)
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)
        self.settings_btn.clicked.connect(self.show_settings)
        
        # ボタンをレイアウトに追加
        for btn in buttons:
            top_bar_layout.addWidget(btn)
        
        parent_layout.addWidget(top_bar)
    
    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        # 受注者入力項目セクション（新しく追加）
        input_group = QGroupBox("受注者入力項目")
        input_layout = QVBoxLayout()
        
        # 対応者名
        input_layout.addWidget(QLabel("対応者名"))
        self.operator_input = QLineEdit()
        input_layout.addWidget(self.operator_input)
        
        # 出やすい時間帯
        input_layout.addWidget(QLabel("出やすい時間帯"))
        self.available_time_input = QLineEdit()
        self.available_time_input.setPlaceholderText("AMPM希望　固定or携帯　000-0000-0000")
        input_layout.addWidget(self.available_time_input)
        
        # 契約者名
        input_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        input_layout.addWidget(self.contractor_input)
        
        # フリガナ
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ"))
        self.furigana_mode_combo = CustomComboBox()
        self.furigana_mode_combo.addItems(["自動", "手動"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        input_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        input_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QVBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        # 生年月日の入力部分を横並びにする
        birth_input_layout = QHBoxLayout()
        birth_input_layout.setContentsMargins(0, 0, 0, 0)  # マージンを0に設定
        birth_input_layout.setSpacing(5)  # 間隔を5pxに設定
        
        # コンテナウィジェットを作成して固定幅を設定
        birth_container = QWidget()
        birth_container.setFixedWidth(320)  # 固定幅を280pxから320pxに増加
        birth_container_layout = QHBoxLayout(birth_container)
        birth_container_layout.setContentsMargins(0, 0, 0, 0)
        birth_container_layout.setSpacing(8)  # 間隔を5pxから8pxに増加
        
        self.era_combo = CustomComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        self.era_combo.setFixedWidth(80)  # 幅を70pxから80pxに増加
        birth_container_layout.addWidget(self.era_combo)
        
        # 年の入力部分用のコンテナ
        year_container = QWidget()
        year_layout = QHBoxLayout(year_container)
        year_layout.setContentsMargins(0, 0, 0, 0)
        year_layout.setSpacing(2)  # 年の数字と「年」の間隔を調整
        
        self.year_combo = CustomComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))
        self.year_combo.setFixedWidth(65)  # 幅を60pxから65pxに増加
        year_layout.addWidget(self.year_combo)
        
        year_label = QLabel("年")
        year_label.setFixedWidth(20)
        year_layout.addWidget(year_label)
        
        birth_container_layout.addWidget(year_container)
        
        # 月の入力部分用のコンテナ
        month_container = QWidget()
        month_layout = QHBoxLayout(month_container)
        month_layout.setContentsMargins(0, 0, 0, 0)
        month_layout.setSpacing(2)
        
        self.month_combo = CustomComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))
        self.month_combo.setFixedWidth(45)
        month_layout.addWidget(self.month_combo)
        
        month_label = QLabel("月")
        month_label.setFixedWidth(20)
        month_layout.addWidget(month_label)
        
        birth_container_layout.addWidget(month_container)
        
        # 日の入力部分用のコンテナ
        day_container = QWidget()
        day_layout = QHBoxLayout(day_container)
        day_layout.setContentsMargins(0, 0, 0, 0)
        day_layout.setSpacing(2)
        
        self.day_combo = CustomComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.setEditable(True)
        self.day_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.day_combo.lineEdit().setMaxLength(2)
        self.day_combo.lineEdit().setValidator(QIntValidator(1, 31))
        self.day_combo.setFixedWidth(45)
        day_layout.addWidget(self.day_combo)
        
        day_label = QLabel("日")
        day_label.setFixedWidth(20)
        day_layout.addWidget(day_label)
        
        birth_container_layout.addWidget(day_container)
        
        # コンテナをレイアウトに追加
        birth_input_layout.addWidget(birth_container)
        birth_input_layout.addStretch()  # 右側に余白を追加
        
        birth_layout.addLayout(birth_input_layout)
        input_layout.addLayout(birth_layout)
        
        # 受注者名
        input_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        input_layout.addWidget(self.order_person_input)
        
        # 料金認識を追加（移動）
        input_layout.addWidget(QLabel("料金認識"))
        self.fee_input = QLineEdit()
        self.fee_input.setText("2500円～3000円")
        input_layout.addWidget(self.fee_input)
        
        # ネット利用
        input_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = CustomComboBox()
        self.net_usage_combo.addItems(["なし", "あり"])
        input_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        input_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = CustomComboBox()
        self.family_approval_combo.addItems(["ok", "なし"])
        input_layout.addWidget(self.family_approval_combo)
        
        # 他番号
        input_layout.addWidget(QLabel("他番号"))
        self.other_number_input = QLineEdit()
        self.other_number_input.setText("なし")
        input_layout.addWidget(self.other_number_input)
        
        # 電話機
        input_layout.addWidget(QLabel("電話機"))
        self.phone_device_input = QLineEdit()
        self.phone_device_input.setText("プッシュホン")
        input_layout.addWidget(self.phone_device_input)
        
        # 禁止回線
        input_layout.addWidget(QLabel("禁止回線"))
        self.forbidden_line_input = QLineEdit()
        self.forbidden_line_input.setText("なし")
        input_layout.addWidget(self.forbidden_line_input)
        
        # ND
        input_layout.addWidget(QLabel("ND"))
        self.nd_input = QLineEdit()
        input_layout.addWidget(self.nd_input)
        
        # リストとの関係性（表示を「名義人の○○」の形式に変更）
        relationship_layout = QHBoxLayout()
        relationship_layout.addWidget(QLabel("備考："))
        self.relationship_input = QLineEdit()
        self.relationship_input.setPlaceholderText("名義人の...")
        relationship_layout.addWidget(self.relationship_input)
        input_layout.addLayout(relationship_layout)
        
        input_group.setLayout(input_layout)
        parent_layout.addWidget(input_group)
        
        # 住所情報セクション
        address_group = QGroupBox("住所情報")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        address_layout.addWidget(self.address_input)
        
        # 住所フリガナ
        address_layout.addWidget(QLabel("住所フリガナ"))
        self.address_furigana_input = QLineEdit()
        address_layout.addWidget(self.address_furigana_input)
        
        # マップアイコンボタン
        self.map_btn = QPushButton()
        self.map_btn.setFixedSize(24, 24)
        
        # アプリケーションの実行ディレクトリからの絶対パスを設定
        app_dir = os.path.dirname(os.path.abspath(__file__))
        root_dir = os.path.dirname(app_dir)  # uiフォルダの親ディレクトリ
        map_icon_path = os.path.join(root_dir, "map.png")
        
        # アイコンが存在する場合のみ設定
        if os.path.exists(map_icon_path):
            self.map_btn.setIcon(QIcon(map_icon_path))
        else:
            # アイコンが見つからない場合は代替テキストを設定
            self.map_btn.setText("🗺️")
            logging.warning(f"マップアイコン画像が見つかりません: {map_icon_path}")
            
        self.map_btn.setToolTip("Googleマップで住所を検索")
        self.map_btn.setStyleSheet("""
            QPushButton {
                border: none;
                background-color: transparent;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
                border-radius: 12px;
            }
        """)
        address_layout.addWidget(self.map_btn)
        
        # 提供エリア検索ボタンを追加
        self.area_search_btn = QPushButton("提供エリア検索")
        self.area_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }
        """)
        address_layout.addWidget(self.area_search_btn)
        
        # 提供エリア検索結果表示用のラベル
        area_result_container = QWidget()
        area_result_layout = QVBoxLayout(area_result_container)
        area_result_layout.setContentsMargins(0, 0, 0, 0)
        area_result_layout.setSpacing(2)

        self.area_result_label = QLabel("提供エリア: 未検索")
        self.area_result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: #f8f8f8;
            }
        """)
        area_result_layout.addWidget(self.area_result_label)

        # プログレスバー（初期状態では非表示）
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # 不定のプログレスバー
        self.progress_bar.setFixedHeight(2)  # 高さを2ピクセルに設定
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #E3F2FD;
            }
            QProgressBar::chunk {
                background-color: #3498DB;
            }
        """)
        self.progress_bar.hide()  # 初期状態では非表示
        area_result_layout.addWidget(self.progress_bar)

        address_layout.addWidget(area_result_container)
        
        address_group.setLayout(address_layout)
        parent_layout.addWidget(address_group)
        
        # リスト情報セクション
        list_group = QGroupBox("リスト情報")
        list_layout = QVBoxLayout()
        
        # リスト名
        list_layout.addWidget(QLabel("リスト名"))
        self.list_name_input = QLineEdit()
        list_layout.addWidget(self.list_name_input)
        
        # リストフリガナ
        list_furigana_layout = QHBoxLayout()
        list_furigana_layout.addWidget(QLabel("リストフリガナ"))
        self.list_furigana_mode_combo = CustomComboBox()
        self.list_furigana_mode_combo.addItems(["自動", "手動"])
        list_furigana_layout.addWidget(self.list_furigana_mode_combo)
        list_layout.addLayout(list_furigana_layout)
        self.list_furigana_input = QLineEdit()
        list_layout.addWidget(self.list_furigana_input)
        
        # 電話番号
        list_layout.addWidget(QLabel("電話番号"))
        self.list_phone_input = QLineEdit()
        list_layout.addWidget(self.list_phone_input)
        
        # リスト郵便番号
        list_layout.addWidget(QLabel("リスト郵便番号"))
        self.list_postal_code_input = QLineEdit()
        list_layout.addWidget(self.list_postal_code_input)
        
        # リスト住所
        list_layout.addWidget(QLabel("リスト住所"))
        list_address_container = QHBoxLayout()
        self.list_address_input = QLineEdit()
        list_address_container.addWidget(self.list_address_input)
        list_layout.addLayout(list_address_container)
        
        # リスト住所フリガナモード
        list_address_furigana_layout = QHBoxLayout()
        list_address_furigana_layout.addWidget(QLabel("リスト住所フリガナ"))
        self.list_address_furigana_mode_combo = CustomComboBox()
        self.list_address_furigana_mode_combo.addItems(["自動", "手動"])
        list_address_furigana_layout.addWidget(self.list_address_furigana_mode_combo)
        list_layout.addLayout(list_address_furigana_layout)
        self.list_address_furigana_input = QLineEdit()
        list_layout.addWidget(self.list_address_furigana_input)
        
        list_group.setLayout(list_layout)
        parent_layout.addWidget(list_group)
        
        # 受注情報セクション
        order_group = QGroupBox("受注情報")
        order_layout = QVBoxLayout()
        
        # 現状回線
        order_layout.addWidget(QLabel("現状回線"))
        self.current_line_combo = CustomComboBox()
        self.current_line_combo.addItems(["アナログ"])
        order_layout.addWidget(self.current_line_combo)
        
        # 受注日（本日自動入力）
        order_layout.addWidget(QLabel("受注日"))
        self.order_date_input = QLineEdit()
        # 0埋めなしの月/日フォーマットを生成
        now = datetime.datetime.now()
        month = str(now.month)  # 0埋めなしの月
        day = str(now.day)      # 0埋めなしの日
        self.order_date_input.setText(f"{month}/{day}")
        self.order_date_input.setReadOnly(True)
        order_layout.addWidget(self.order_date_input)
        
        # 提供判定
        order_layout.addWidget(QLabel("提供判定"))
        self.judgment_combo = CustomComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        parent_layout.addWidget(order_group)
    
    def create_preview_area(self, parent_layout):
        """プレビューエリアを作成"""
        # プレビューラベル
        preview_label = QLabel("CTIフォーマットプレビュー")
        
        # プレビューテキストエリア
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet("""
            QTextEdit {
                background-color: #f8f8f8;
                color: #333333;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                font-family: 'MS Gothic', monospace;
            }
        """)
        
        parent_layout.addWidget(preview_label)
        parent_layout.addWidget(self.preview_text)
    
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
        
        # 名前とフリガナのバリデーション用のシグナル
        self.contractor_input.textChanged.connect(self.validate_contractor_name)
        self.furigana_input.textChanged.connect(self.validate_furigana_input)
        self.list_name_input.textChanged.connect(self.validate_list_name)
        self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
        
        # フリガナ自動変換のシグナル
        self.contractor_input.textChanged.connect(self.auto_generate_furigana)
        self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
        
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
        self.nd_input.textChanged.connect(self.reset_background_color)  # NDの背景色リセット
        
        # ボタンのシグナル接続
        self.area_search_btn.clicked.connect(self.search_service_area)
        
        # マップボタンのシグナル接続
        self.map_btn.clicked.connect(self.open_street_view)
    
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
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
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

