"""
メインウィンドウ

このモジュールは、アプリケーションのメインウィンドウを
提供します。
"""

import datetime
import logging
import json
import os
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QComboBox, QPushButton,
                              QTextEdit, QGroupBox, QMessageBox, QScrollArea,
                              QApplication)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QIntValidator, QClipboard

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from ui.main_window_functions import MainWindowFunctions


class MainWindow(QMainWindow, MainWindowFunctions):
    """メインウィンドウクラス"""
    
    def __init__(self):
        """メインウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(1000, 800)
        
        # クリップボード監視用の変数
        self.clipboard = QApplication.clipboard()
        self.last_clipboard_text = ""
        self.clipboard_timer = QTimer()
        self.clipboard_timer.timeout.connect(self.check_clipboard)
        
        # 受注者名の初期化
        self.order_person = ""
        
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
        
        # メイン部分のレイアウト
        content_layout = QHBoxLayout()
        
        # 入力フォームエリア（左側70%）をスクロール可能に
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        self.create_input_form(form_layout)
        
        # スクロールエリアの作成
        scroll_area = QScrollArea()
        scroll_area.setWidget(form_widget)
        scroll_area.setWidgetResizable(True)  # ウィジェットのリサイズを許可
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
                background: none;
                height: 0px;
            }
        """)
        
        # スクロールエリアをレイアウトに追加（左側70%）
        content_layout.addWidget(scroll_area, 70)
        
        # プレビューエリア（右側30%）
        preview_group = QGroupBox("プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        self.create_preview_area(preview_layout)
        content_layout.addWidget(preview_group, 30)
        
        # コンテンツレイアウトをメインレイアウトに追加
        main_layout.addLayout(content_layout)
        
        # シグナルの設定
        self.setup_signals()
        
        # Google Sheetsの設定
        self.setup_google_sheets()
        
        # クリップボード監視の初期化
        self.toggle_clipboard_monitor()
    
    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        # トップバーのレイアウト
        top_bar = QHBoxLayout()
        
        # 設定ボタン
        self.settings_btn = QPushButton("設定")
        self.settings_btn.setFixedWidth(80)
        self.settings_btn.setStyleSheet("""
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
        top_bar.addWidget(self.settings_btn)
        
        # クリップボード監視トグルボタン
        self.clipboard_toggle_btn = QPushButton("クリップボード監視: オフ")
        self.clipboard_toggle_btn.setFixedWidth(200)
        self.clipboard_toggle_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:pressed {
                background-color: #b71c1c;
            }
        """)
        top_bar.addWidget(self.clipboard_toggle_btn)
        
        # スペーサー
        top_bar.addStretch()
        
        # CTIフォーマットコピーボタン
        self.cti_copy_btn = QPushButton("CTIフォーマットをコピー")
        self.cti_copy_btn.setFixedWidth(200)
        self.cti_copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0b7dda;
            }
            QPushButton:pressed {
                background-color: #0a69b7;
            }
        """)
        top_bar.addWidget(self.cti_copy_btn)
        
        # スプレッドシート転記ボタン
        self.spreadsheet_btn = QPushButton("スプレッドシートに転記")
        self.spreadsheet_btn.setFixedWidth(200)
        self.spreadsheet_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e68a00;
            }
            QPushButton:pressed {
                background-color: #cc7a00;
            }
        """)
        top_bar.addWidget(self.spreadsheet_btn)
        
        # クリアボタン
        self.clear_btn = QPushButton("クリア")
        self.clear_btn.setFixedWidth(80)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #9E9E9E;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #757575;
            }
            QPushButton:pressed {
                background-color: #616161;
            }
        """)
        top_bar.addWidget(self.clear_btn)
        
        # トップバーをメインレイアウトに追加
        parent_layout.addLayout(top_bar)
    
    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        # 基本情報セクション
        basic_info_group = QGroupBox("基本情報")
        basic_layout = QVBoxLayout()
        
        # 対応者名
        basic_layout.addWidget(QLabel("対応者名"))
        self.operator_input = QLineEdit()
        basic_layout.addWidget(self.operator_input)
        
        # 携帯電話番号
        basic_layout.addWidget(QLabel("携帯電話番号"))
        mobile_layout = QHBoxLayout()
        self.mobile_type_combo = QComboBox()
        self.mobile_type_combo.addItems(["入力", "なし"])
        self.mobile_type_combo.currentTextChanged.connect(self.toggle_mobile_input)
        mobile_layout.addWidget(self.mobile_type_combo)
        self.mobile_input = QLineEdit()
        mobile_layout.addWidget(self.mobile_input)
        basic_layout.addLayout(mobile_layout)
        
        # 契約者名
        basic_layout.addWidget(QLabel("契約者名"))
        self.contractor_input = QLineEdit()
        basic_layout.addWidget(self.contractor_input)
        
        # フリガナ
        basic_layout.addWidget(QLabel("フリガナ"))
        self.furigana_input = QLineEdit()
        basic_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_layout = QHBoxLayout()
        birth_layout.addWidget(QLabel("生年月日"))
        
        self.era_combo = QComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        birth_layout.addWidget(self.era_combo)
        
        self.year_combo = QComboBox()
        # 初期値として昭和の年を設定
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)  # 最大4桁
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))  # 1-9999の範囲で制限
        birth_layout.addWidget(self.year_combo)
        birth_layout.addWidget(QLabel("年"))
        
        self.month_combo = QComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)  # 最大2桁
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))  # 1-12の範囲で制限
        birth_layout.addWidget(self.month_combo)
        birth_layout.addWidget(QLabel("月"))
        
        self.day_combo = QComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.setEditable(True)
        self.day_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.day_combo.lineEdit().setMaxLength(2)  # 最大2桁
        self.day_combo.lineEdit().setValidator(QIntValidator(1, 31))  # 1-31の範囲で制限
        birth_layout.addWidget(self.day_combo)
        birth_layout.addWidget(QLabel("日"))
        
        basic_layout.addLayout(birth_layout)
        basic_info_group.setLayout(basic_layout)
        parent_layout.addWidget(basic_info_group)
        
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
        address_layout.addWidget(self.area_result_label)
        
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
        list_layout.addWidget(QLabel("リストフリガナ"))
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
        self.list_address_input = QLineEdit()
        list_layout.addWidget(self.list_address_input)
        
        list_group.setLayout(list_layout)
        parent_layout.addWidget(list_group)
        
        # 受注情報セクション
        order_group = QGroupBox("受注情報")
        order_layout = QVBoxLayout()
        
        # 現状回線
        order_layout.addWidget(QLabel("現状回線"))
        self.current_line_combo = QComboBox()
        self.current_line_combo.addItems(["アナログ"])
        order_layout.addWidget(self.current_line_combo)
        
        # 受注日（本日自動入力）
        order_layout.addWidget(QLabel("受注日"))
        self.order_date_input = QLineEdit()
        self.order_date_input.setText(datetime.datetime.now().strftime("%Y/%m/%d"))
        self.order_date_input.setReadOnly(True)
        order_layout.addWidget(self.order_date_input)
        
        # 受注者名
        order_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        order_layout.addWidget(self.order_person_input)
        
        # 提供判定
        order_layout.addWidget(QLabel("提供判定"))
        self.judgment_combo = QComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        parent_layout.addWidget(order_group)
        
        # その他情報セクション
        other_group = QGroupBox("その他情報")
        other_layout = QVBoxLayout()
        
        # 料金認識
        other_layout.addWidget(QLabel("料金認識"))
        self.fee_input = QLineEdit()
        self.fee_input.setText("3000円～3500円")
        other_layout.addWidget(self.fee_input)
        
        # ネット利用
        other_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = QComboBox()
        self.net_usage_combo.addItems(["あり", "なし"])
        other_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        other_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = QComboBox()
        self.family_approval_combo.addItems(["あり", "なし"])
        other_layout.addWidget(self.family_approval_combo)
        
        # 備考
        other_layout.addWidget(QLabel("備考"))
        self.remarks_input = QTextEdit()
        self.remarks_input.setMaximumHeight(100)
        other_layout.addWidget(self.remarks_input)
        
        other_group.setLayout(other_layout)
        parent_layout.addWidget(other_group)
    
    def create_preview_area(self, parent_layout):
        """プレビューエリアを作成"""
        # プレビューテキストエリア
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setPlaceholderText("ここにプレビューが表示されます")
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
        parent_layout.addWidget(self.preview_text)
    
    def setup_signals(self):
        """シグナルの接続"""
        # 自動フォーマット用のシグナル
        self.mobile_input.textChanged.connect(self.format_phone_number)
        self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
        self.postal_code_input.textChanged.connect(self.format_postal_code)
        self.postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.list_postal_code_input.textChanged.connect(self.format_postal_code)
        self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.address_input.textChanged.connect(self.convert_to_half_width)
        self.list_address_input.textChanged.connect(self.convert_to_half_width)
        self.era_combo.currentTextChanged.connect(self.update_year_combo)
        
        # ボタンのシグナル接続
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.generate_cti_format)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)
        
        # 提供エリア検索ボタンのシグナル接続
        self.area_search_btn.clicked.connect(self.search_service_area)
        
        # 設定ボタンのシグナル接続
        self.settings_btn.clicked.connect(self.show_settings)
        
        # クリップボード監視ボタンのシグナル接続
        self.clipboard_toggle_btn.clicked.connect(self.toggle_clipboard_monitor) 