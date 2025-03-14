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
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QComboBox, QPushButton,
                              QTextEdit, QGroupBox, QMessageBox, QScrollArea,
                              QApplication, QToolTip)
from PySide6.QtCore import Qt, QTimer, QPoint
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from ui.main_window_functions import MainWindowFunctions
from utils.string_utils import validate_name, validate_furigana


class MainWindow(QMainWindow, MainWindowFunctions):
    """メインウィンドウクラス"""
    
    def __init__(self):
        """メインウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(800, 600)  # 最小ウィンドウサイズを縮小
        
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
        
        # フォントサイズの適用
        self.apply_font_size()
    
    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        top_bar = QWidget()
        top_bar.setStyleSheet("background-color: #2C3E50; color: white;")
        top_bar_layout = QHBoxLayout(top_bar)
        
        # クリップボード監視ボタンを追加
        self.clipboard_toggle_btn = QPushButton("クリップボード監視")
        self.clipboard_toggle_btn.setCheckable(True)  # トグルボタンにする
        self.clipboard_toggle_btn.setStyleSheet("""
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
                background-color: #2C3E50;
            }
            QPushButton:checked {
                background-color: #27AE60;
            }
            QPushButton:hover {
                background-color: #34495E;
                border: 1px solid #3498DB;
            }
            QPushButton:pressed {
                background-color: #2980B9;
            }
        """)
        top_bar_layout.addWidget(self.clipboard_toggle_btn)
        
        # 既存のボタン
        self.settings_btn = QPushButton("設定")
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("CTIコピー")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
        
        # スクリーンショット表示ボタンを追加
        self.screenshot_btn = QPushButton("スクリーンショット表示")
        self.screenshot_btn.setEnabled(False)  # 初期状態は無効
        
        # 既存のボタンにスタイルを適用
        button_style = """
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
                background-color: #2C3E50;
            }
            QPushButton:hover {
                background-color: #34495E;
                border: 1px solid #3498DB;
            }
            QPushButton:pressed {
                background-color: #2980B9;
            }
            QPushButton:disabled {
                color: #95A5A6;
                border: 1px solid #95A5A6;
                background-color: #34495E;
            }
        """
        
        self.settings_btn.setStyleSheet(button_style)
        self.clear_btn.setStyleSheet(button_style)
        self.cti_copy_btn.setStyleSheet(button_style)
        self.spreadsheet_btn.setStyleSheet(button_style)
        self.screenshot_btn.setStyleSheet(button_style)
        
        # ボタンをレイアウトに追加
        top_bar_layout.addWidget(self.settings_btn)
        top_bar_layout.addWidget(self.clear_btn)
        top_bar_layout.addWidget(self.cti_copy_btn)
        top_bar_layout.addWidget(self.spreadsheet_btn)
        top_bar_layout.addWidget(self.screenshot_btn)
        
        parent_layout.addWidget(top_bar)
    
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
        # 既存のシグナル設定
        self.settings_btn.clicked.connect(self.show_settings_dialog)
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.generate_cti_format)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)
        self.clipboard_toggle_btn.clicked.connect(self.toggle_clipboard_monitor)
        
        # スクリーンショット表示ボタンのシグナル設定
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        
        # 自動フォーマット用のシグナル
        self.mobile_input.textChanged.connect(self.format_phone_number)
        self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
        self.postal_code_input.textChanged.connect(self.format_postal_code)
        self.postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.list_postal_code_input.textChanged.connect(self.format_postal_code)
        self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
        self.address_input.textChanged.connect(self.convert_hyphen_to_half_width)
        self.list_address_input.textChanged.connect(self.convert_hyphen_to_half_width_list)
        self.era_combo.currentTextChanged.connect(self.update_year_combo)
        
        # 名前とフリガナのバリデーション用のシグナル
        self.contractor_input.textChanged.connect(self.validate_contractor_name)
        self.furigana_input.textChanged.connect(self.validate_furigana_input)
        self.list_name_input.textChanged.connect(self.validate_list_name)
        self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
        
        # ボタンのシグナル接続
        self.area_search_btn.clicked.connect(self.search_service_area)
    
    def show_screenshot(self):
        """最新のスクリーンショットを表示"""
        if hasattr(self, 'latest_screenshot_path') and os.path.exists(self.latest_screenshot_path):
            # スクリーンショット表示用のダイアログを作成
            dialog = QMessageBox(self)
            dialog.setWindowTitle("スクリーンショット")
            
            # スクリーンショットを読み込んでラベルに設定
            pixmap = QPixmap(self.latest_screenshot_path)
            
            # スクリーンサイズの80%を上限とする
            screen_size = QApplication.primaryScreen().size()
            max_width = int(screen_size.width() * 0.8)
            max_height = int(screen_size.height() * 0.8)
            
            # 画像のサイズを調整
            if pixmap.width() > max_width or pixmap.height() > max_height:
                pixmap = pixmap.scaled(max_width, max_height, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            
            # ラベルを作成して画像を設定
            label = QLabel()
            label.setPixmap(pixmap)
            
            # ダイアログにラベルを設定
            dialog.layout().addWidget(label, 0, 0, 1, dialog.layout().columnCount())
            dialog.setStyleSheet("QMessageBox { background-color: white; }")
            
            # OKボタンのみ表示
            dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
            
            # ダイアログを表示
            dialog.exec()
        else:
            QMessageBox.warning(self, "エラー", "スクリーンショットが見つかりません。")

    def update_screenshot_button(self, screenshot_path=None):
        """スクリーンショット表示ボタンの状態を更新"""
        if screenshot_path and os.path.exists(screenshot_path):
            self.latest_screenshot_path = screenshot_path
            self.screenshot_btn.setEnabled(True)
        else:
            self.screenshot_btn.setEnabled(False)

    def convert_hyphen_to_half_width(self, text=None):
        """住所の数字とハイフンを半角に変換"""
        if text is None:
            text = self.address_input.text()
        
        # 全角数字と全角ハイフンの変換マップ
        conversion_map = str.maketrans({
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
            '－': '-', 'ー': '-', '―': '-', '‐': '-', '−': '-'
        })
        
        # 変換を実行
        half_width_text = text.translate(conversion_map)
        
        # テキストが変更された場合のみ更新
        if text != half_width_text:
            cursor_pos = self.address_input.cursorPosition()
            self.address_input.setText(half_width_text)
            self.address_input.setCursorPosition(cursor_pos)
        
        return half_width_text

    def convert_hyphen_to_half_width_list(self, text=None):
        """リスト住所の数字とハイフンを半角に変換"""
        if text is None:
            text = self.list_address_input.text()
        
        # 全角数字と全角ハイフンの変換マップ
        conversion_map = str.maketrans({
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
            '－': '-', 'ー': '-', '―': '-', '‐': '-', '−': '-'
        })
        
        # 変換を実行
        half_width_text = text.translate(conversion_map)
        
        # テキストが変更された場合のみ更新
        if text != half_width_text:
            cursor_pos = self.list_address_input.cursorPosition()
            self.list_address_input.setText(half_width_text)
            self.list_address_input.setCursorPosition(cursor_pos)
        
        return half_width_text

    def validate_contractor_name(self):
        """契約者名のバリデーション"""
        text = self.contractor_input.text()
        if not validate_name(text):
            self.contractor_input.setStyleSheet("""
                QLineEdit {
                    background-color: #FFD7D7;
                    border: 2px solid #FF8080;
                }
            """)
            QToolTip.showText(
                self.contractor_input.mapToGlobal(QPoint(0, 0)),
                "名前に数字を含めることはできません",
                self.contractor_input,
            )
        else:
            self.contractor_input.setStyleSheet("")
            QToolTip.hideText()

    def validate_furigana_input(self):
        """フリガナのバリデーション"""
        text = self.furigana_input.text()
        if not validate_furigana(text):
            self.furigana_input.setStyleSheet("""
                QLineEdit {
                    background-color: #FFD7D7;
                    border: 2px solid #FF8080;
                }
            """)
            QToolTip.showText(
                self.furigana_input.mapToGlobal(QPoint(0, 0)),
                "フリガナに数字や不適切な文字を含めることはできません",
                self.furigana_input
            )
        else:
            self.furigana_input.setStyleSheet("")
            QToolTip.hideText()

    def validate_list_name(self):
        """リスト名のバリデーション"""
        text = self.list_name_input.text()
        if not validate_name(text):
            self.list_name_input.setStyleSheet("""
                QLineEdit {
                    background-color: #FFD7D7;
                    border: 2px solid #FF8080;
                }
            """)
            QToolTip.showText(
                self.list_name_input.mapToGlobal(QPoint(0, 0)),
                "名前に数字を含めることはできません",
                self.list_name_input
            )
        else:
            self.list_name_input.setStyleSheet("")
            QToolTip.hideText()

    def validate_list_furigana(self):
        """リストフリガナのバリデーション"""
        text = self.list_furigana_input.text()
        if not validate_furigana(text):
            self.list_furigana_input.setStyleSheet("""
                QLineEdit {
                    background-color: #FFD7D7;
                    border: 2px solid #FF8080;
                }
            """)
            QToolTip.showText(
                self.list_furigana_input.mapToGlobal(QPoint(0, 0)),
                "フリガナに数字や不適切な文字を含めることはできません",
                self.list_furigana_input
            )
        else:
            self.list_furigana_input.setStyleSheet("")
            QToolTip.hideText()

    def analyze_clipboard_content(self, text):
        """クリップボードの内容を解析して適切なフィールドに入力"""
        # 電話番号（ハイフンあり/なし）のパターン
        phone_pattern = re.compile(r'(\d{2,4}[-\s]?\d{2,4}[-\s]?\d{4})')
        phone_matches = phone_pattern.finditer(text)
        
        # 郵便番号（ハイフンあり/なし）のパターン
        postal_pattern = re.compile(r'(\d{3}[-\s]?\d{4})')
        postal_match = postal_pattern.search(text)
        
        # 電話番号の処理
        for match in phone_matches:
            phone_number = match.group(1)
            # 携帯電話番号の判定（070, 080, 090で始まる番号）
            if phone_number.replace('-', '').replace(' ', '').startswith(('070', '080', '090')):
                self.mobile_input.setText(phone_number)
                self.mobile_type_combo.setCurrentText("入力")
            else:
                self.list_phone_input.setText(phone_number)
        
        # 郵便番号の処理
        if postal_match:
            postal_code = postal_match.group(1)
            self.postal_code_input.setText(postal_code)
            self.list_postal_code_input.setText(postal_code)
        
        # 住所らしき文字列（漢字とカタカナが含まれる長い文字列）
        if len(text) > 10 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            self.address_input.setText(text)
            self.list_address_input.setText(text)
        
        # カタカナのみの文字列（フリガナとして扱う）
        if all(ord(c) >= 0x30A0 and ord(c) <= 0x30FF or c.isspace() for c in text):
            self.list_furigana_input.setText(text)
        
        # その他の文字列（名前として扱う）
        if len(text) <= 20 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            # 数字が含まれていない場合のみ、名前として処理
            if validate_name(text):
                self.list_name_input.setText(text)
                # フリガナが空の場合は、カタカナ変換を試みる
                if not self.list_furigana_input.text():
                    try:
                        import pykakasi
                        kakasi = pykakasi.kakasi()
                        result = kakasi.convert(text)
                        katakana = ''.join([item['kana'] for item in result])
                        self.list_furigana_input.setText(katakana)
                    except:
                        pass  # カタカナ変換に失敗した場合は何もしない

    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.format_template = settings.get('format_template', '')
                    self.font_size = settings.get('font_size', 9)
            else:
                self.format_template = ''
                self.font_size = 9
        except Exception as e:
            logging.error(f"設定の読み込みに失敗しました: {str(e)}")
            self.format_template = ''
            self.font_size = 9

    def apply_font_size(self):
        """アプリケーション全体のフォントサイズを設定する"""
        # アプリケーション全体のデフォルトフォントを設定
        font = QApplication.font()
        font.setPointSize(self.font_size)
        QApplication.setFont(font)
        
        # メインウィンドウ内のすべてのウィジェットにフォントを適用
        for widget in self.findChildren(QWidget):
            widget_font = widget.font()
            widget_font.setPointSize(self.font_size)
            widget.setFont(widget_font)
        
        # プレビューテキストエリアのフォントサイズを設定
        preview_font = self.preview_text.font()
        preview_font.setPointSize(self.font_size)
        self.preview_text.setFont(preview_font)
        
        # ウィジェットを更新
        self.update()

    def show_settings_dialog(self):
        """設定ダイアログを表示する"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            settings = dialog.get_settings()
            self.format_template = settings['format_template']
            self.font_size = settings['font_size']
            self.apply_font_size()
            self.load_settings()  # 設定を再読み込み 