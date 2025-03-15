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
import requests
from urllib.parse import quote
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QComboBox, QPushButton,
                              QTextEdit, QGroupBox, QMessageBox, QScrollArea,
                              QApplication, QToolTip)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from ui.main_window_functions import MainWindowFunctions
from utils.string_utils import validate_name, validate_furigana
from utils.furigana_utils import convert_to_furigana
from services.oneclick import OneClickService


class FocusWheelComboBox(QComboBox):
    """フォーカス時のみホイールで値を変更できるコンボボックス"""
    
    def __init__(self):
        """コンボボックスの初期化"""
        super().__init__()
        self.clicked = False  # クリック状態を追跡する変数を追加
    
    def wheelEvent(self, event):
        """ホイールイベントの処理"""
        if self.clicked and self.hasFocus():  # クリックされていてフォーカスがある場合のみスクロール可能
            super().wheelEvent(event)
        else:
            # フォーカスがない場合は、イベントを親ウィジェットに伝播
            parent = self.parent()
            while parent is not None:
                if isinstance(parent, QScrollArea):
                    parent.wheelEvent(event)
                    break
                parent = parent.parent()
            if parent is None:
                event.ignore()

    def mousePressEvent(self, event):
        """マウスクリックイベントの処理"""
        super().mousePressEvent(event)
        self.clicked = True  # クリック状態をTrueに設定
        self.setFocus()  # フォーカスを設定

    def focusOutEvent(self, event):
        """フォーカスが外れた時の処理"""
        super().focusOutEvent(event)
        self.clicked = False  # クリック状態をリセット

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
        
        # フォーマットテンプレートの読み込み
        self.load_settings()
        
        # フォントサイズの適用
        self.apply_font_size()
        
        # CTI連携サービスの初期化
        self.cti_service = OneClickService()
    
    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        top_bar = QWidget()
        top_bar.setStyleSheet("background-color: #2C3E50; color: white;")
        top_bar_layout = QHBoxLayout(top_bar)
        
        # わかりやすいモードボタン
        self.easy_mode_btn = QPushButton("わかりやすいモード")
        self.easy_mode_btn.setStyleSheet("""
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
                background-color: #FF6B6B;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FF8787;
            }
            QPushButton:pressed {
                background-color: #FF6B6B;
            }
        """)
        self.easy_mode_btn.clicked.connect(self.start_easy_mode)
        top_bar_layout.addWidget(self.easy_mode_btn)
        
        # ワンクリック取得ボタン
        self.oneclick_btn = QPushButton("ワンクリック取得")
        self.oneclick_btn.setStyleSheet("""
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
                background-color: #27AE60;
            }
            QPushButton:hover {
                background-color: #2ECC71;
            }
            QPushButton:pressed {
                background-color: #27AE60;
            }
        """)
        self.oneclick_btn.clicked.connect(self.fetch_cti_data)
        top_bar_layout.addWidget(self.oneclick_btn)
        
        # 既存のボタン
        self.settings_btn = QPushButton("設定")
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("CTIコピー")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
        
        # クリップボード監視トグルボタン
        self.clipboard_toggle_btn = QPushButton("クリップボード監視")
        self.clipboard_toggle_btn.setCheckable(True)
        
        # スクリーンショット表示ボタン
        self.screenshot_btn = QPushButton("スクリーンショット")
        
        # ボタンのスタイル設定
        button_style = """
            QPushButton {
                color: white;
                border: 1px solid white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #34495E;
            }
            QPushButton:pressed {
                background-color: #2C3E50;
            }
        """
        
        for btn in [self.settings_btn, self.clear_btn, 
                   self.cti_copy_btn, self.spreadsheet_btn,
                   self.clipboard_toggle_btn, self.screenshot_btn]:
            btn.setStyleSheet(button_style)
        
        # ボタンをレイアウトに追加
        top_bar_layout.addWidget(self.settings_btn)
        top_bar_layout.addWidget(self.clear_btn)
        top_bar_layout.addWidget(self.cti_copy_btn)
        top_bar_layout.addWidget(self.spreadsheet_btn)
        top_bar_layout.addWidget(self.clipboard_toggle_btn)
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
        self.mobile_type_combo = FocusWheelComboBox()
        self.mobile_type_combo.addItems(["あり", "なし"])
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
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ"))
        self.furigana_mode_combo = QComboBox()
        self.furigana_mode_combo.addItems(["はい", "いいえ"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        basic_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        basic_layout.addWidget(self.furigana_input)
        
        # 生年月日
        birth_date_group = QGroupBox("生年月日")
        birth_date_layout = QHBoxLayout()
        
        # 元号選択
        self.era_combo = FocusWheelComboBox()
        self.era_combo.addItems(["", "昭和", "平成", "令和"])
        birth_date_layout.addWidget(self.era_combo)
        
        # 年選択
        self.year_combo = FocusWheelComboBox()
        self.year_combo.addItems([""] + [str(i) for i in range(1, 65)])
        birth_date_layout.addWidget(self.year_combo)
        birth_date_layout.addWidget(QLabel("年"))
        
        # 月選択
        self.month_combo = FocusWheelComboBox()
        self.month_combo.addItems([""] + [str(i) for i in range(1, 13)])
        birth_date_layout.addWidget(self.month_combo)
        birth_date_layout.addWidget(QLabel("月"))
        
        # 日選択
        self.day_combo = FocusWheelComboBox()
        self.day_combo.addItems([""] + [str(i) for i in range(1, 32)])
        birth_date_layout.addWidget(self.day_combo)
        birth_date_layout.addWidget(QLabel("日"))
        
        birth_date_group.setLayout(birth_date_layout)
        basic_layout.addWidget(birth_date_group)
        
        # 住所情報セクション
        address_group = QGroupBox("住所情報")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        address_container = QHBoxLayout()
        self.address_input = QLineEdit()
        address_container.addWidget(self.address_input)
        
        # マップアイコンボタン
        self.map_btn = QPushButton()
        self.map_btn.setFixedSize(24, 24)
        self.map_btn.setIcon(QIcon("map.png"))
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
        address_container.addWidget(self.map_btn)
        address_layout.addLayout(address_container)
        
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
        basic_layout.addWidget(address_group)
        
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
        self.list_furigana_mode_combo = QComboBox()
        self.list_furigana_mode_combo.addItems(["はい", "いいえ"])
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
        self.list_address_input = QLineEdit()
        list_layout.addWidget(self.list_address_input)
        
        list_group.setLayout(list_layout)
        basic_layout.addWidget(list_group)
        
        # 受注情報セクション
        order_group = QGroupBox("受注情報")
        order_layout = QVBoxLayout()
        
        # 現状回線
        order_layout.addWidget(QLabel("現状回線"))
        self.current_line_combo = FocusWheelComboBox()
        self.current_line_combo.addItems(["", "アナログ", "光電話", "その他"])
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
        self.judgment_combo = FocusWheelComboBox()
        self.judgment_combo.addItems(["", "はい", "いいえ"])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        basic_layout.addWidget(order_group)
        
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
        self.net_usage_combo = FocusWheelComboBox()
        self.net_usage_combo.addItems(["", "はい", "いいえ"])
        other_layout.addWidget(self.net_usage_combo)
        
        # 家族了承
        other_layout.addWidget(QLabel("家族了承"))
        self.family_approval_combo = FocusWheelComboBox()
        self.family_approval_combo.addItems(["", "はい", "いいえ"])
        other_layout.addWidget(self.family_approval_combo)
        
        # 備考
        other_layout.addWidget(QLabel("備考"))
        self.remarks_input = QTextEdit()
        self.remarks_input.setMaximumHeight(100)
        other_layout.addWidget(self.remarks_input)
        
        other_group.setLayout(other_layout)
        basic_layout.addWidget(other_group)
        
        basic_info_group.setLayout(basic_layout)
        parent_layout.addWidget(basic_info_group)
    
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
        self.settings_btn.clicked.connect(self.show_settings)
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
        self.address_input.textChanged.connect(self.convert_to_half_width)
        self.list_address_input.textChanged.connect(self.convert_to_half_width)
        
        # スクロール選択項目の選択変更時に背景色をリセット
        self.era_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.era_combo))
        self.year_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.year_combo))
        self.month_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.month_combo))
        self.day_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.day_combo))
        self.current_line_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.current_line_combo))
        self.judgment_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.judgment_combo))
        self.net_usage_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.net_usage_combo))
        self.family_approval_combo.currentIndexChanged.connect(lambda: self.reset_combo_style(self.family_approval_combo))
        
        # 携帯電話番号の入力チェック
        self.mobile_type_combo.currentTextChanged.connect(self.check_mobile_input)
        self.mobile_input.textChanged.connect(self.check_mobile_input)
        
        # クリック時の背景色リセット
        self.operator_input.mousePressEvent = lambda e: self.reset_input_style(self.operator_input, e)
        self.contractor_input.mousePressEvent = lambda e: self.reset_input_style(self.contractor_input, e)
        self.mobile_input.mousePressEvent = lambda e: self.reset_input_style(self.mobile_input, e)
        
        # 名前とフリガナのバリデーション用のシグナル
        self.contractor_input.textChanged.connect(self.validate_contractor_name)
        self.furigana_input.textChanged.connect(self.validate_furigana_input)
        self.list_name_input.textChanged.connect(self.validate_list_name)
        self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
        
        # フリガナ自動変換のシグナル
        self.contractor_input.textChanged.connect(self.auto_generate_furigana)
        self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
        
        # ボタンのシグナル接続
        self.area_search_btn.clicked.connect(self.search_service_area)
        
        # マップボタンのシグナル接続
        self.map_btn.clicked.connect(self.open_street_view)
    
    def fetch_cti_data(self):
        """
        CTIメインウィンドウからデータを取得し、
        フォームに反映します。
        """
        try:
            # データ取得
            data = self.cti_service.get_all_fields_data()
            if not data:
                QMessageBox.warning(
                    self,
                    "データ取得エラー",
                    "CTIメインウィンドウからデータを取得できませんでした。\n"
                    "CTIメインウィンドウが開いているか確認してください。"
                )
                return

            # フォームに反映
            self.list_name_input.setText(data.customer_name)
            self.address_input.setText(data.address)
            self.list_address_input.setText(data.address)
            self.list_phone_input.setText(data.phone)
            self.postal_code_input.setText(data.postal_code)
            self.list_postal_code_input.setText(data.postal_code)

            # リスト名と契約者名の確認ダイアログを表示
            reply = QMessageBox.question(
                self,
                "わかりやすい確認",
                "リスト名と契約者名は同じですか？\n\n"
                "「はい」を選ぶと、リスト名が対応者名と契約者名に自動的にコピーされます。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # リスト名を対応者名と契約者名にコピー
                self.operator_input.setText(data.customer_name)
                self.contractor_input.setText(data.customer_name)

            # 成功メッセージ
            self.statusBar().showMessage("データを取得しました", 5000)

        except Exception as e:
            logging.error(f"データ取得中にエラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"データ取得中にエラーが発生しました。\n{str(e)}"
            )

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
            # 入力フィールドの背景色を変更して警告
            self.contractor_input.setStyleSheet("background-color: #FFE4E1;")  # 薄い赤色
        else:
            # 正常な入力の場合は背景色をリセット
            self.contractor_input.setStyleSheet("")
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
        else:
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
        else:
            self.statusBar().clearMessage() 

    def validate_list_furigana(self):
        """リストフリガナのバリデーション"""
        text = self.list_furigana_input.text()
        if not validate_furigana(text):
            QToolTip.showText(
                self.list_furigana_input.mapToGlobal(QPoint(0, 0)),
                "フリガナに数字や不適切な文字を含めることはできません",
                self.list_furigana_input
            )
        else:
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
                self.mobile_type_combo.setCurrentText("あり")
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
            # デフォルト値の設定
            self.format_template = ""
            self.settings = {}
            
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.format_template = settings.get('format_template', "")
                    self.settings = settings
                    
                    # 受注者名を設定
                    order_person = settings.get('order_person', '')
                    if order_person:
                        self.order_person = order_person
                        self.order_person_input.setText(order_person)
                        # 受注者名の入力フィールドを読み取り専用に設定
                        self.order_person_input.setReadOnly(True)
        except Exception as e:
            logging.error(f"設定ファイルの読み込みに失敗しました: {str(e)}")
            # エラーが発生しても以前の設定を保持
            if hasattr(self, 'order_person') and self.order_person:
                self.order_person_input.setText(self.order_person)
                self.order_person_input.setReadOnly(True)

    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            try:
                settings = dialog.get_settings()
                self.format_template = settings['format_template']
                self.settings = settings
                
                # フォントサイズの適用
                self.apply_font_size()
                
                # 受注者名の更新
                order_person = settings.get('order_person', '')
                if order_person:
                    self.order_person = order_person
                    self.order_person_input.setText(order_person)
                    # 受注者名の入力フィールドを読み取り専用に設定
                    self.order_person_input.setReadOnly(True)
                
                # 設定を保存
                with open(self.settings_file, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, ensure_ascii=False, indent=4)
                    
            except Exception as e:
                logging.error(f"設定の保存に失敗しました: {str(e)}")
                QMessageBox.warning(
                    self,
                    "エラー",
                    f"設定の保存に失敗しました: {str(e)}"
                )

    def generate_cti_format(self):
        """CTIフォーマットを生成してクリップボードにコピー"""
        try:
            # 必須項目のチェック
            empty_fields = []
            
            # 各フィールドの未入力チェック
            if not self.operator_input.text():
                empty_fields.append("対応者名")
            if not self.contractor_input.text():
                empty_fields.append("契約者名")
            if not self.furigana_input.text():
                empty_fields.append("フリガナ")
            if not self.postal_code_input.text():
                empty_fields.append("郵便番号")
            if not self.address_input.text():
                empty_fields.append("住所")
            if not self.list_name_input.text():
                empty_fields.append("リスト名")
            if not self.list_furigana_input.text():
                empty_fields.append("リストフリガナ")
            if not self.list_phone_input.text():
                empty_fields.append("電話番号")
            if not self.list_postal_code_input.text():
                empty_fields.append("リスト郵便番号")
            if not self.list_address_input.text():
                empty_fields.append("リスト住所")
            if not self.current_line_combo.currentText():
                empty_fields.append("現状回線")
            if not self.order_person_input.text():
                empty_fields.append("受注者名")
            if not self.judgment_combo.currentText():
                empty_fields.append("提供判定")
            if not self.net_usage_combo.currentText():
                empty_fields.append("ネット利用")
            if not self.family_approval_combo.currentText():
                empty_fields.append("家族了承")
            
            # 未入力項目がある場合は警告メッセージを表示
            if empty_fields:
                message = "以下の項目が未入力です：\n\n" + "\n".join(empty_fields) + "\n\nこのままコピーを続けますか？"
                reply = QMessageBox.warning(
                    self,
                    "未入力項目の確認",
                    message,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                
                if reply == QMessageBox.StandardButton.No:
                    return
            
            # 生年月日の生成
            era = self.era_combo.currentText()
            year = self.year_combo.currentText()
            month = self.month_combo.currentText()
            day = self.day_combo.currentText()
            
            # 和暦から西暦への変換
            if era == "昭和":
                year = str(int(year) + 1925)
            elif era == "平成":
                year = str(int(year) + 1988)
            
            birth_date = f"{year}/{month}/{day}"
            
            # GoogleマップのURL生成
            address = self.address_input.text()
            google_maps_url = f"https://www.google.com/maps/search/?api=1&query={quote(address)}"
            
            # フォーマットテンプレートに値を埋め込む
            formatted_text = self.format_template.format(
                operator=self.operator_input.text(),
                mobile=self.mobile_input.text() if self.mobile_type_combo.currentText() == "あり" else "なし",
                contractor=self.contractor_input.text(),
                furigana=self.furigana_input.text(),
                birth_date=birth_date,
                postal_code=self.postal_code_input.text(),
                address=self.address_input.text(),
                list_name=self.list_name_input.text(),
                list_furigana=self.list_furigana_input.text(),
                list_phone=self.list_phone_input.text(),
                list_postal_code=self.list_postal_code_input.text(),
                list_address=self.list_address_input.text(),
                current_line=self.current_line_combo.currentText(),
                order_date=self.order_date_input.text(),
                order_person=self.order_person_input.text(),
                judgment=self.judgment_combo.currentText(),
                fee=self.fee_input.text(),
                net_usage=self.net_usage_combo.currentText(),
                family_approval=self.family_approval_combo.currentText(),
                remarks=self.remarks_input.toPlainText(),
                google_maps_url=google_maps_url
            )
            
            # プレビューに表示
            self.preview_text.setText(formatted_text)
            
            # クリップボードにコピー
            clipboard = QApplication.clipboard()
            clipboard.setText(formatted_text)
            
            # 成功メッセージを表示
            self.statusBar().showMessage("CTIフォーマットをクリップボードにコピーしました", 5000)
            
        except KeyError as e:
            QMessageBox.warning(
                self,
                "フォーマットエラー",
                f"テンプレートに未定義のプレースホルダーが含まれています: {str(e)}"
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "エラー",
                f"CTIフォーマットの生成中にエラーが発生しました: {str(e)}"
            )

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # 受注者名以外の入力フィールドをクリア
        self.operator_input.clear()
        self.mobile_input.clear()
        self.contractor_input.clear()
        self.furigana_input.clear()
        self.postal_code_input.clear()
        self.address_input.clear()
        self.list_name_input.clear()
        self.list_furigana_input.clear()
        self.list_phone_input.clear()
        self.list_postal_code_input.clear()
        self.list_address_input.clear()
        self.remarks_input.clear()
        
        # コンボボックスを初期値に戻す
        self.mobile_type_combo.setCurrentIndex(0)
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)
        self.furigana_mode_combo.setCurrentIndex(0)
        self.list_furigana_mode_combo.setCurrentIndex(0)
        
        # 生年月日コンボボックスを初期値に戻す
        self.era_combo.setCurrentIndex(0)
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
        
        # プレビューをクリア
        self.preview_text.clear()
        
        # 受注日を今日の日付に更新
        self.order_date_input.setText(datetime.datetime.now().strftime("%Y/%m/%d"))
        
        # 料金認識を初期値に戻す
        self.fee_input.setText("3000円～3500円")
        
        # エリア検索結果をリセット
        self.area_result_label.setText("提供エリア: 未検索")

    def start_easy_mode(self):
        """わかりやすいモードを開始"""
        try:
            # まず、CTIデータを取得
            data = self.cti_service.get_all_fields_data()
            if not data:
                QMessageBox.warning(
                    self,
                    "データ取得エラー",
                    "CTIメインウィンドウからデータを取得できませんでした。\n"
                    "CTIメインウィンドウが開いているか確認してください。",
                    QMessageBox.StandardButton.Ok
                )
                return

            # フォームに反映
            self.list_name_input.setText(data.customer_name)
            self.address_input.setText(data.address)
            self.list_address_input.setText(data.address)
            self.list_phone_input.setText(data.phone)
            self.postal_code_input.setText(data.postal_code)
            self.list_postal_code_input.setText(data.postal_code)

            # リスト名と契約者名の確認ダイアログを表示
            reply = QMessageBox.question(
                self,
                "わかりやすい確認",
                "リスト名と契約者名は同じですか？\n\n"
                "「はい」を選ぶと、リスト名が対応者名と契約者名に自動的にコピーされます。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                # リスト名を対応者名と契約者名にコピー
                self.operator_input.setText(data.customer_name)
                self.contractor_input.setText(data.customer_name)
                
                # フリガナを自動生成
                try:
                    furigana = convert_to_furigana(data.customer_name)
                    self.furigana_input.setText(furigana)
                    self.list_furigana_input.setText(furigana)
                    # フリガナモードを「はい」に設定
                    self.furigana_mode_combo.setCurrentText("はい")
                    self.list_furigana_mode_combo.setCurrentText("はい")
                except Exception as e:
                    logging.error(f"フリガナの自動生成に失敗しました: {str(e)}")
            else:
                # 契約者名が入力された時点でフリガナを自動生成するように設定
                self.contractor_input.textChanged.connect(self.auto_generate_both_furigana)

            # 未入力項目のチェックと案内
            empty_fields = []
            field_positions = {}  # フィールドの位置情報を保存

            # 各フィールドの未入力チェック
            if not self.operator_input.text():
                empty_fields.append("対応者名")
                field_positions["対応者名"] = self.operator_input.mapToGlobal(QPoint(0, 0))
            if not self.contractor_input.text():
                empty_fields.append("契約者名")
                field_positions["契約者名"] = self.contractor_input.mapToGlobal(QPoint(0, 0))
            if not self.furigana_input.text():
                empty_fields.append("フリガナ")
                field_positions["フリガナ"] = self.furigana_input.mapToGlobal(QPoint(0, 0))

            # 携帯電話番号のチェック
            if self.mobile_type_combo.currentText() == "あり" and not self.mobile_input.text():
                empty_fields.append("携帯電話番号")
                field_positions["携帯電話番号"] = self.mobile_input.mapToGlobal(QPoint(0, 0))
                self.mobile_input.setStyleSheet("background-color: #FFE4E1;")

            # 生年月日のチェック
            if self.era_combo.currentIndex() == 0:
                empty_fields.append("生年月日（元号）")
                field_positions["生年月日（元号）"] = self.era_combo.mapToGlobal(QPoint(0, 0))
            if self.year_combo.currentIndex() == 0:
                empty_fields.append("生年月日（年）")
                field_positions["生年月日（年）"] = self.year_combo.mapToGlobal(QPoint(0, 0))
            if self.month_combo.currentIndex() == 0:
                empty_fields.append("生年月日（月）")
                field_positions["生年月日（月）"] = self.month_combo.mapToGlobal(QPoint(0, 0))
            if self.day_combo.currentIndex() == 0:
                empty_fields.append("生年月日（日）")
                field_positions["生年月日（日）"] = self.day_combo.mapToGlobal(QPoint(0, 0))

            # その他の必須項目チェック
            if self.current_line_combo.currentIndex() == 0:
                empty_fields.append("現状回線")
                field_positions["現状回線"] = self.current_line_combo.mapToGlobal(QPoint(0, 0))
            if self.judgment_combo.currentIndex() == 0:
                empty_fields.append("提供判定")
                field_positions["提供判定"] = self.judgment_combo.mapToGlobal(QPoint(0, 0))
            if self.net_usage_combo.currentIndex() == 0:
                empty_fields.append("ネット利用")
                field_positions["ネット利用"] = self.net_usage_combo.mapToGlobal(QPoint(0, 0))
            if self.family_approval_combo.currentIndex() == 0:
                empty_fields.append("家族了承")
                field_positions["家族了承"] = self.family_approval_combo.mapToGlobal(QPoint(0, 0))

            # 未入力項目がある場合は案内メッセージを表示
            if empty_fields:
                message = "以下の項目を入力してください：\n\n"
                for field in empty_fields:
                    message += f"・{field}\n"
                    # 対応するフィールドの背景色を変更
                    if field == "対応者名":
                        self.operator_input.setStyleSheet("background-color: #FFE4E1;")
                    elif field == "契約者名":
                        self.contractor_input.setStyleSheet("background-color: #FFE4E1;")
                    # コンボボックスの背景色変更
                    elif "生年月日" in field:
                        if "元号" in field:
                            self.era_combo.setStyleSheet("background-color: #FFE4E1;")
                        elif "年" in field:
                            self.year_combo.setStyleSheet("background-color: #FFE4E1;")
                        elif "月" in field:
                            self.month_combo.setStyleSheet("background-color: #FFE4E1;")
                        elif "日" in field:
                            self.day_combo.setStyleSheet("background-color: #FFE4E1;")
                    elif field == "現状回線":
                        self.current_line_combo.setStyleSheet("background-color: #FFE4E1;")
                    elif field == "提供判定":
                        self.judgment_combo.setStyleSheet("background-color: #FFE4E1;")
                    elif field == "ネット利用":
                        self.net_usage_combo.setStyleSheet("background-color: #FFE4E1;")
                    elif field == "家族了承":
                        self.family_approval_combo.setStyleSheet("background-color: #FFE4E1;")

                QMessageBox.information(
                    self,
                    "入力案内",
                    message,
                    QMessageBox.StandardButton.Ok
                )

            # 成功メッセージ
            self.statusBar().showMessage("わかりやすいモードで入力をサポートしています", 5000)

        except Exception as e:
            logging.error(f"わかりやすいモード実行中にエラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"わかりやすいモードの実行中にエラーが発生しました。\n{str(e)}",
                QMessageBox.StandardButton.Ok
            )

    def reset_field_styles(self):
        """入力フィールドのスタイルをリセット"""
        # テキスト入力フィールドのリセット
        for field in [self.operator_input, self.contractor_input, self.furigana_input]:
            field.setStyleSheet("")
        
        # コンボボックスのリセット
        for combo in [self.era_combo, self.year_combo, self.month_combo, self.day_combo,
                     self.current_line_combo, self.judgment_combo, self.net_usage_combo,
                     self.family_approval_combo]:
            combo.setStyleSheet("")

    def auto_generate_both_furigana(self, text):
        """契約者名からフリガナとリストフリガナを自動生成"""
        if text and self.furigana_mode_combo.currentText() == "はい":
            try:
                furigana = convert_to_furigana(text)
                self.furigana_input.setText(furigana)
                self.list_furigana_input.setText(furigana)
            except Exception as e:
                logging.error(f"フリガナの自動生成に失敗しました: {str(e)}")

    def reset_combo_style(self, combo):
        """コンボボックスの背景色をリセット"""
        combo.setStyleSheet("")

    def check_mobile_input(self):
        """携帯電話番号の入力チェック"""
        if self.mobile_type_combo.currentText() == "あり":
            if not self.mobile_input.text():
                self.mobile_input.setStyleSheet("background-color: #FFE4E1;")
            else:
                self.mobile_input.setStyleSheet("")
        else:
            self.mobile_input.setStyleSheet("")

    def reset_input_style(self, widget, event):
        """入力フィールドの背景色をリセット"""
        widget.setStyleSheet("")
        # 元のmousePressEventを呼び出す
        QLineEdit.mousePressEvent(widget, event)

    def reset_combo_style_on_click(self, combo, event):
        """コンボボックスの背景色をクリック時にリセット"""
        combo.setStyleSheet("")
        # 元のmousePressEventを呼び出す
        FocusWheelComboBox.mousePressEvent(combo, event)