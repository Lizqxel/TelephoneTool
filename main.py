"""
コールセンター業務効率化ツール

このスクリプトは、コールセンター業務の効率化を目的としたGUIアプリケーションです。
PySide6を使用してUIを構築し、Google Spreadsheetsとの連携機能を提供します。

主な機能：
- 顧客情報の入力
- CTIフォーマットの生成
- スプレッドシートへのデータ転記
- クリップボード監視機能
- 提供エリア確認機能
"""

import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QLabel, QLineEdit, QComboBox,
                              QPushButton, QTextEdit, QGroupBox, QMessageBox,
                              QScrollArea, QDialog, QTabWidget)
from PySide6.QtCore import Qt, QTimer, Signal, QObject
from PySide6.QtGui import QFont, QIntValidator, QClipboard
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import threading
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import logging

# 自作モジュールのインポート
from area_checker import AreaChecker

# デフォルトのフォーマットテンプレート
DEFAULT_FORMAT = """対応者（お客様の名前）：{operator}
工事希望日
★出やすい時間帯：携帯：{mobile}
★電話取次：アナログ→光電話
★電話OP：
★無線
契約者(書類名義)：{contractor}
フリガナ：{furigana}
生年月日：{birth_date}
郵便番号：{postal_code}
住所：{address}
リスト名：{list_name}
リスト名フリガナ：{list_furigana}
電話番号：{list_phone}
リスト郵便番号：{list_postal_code}
リスト住所：{list_address}
現状回線：{current_line}
受注日：{order_date}
受注者：{order_person}
提供判定：{judgment}

料金認識：{fee}
ネット利用：{net_usage}
家族了承：{family_approval}

他番号：なし
電話機：プッシュホン
禁止回線：なし
ND：

備考：{remarks}
お客様が今使っている回線：アナログ
案内料金：2500円
※リスト名との関係性："""

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # レイアウトの設定
        layout = QVBoxLayout(self)
        
        # タブウィジェットの作成
        tab_widget = QTabWidget()
        
        # スプレッドシート設定タブ
        spreadsheet_tab = QWidget()
        spreadsheet_layout = QVBoxLayout(spreadsheet_tab)
        
        # スプレッドシートID入力
        spreadsheet_layout.addWidget(QLabel("スプレッドシートID:"))
        self.spreadsheet_id_input = QLineEdit()
        spreadsheet_layout.addWidget(self.spreadsheet_id_input)
        
        # 受注者名設定
        spreadsheet_layout.addWidget(QLabel("受注者名:"))
        self.order_person_input = QLineEdit()
        spreadsheet_layout.addWidget(self.order_person_input)
        
        spreadsheet_layout.addStretch()
        
        # フォーマット設定タブ
        format_tab = QWidget()
        format_layout = QVBoxLayout(format_tab)
        
        # フォーマットテンプレート編集
        format_layout.addWidget(QLabel("CTIフォーマットテンプレート:"))
        self.format_template_input = QTextEdit()
        format_layout.addWidget(self.format_template_input)
        
        # プレースホルダーの説明
        placeholders_label = QLabel(
            "利用可能なプレースホルダー:\n"
            "{operator} - 対応者名\n"
            "{mobile} - 携帯電話番号\n"
            "{contractor} - 契約者名\n"
            "{furigana} - フリガナ\n"
            "{birth_date} - 生年月日\n"
            "{postal_code} - 郵便番号\n"
            "{address} - 住所\n"
            "{list_name} - リスト名\n"
            "{list_furigana} - リストフリガナ\n"
            "{list_phone} - 電話番号\n"
            "{list_postal_code} - リスト郵便番号\n"
            "{list_address} - リスト住所\n"
            "{current_line} - 現状回線\n"
            "{order_date} - 受注日\n"
            "{order_person} - 受注者名\n"
            "{judgment} - 提供判定\n"
            "{fee} - 料金認識\n"
            "{net_usage} - ネット利用\n"
            "{family_approval} - 家族了承\n"
            "{remarks} - 備考"
        )
        placeholders_label.setStyleSheet("color: gray;")
        format_layout.addWidget(placeholders_label)
        
        # タブの追加
        tab_widget.addTab(spreadsheet_tab, "スプレッドシート設定")
        tab_widget.addTab(format_tab, "フォーマット設定")
        layout.addWidget(tab_widget)
        
        # ボタンの配置
        button_layout = QHBoxLayout()
        save_button = QPushButton("保存")
        save_button.clicked.connect(self.save_settings)
        cancel_button = QPushButton("キャンセル")
        cancel_button.clicked.connect(self.reject)
        reset_button = QPushButton("初期値に戻す")
        reset_button.clicked.connect(self.reset_to_default)
        button_layout.addWidget(reset_button)
        button_layout.addStretch()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        # 設定の読み込み
        self.load_settings()

    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.spreadsheet_id_input.setText(settings.get('spreadsheet_id', ''))
                    self.format_template_input.setText(settings.get('format_template', DEFAULT_FORMAT))
                    self.order_person_input.setText(settings.get('order_person', ''))
            else:
                self.reset_to_default()
        except Exception as e:
            print(f"設定の読み込みエラー: {str(e)}")
            self.reset_to_default()

    def save_settings(self):
        """設定をファイルに保存"""
        try:
            settings = {
                'spreadsheet_id': self.spreadsheet_id_input.text(),
                'format_template': self.format_template_input.toPlainText(),
                'order_person': self.order_person_input.text()
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.accept()
        except Exception as e:
            print(f"設定の保存エラー: {str(e)}")

    def reset_to_default(self):
        """設定を初期値に戻す"""
        self.spreadsheet_id_input.setText("1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw")
        self.format_template_input.setText(DEFAULT_FORMAT)
        self.order_person_input.clear()

    def get_settings(self):
        """現在の設定を取得"""
        return {
            'spreadsheet_id': self.spreadsheet_id_input.text(),
            'format_template': self.format_template_input.toPlainText(),
            'order_person': self.order_person_input.text()
        }

class AreaCheckSignals(QObject):
    """エリア確認用のシグナルクラス"""
    result_signal = Signal(dict)

class MainWindow(QMainWindow):
    """
    メインウィンドウクラス
    
    アプリケーションのメインウィンドウとUIコンポーネントを管理します。
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(1000, 800)
        
        # エリア確認用のシグナルを初期化
        self.area_signals = AreaCheckSignals()
        self.area_signals.result_signal.connect(
            lambda result: self.show_area_result(result['status'], result['message'])
        )
        
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
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        content_layout.addWidget(scroll_area, 70)
        
        # プレビューエリア（右側30%）
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.create_preview_area(preview_layout)
        content_layout.addWidget(preview_widget, 30)
        
        main_layout.addLayout(content_layout)

        # 全てのUIコンポーネントを作成した後にシグナルを接続
        self.setup_signals()

        # Google Sheets APIの設定
        self.CREDENTIALS_PATH = "telephonetool-42f9d38d14ee.json"
        self.SPREADSHEET_ID = "1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw"
        
        try:
            self.setup_google_sheets()
        except Exception as e:
            QMessageBox.warning(self, "警告", "スプレッドシートへの初期接続に失敗しました。\n転記機能は再接続後に使用可能になります。")
            print(f"Google Sheets設定エラー: {str(e)}")

    def create_top_bar(self, parent_layout):
        """トップバーを作成"""
        top_bar = QWidget()
        top_bar.setStyleSheet("background-color: #2C3E50; color: white;")
        top_bar_layout = QHBoxLayout(top_bar)
        
        # クリップボード監視ボタンを追加
        self.clipboard_monitor_btn = QPushButton("クリップボード監視")
        self.clipboard_monitor_btn.setCheckable(True)  # トグルボタンにする
        self.clipboard_monitor_btn.setStyleSheet("""
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
        self.clipboard_monitor_btn.clicked.connect(self.toggle_clipboard_monitor)
        top_bar_layout.addWidget(self.clipboard_monitor_btn)
        
        # 既存のボタン
        self.settings_btn = QPushButton("設定")
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("CTIコピー")
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
        
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
        """
        
        self.settings_btn.setStyleSheet(button_style)
        self.clear_btn.setStyleSheet(button_style)
        self.cti_copy_btn.setStyleSheet(button_style)
        self.spreadsheet_btn.setStyleSheet(button_style)
        
        # 設定ボタンのシグナル接続
        self.settings_btn.clicked.connect(self.show_settings)
        
        # ボタンをレイアウトに追加
        top_bar_layout.addWidget(self.settings_btn)
        top_bar_layout.addWidget(self.clear_btn)
        top_bar_layout.addWidget(self.cti_copy_btn)
        top_bar_layout.addWidget(self.spreadsheet_btn)
        
        parent_layout.addWidget(top_bar)

    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        form_group = QGroupBox("顧客情報入力")
        form_layout = QVBoxLayout()
        
        # 入力フィールドのグリッドレイアウト
        grid_layout = QVBoxLayout()
        
        # 顧客情報入力セクション
        customer_info_layout = QVBoxLayout()
        customer_info_layout.setSpacing(10)
        
        # 契約者名
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("契約者名:"))
        self.contractor_input = QLineEdit()
        self.contractor_input.setPlaceholderText("例: 山田 太郎")
        name_layout.addWidget(self.contractor_input)
        customer_info_layout.addLayout(name_layout)
        
        # フリガナ
        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガナ:"))
        self.furigana_input = QLineEdit()
        self.furigana_input.setPlaceholderText("例: ヤマダ タロウ")
        furigana_layout.addWidget(self.furigana_input)
        customer_info_layout.addLayout(furigana_layout)
        
        # 生年月日
        birth_layout = QHBoxLayout()
        birth_layout.addWidget(QLabel("生年月日:"))
        
        birth_date_layout = QHBoxLayout()
        self.year_combo = QComboBox()
        current_year = datetime.datetime.now().year
        for year in range(current_year - 100, current_year - 17):
            self.year_combo.addItem(str(year))
        birth_date_layout.addWidget(self.year_combo)
        birth_date_layout.addWidget(QLabel("年"))
        
        self.month_combo = QComboBox()
        for month in range(1, 13):
            self.month_combo.addItem(str(month))
        birth_date_layout.addWidget(self.month_combo)
        birth_date_layout.addWidget(QLabel("月"))
        
        self.day_combo = QComboBox()
        for day in range(1, 32):
            self.day_combo.addItem(str(day))
        birth_date_layout.addWidget(self.day_combo)
        birth_date_layout.addWidget(QLabel("日"))
        
        birth_layout.addLayout(birth_date_layout)
        customer_info_layout.addLayout(birth_layout)
        
        # 電話番号
        phone_layout = QHBoxLayout()
        phone_layout.addWidget(QLabel("電話番号:"))
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("例: 090-1234-5678")
        phone_layout.addWidget(self.phone_input)
        customer_info_layout.addLayout(phone_layout)
        
        # 携帯電話
        mobile_layout = QHBoxLayout()
        mobile_layout.addWidget(QLabel("携帯電話:"))
        self.mobile_input = QLineEdit()
        self.mobile_input.setPlaceholderText("例: 090-1234-5678")
        mobile_layout.addWidget(self.mobile_input)
        customer_info_layout.addLayout(mobile_layout)
        
        # 郵便番号
        postal_layout = QHBoxLayout()
        postal_layout.addWidget(QLabel("郵便番号:"))
        self.postal_code_input = QLineEdit()
        self.postal_code_input.setPlaceholderText("例: 123-4567")
        postal_layout.addWidget(self.postal_code_input)
        customer_info_layout.addLayout(postal_layout)
        
        # 住所
        address_layout = QHBoxLayout()
        address_layout.addWidget(QLabel("住所:"))
        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText("例: 東京都新宿区西新宿2-8-1")
        address_layout.addWidget(self.address_input)
        customer_info_layout.addLayout(address_layout)
        
        # エリア確認セクション
        area_check_group = QGroupBox("提供エリア確認")
        area_check_layout = QVBoxLayout()
        
        # 都道府県選択
        prefecture_layout = QHBoxLayout()
        prefecture_layout.addWidget(QLabel("都道府県:"))
        self.prefecture_combo = QComboBox()
        # NTT西日本エリアの都道府県を追加
        prefectures = [
            "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
            "鳥取県", "島根県", "岡山県", "広島県", "山口県",
            "徳島県", "香川県", "愛媛県", "高知県",
            "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
        ]
        for pref in prefectures:
            self.prefecture_combo.addItem(pref)
        prefecture_layout.addWidget(self.prefecture_combo)
        area_check_layout.addLayout(prefecture_layout)
        
        # 市区町村
        city_layout = QHBoxLayout()
        city_layout.addWidget(QLabel("市区町村:"))
        self.city_input = QLineEdit()
        self.city_input.setPlaceholderText("例: 山口市")
        city_layout.addWidget(self.city_input)
        area_check_layout.addLayout(city_layout)
        
        # 町名
        town_layout = QHBoxLayout()
        town_layout.addWidget(QLabel("町名:"))
        self.town_input = QLineEdit()
        self.town_input.setPlaceholderText("例: 小郡黄金町")
        town_layout.addWidget(self.town_input)
        area_check_layout.addLayout(town_layout)
        
        # 番地
        block_layout = QHBoxLayout()
        block_layout.addWidget(QLabel("番地:"))
        self.block_input = QLineEdit()
        self.block_input.setPlaceholderText("例: 1-1")
        block_layout.addWidget(self.block_input)
        area_check_layout.addLayout(block_layout)
        
        # エリア確認ボタン
        button_layout = QHBoxLayout()
        self.area_check_btn = QPushButton("エリア確認")
        self.area_check_btn.clicked.connect(self.check_area)
        button_layout.addWidget(self.area_check_btn)
        area_check_layout.addLayout(button_layout)
        
        # エリア確認結果表示ラベル
        self.area_result_label = QLabel()
        self.area_result_label.setWordWrap(True)
        self.area_result_label.setMinimumHeight(80)
        self.area_result_label.hide()
        area_check_layout.addWidget(self.area_result_label)
        
        # 判定結果選択
        judgment_layout = QHBoxLayout()
        judgment_layout.addWidget(QLabel("判定結果:"))
        self.judgment_combo = QComboBox()
        self.judgment_combo.addItems(["", "OK", "NG"])
        judgment_layout.addWidget(self.judgment_combo)
        area_check_layout.addLayout(judgment_layout)
        
        area_check_group.setLayout(area_check_layout)
        
        # リスト情報セクション
        list_info_group = QGroupBox("リスト情報")
        list_info_layout = QVBoxLayout()
        
        # リスト名
        list_name_layout = QHBoxLayout()
        list_name_layout.addWidget(QLabel("リスト名:"))
        self.list_name_input = QLineEdit()
        list_name_layout.addWidget(self.list_name_input)
        list_info_layout.addLayout(list_name_layout)
        
        # リストフリガナ
        list_furigana_layout = QHBoxLayout()
        list_furigana_layout.addWidget(QLabel("リストフリガナ:"))
        self.list_furigana_input = QLineEdit()
        list_furigana_layout.addWidget(self.list_furigana_input)
        list_info_layout.addLayout(list_furigana_layout)
        
        # リスト電話番号
        list_phone_layout = QHBoxLayout()
        list_phone_layout.addWidget(QLabel("リスト電話番号:"))
        self.list_phone_input = QLineEdit()
        list_phone_layout.addWidget(self.list_phone_input)
        list_info_layout.addLayout(list_phone_layout)
        
        # リスト郵便番号
        list_postal_layout = QHBoxLayout()
        list_postal_layout.addWidget(QLabel("リスト郵便番号:"))
        self.list_postal_code_input = QLineEdit()
        list_postal_layout.addWidget(self.list_postal_code_input)
        list_info_layout.addLayout(list_postal_layout)
        
        # リスト住所
        list_address_layout = QHBoxLayout()
        list_address_layout.addWidget(QLabel("リスト住所:"))
        self.list_address_input = QLineEdit()
        list_address_layout.addWidget(self.list_address_input)
        list_info_layout.addLayout(list_address_layout)
        
        list_info_group.setLayout(list_info_layout)
        
        # 各セクションをメインレイアウトに追加
        grid_layout.addWidget(QLabel("【基本情報】"))
        grid_layout.addLayout(customer_info_layout)
        grid_layout.addWidget(area_check_group)
        grid_layout.addWidget(list_info_group)
        
        form_layout.addLayout(grid_layout)
        form_group.setLayout(form_layout)
        parent_layout.addWidget(form_group)

    def create_preview_area(self, parent_layout):
        """プレビューエリアを作成"""
        preview_label = QLabel("CTIフォーマットプレビュー")
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        
        parent_layout.addWidget(preview_label)
        parent_layout.addWidget(self.preview_text)

    def format_phone_number(self):
        """電話番号の自動フォーマット処理"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 11桁以内に制限
        text = text[:11]
        
        # ハイフン挿入
        if len(text) > 7:
            text = f'{text[:3]}-{text[3:7]}-{text[7:]}'
        elif len(text) > 3:
            text = f'{text[:3]}-{text[3:]}'
        
        sender.setText(text)

    def format_postal_code(self):
        """郵便番号の自動フォーマット処理"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 7桁以内に制限
        text = text[:7]
        
        # ハイフン挿入
        if len(text) > 3:
            text = f'{text[:3]}-{text[3:]}'
        
        sender.setText(text)

    def convert_to_half_width(self):
        """全角文字を半角に変換"""
        sender = self.sender()
        text = sender.text()
        
        # 全角数字を半角に変換
        for i in range(10):
            text = text.replace(chr(0xFF10 + i), str(i))
        
        # 全角ハイフン（マイナス）を半角に変換
        text = text.replace('−', '-')  # 全角マイナス
        text = text.replace('ー', '-')  # 長音符
        text = text.replace('―', '-')  # ダッシュ
        text = text.replace('‐', '-')  # ハイフン
        text = text.replace('－', '-')  # 全角ハイフン
        
        sender.setText(text)

    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.SPREADSHEET_ID = settings.get('spreadsheet_id', '1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw')
                    self.format_template = settings.get('format_template', '')
                    self.order_person = settings.get('order_person', '')
            else:
                self.SPREADSHEET_ID = '1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw'
                self.format_template = ''
                self.order_person = ''
        except Exception as e:
            print(f"設定の読み込みエラー: {str(e)}")
            self.SPREADSHEET_ID = '1Y3M8YZ0ywLdMxCVY6EB8COcPkBaOS6e27jzEJWcU8Tw'
            self.format_template = ''
            self.order_person = ''

    def generate_cti_format(self):
        """CTIフォーマットの生成とプレビュー表示"""
        # 生年月日の生成（和暦を西暦に変換）
        era = self.era_combo.currentText()
        year = int(self.year_combo.currentText())
        month = self.month_combo.currentText()
        day = self.day_combo.currentText()
        
        # 和暦を西暦に変換
        if era == "昭和":
            year = year + 1925  # 昭和元年は1926年
        elif era == "平成":
            year = year + 1988  # 平成元年は1989年
        # 西暦の場合はそのまま
        
        birth_date = f"{year}/{month}/{day}"
        
        # 現在の日付を取得（先頭の0を除去）
        now = datetime.datetime.now()
        month_str = str(now.month)
        day_str = str(now.day)
        order_date = f"{month_str}/{day_str}"
        
        # フォーマットデータの準備
        format_data = {
            'operator': self.operator_input.text(),
            'mobile': "なし" if self.mobile_type_combo.currentText() == "なし" else self.mobile_input.text(),
            'contractor': self.contractor_input.text(),
            'furigana': self.furigana_input.text(),
            'birth_date': birth_date,
            'postal_code': self.postal_code_input.text(),
            'address': self.address_input.text(),
            'list_name': self.list_name_input.text(),
            'list_furigana': self.list_furigana_input.text(),
            'list_phone': self.list_phone_input.text(),
            'list_postal_code': self.list_postal_code_input.text(),
            'list_address': self.list_address_input.text(),
            'current_line': self.current_line_combo.currentText(),
            'order_date': order_date,
            'order_person': self.order_person_input.text(),
            'judgment': self.judgment_combo.currentText(),
            'fee': self.fee_input.text(),
            'net_usage': self.net_usage_combo.currentText(),
            'family_approval': self.family_approval_combo.currentText(),
            'remarks': self.remarks_input.toPlainText()
        }
        
        # テンプレートが設定されている場合はそれを使用し、
        # 設定されていない場合はデフォルトのフォーマットを使用
        if self.format_template:
            cti_text = self.format_template.format(**format_data)
        else:
            cti_text = DEFAULT_FORMAT.format(**format_data)
        
        # プレビューに表示
        self.preview_text.setText(cti_text)
        
        # クリップボードにコピー
        clipboard = QApplication.clipboard()
        clipboard.setText(cti_text)

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        self.mobile_type_combo.setCurrentIndex(0)  # 携帯電話番号の選択を「入力」に戻す
        self.mobile_input.setEnabled(True)  # 入力フィールドを有効化
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
        self.order_person_input.clear()
        
        # コンボボックスを初期値に戻す
        self.era_combo.setCurrentIndex(0)
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)
        
        # 料金認識を初期値に戻す
        self.fee_input.setText("3000円～3500円")
        
        # プレビューエリアをクリア
        self.preview_text.clear()
        
        # 備考欄をクリア
        self.remarks_input.clear()

    def setup_google_sheets(self):
        """Google Sheets APIの設定"""
        try:
            # 認証情報を直接指定
            creds_dict = {
                "type": "service_account",
                "project_id": "telephonetool",
                "private_key_id": "42f9d38d14eec753b9b97f60e25d38993f93cd9c",
                "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCSaSf2fczAQh8Q\nqXtEZgawNuBvjmIvrU2ClEQ3sWDJVql8hNSXG0NCYGksvy3IMRX+71m0BpAn4c5L\n45NsPX7nm8g1NTJG13jvdeS3KdbpBbYK3zN6mVB2O4hxrj5eja78Ik3ih6zVl8Zw\njMYIrgABSclfDre3DZNDAw+brdHC5ltOvnJCPUF3w/PPMEprry74RDz37rdS/a8p\n9KbdTWTl5n3xsUB1oZxiussDCICJjyj0b+pKwM46I6WHs1ZtcGOOazVxK414kAAY\n9y1Xq9iyRmWVdSkbWqs2S8FNbQhbz6TwkZv96aNR8OXDHeLw6+mXbnRGNaZBHEKj\nRhXAeffHAgMBAAECggEAQBl6zg94l4m7WQuidKkob3ivHRgcw5vfrfYkwa9OXQes\nj2AGRRvCACr+kQEoVZer9h+rScZ/0X4qWA5MKlzoFRWeezENkHdgspIObuSJ+x4t\ne6gJvTinQgRBcefj1Xi5bhjEuZNF54OZ9Qek4gLv7KB14cCrTSDL4tBRwopApk+T\njBxV0NOGljN6YoPidanrNzkEc2FMsSQ81icBNLzuoj+GOFPIO3M+hIQdfpTZcjgR\n0p6Q98vfU7vlHD6PvRA6keIDVPpHsPXMOIRkprttsOk7ojquexsBqWe0V0cL3grM\nHiUXVvkNziSZbomSVqUdG1Ntp2BsNvKqMm46pb62FQKBgQDEGDeAQbw8szfmUBI5\ndngtamCzHFBsZO7mq21vCq/FQ3Dj5OgkFbVMhfocpwHQb4ZhkZL8lfQ9PRQ8/cVu\ngGs91DsmlZml+ZxVrnXMZKzdaOmrv6OyJDKN2yCpd1g7azscEwFWpVu/P7E7axZb\neJKVAt0MwsvNCEJIaA8cuMBAxQKBgQC/I10B2KnQ7YRn1706vXTrEa5Pxnvf7BBz\nquCbNrnXVZEHY2XS+662o6GwlFAruz/OXJUwWg6u4+pcP0oIPtGts1SNZrcJdMIg\nyLEFjt2ysnGqeLd/zyrkPK6xeNvV5HQXlLdppTE/0iXm433+pFOtCjqlJ28UoI9F\nSdR3dBLHGwKBgQCIRGPdJtEORWRlEeN4NxFQTgogrV5d1M4HUb1cWsrGhBUg6ONA\noC06nieuXYfvNnDlwGmqSPJO0/EKaTcXkPn1H1RzfaYmJo0zJWcKwDM4MT2gci3p\nDypqVYoe+aZAtEWBPtvBQGu/PR2GMuZ4bhM+pZzCz2Mcec7FzjoiNWi0GQKBgQCZ\njOZF+nIJ9xXcenN5ggQwaCbZzcFsVW+uDIOeDavkcsgs4ExH34svDGtzuOJjD22l\n8bikfGS5WT3IV8u4rgayfZOaeP7oaNUfkzqrFWfDDBnGcm4wDhUOADXzOv2Yaoxc\n+UsTYvMaq09pmi546DiUldghH3ncX1RZvIMkZ6pCKwKBgDL/BSgXS9rLHO3r7KKc\nK4/sSswXcHTVAsu7idaINPWw3pnyyqxBCf+5LjzKFy+VY2DCd8ZqXdU8LoLhXDPZ\ny7f4pze6vxhb93O+VooaS79Yjipk8+/rCsHxajbCz29D8oL+2KYMwfsCRuTzotKU\nSYKjdETtL/oYytV4+2iCLFo5",
                "client_email": "telephonetool@telephonetool.iam.gserviceaccount.com",
                "client_id": "100938209331537488837",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/telephonetool%40telephonetool.iam.gserviceaccount.com",
                "universe_domain": "googleapis.com"
            }
            
            scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']
            
            credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            self.gc = gspread.authorize(credentials)
            self.sheet = self.gc.open_by_key(self.SPREADSHEET_ID).sheet1
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"スプレッドシートへの接続に失敗しました。\n{str(e)}")

    def write_to_spreadsheet(self):
        """スプレッドシートにデータを書き込む"""
        try:
            # スプレッドシートの接続確認と再接続
            if not hasattr(self, 'sheet') or self.sheet is None:
                try:
                    self.setup_google_sheets()
                except Exception as e:
                    QMessageBox.critical(self, "接続エラー", "スプレッドシートへの接続に失敗しました。\n再度お試しください。")
                    return

            # 現在時刻と日付を取得
            current_time = datetime.datetime.now().strftime("%H:%M")
            current_date = datetime.datetime.now().strftime("%Y/%m/%d")
            
            # 電話番号からハイフンを除去
            tel1 = self.list_phone_input.text().replace('-', '')
            tel2 = self.mobile_input.text().replace('-', '')
            
            # 空のリストを作成（スプレッドシートの列数分）
            row_data = [''] * 16  # 画像から見える列数に合わせて調整
            
            # 必要な列にデータを設定（0から始まるインデックスなので、A列が0、B列が1）
            row_data[3] = tel1                          # TEL1（E列）
            row_data[4] = tel2                          # TEL2（F列）
            row_data[7] = current_time                  # 架電時間（I列）
            row_data[10] = current_date                 # トス日（L列）
            
            # スプレッドシートに追記
            self.sheet.append_row(row_data, value_input_option='RAW')
            
            # 成功メッセージを表示
            QMessageBox.information(self, "成功", "スプレッドシートへの転記が完了しました。")
            
        except gspread.exceptions.APIError as e:
            # API制限やトークン期限切れの場合は再接続を試みる
            try:
                self.setup_google_sheets()
                # 再度書き込みを試行
                self.sheet.append_row(row_data, value_input_option='RAW')
                QMessageBox.information(self, "成功", "スプレッドシートへの転記が完了しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"スプレッドシートへの転記に失敗しました。\n{str(e)}")
        except Exception as e:
            # その他のエラー
            QMessageBox.critical(self, "エラー", f"スプレッドシートへの転記に失敗しました。\n{str(e)}")

    def update_year_combo(self):
        """元号に応じて年の選択肢を更新"""
        current_era = self.era_combo.currentText()
        current_year = datetime.datetime.now().year
        
        self.year_combo.clear()
        
        if current_era == "昭和":
            # 昭和1年から64年まで
            self.year_combo.addItems([str(i) for i in range(1, 65)])
        elif current_era == "平成":
            # 平成1年から31年まで
            self.year_combo.addItems([str(i) for i in range(1, 32)])
        else:  # 西暦
            # 1900年から現在の年まで
            self.year_combo.addItems([str(i) for i in range(1900, current_year + 1)])
        
        # 最初の項目を選択
        self.year_combo.setCurrentIndex(0)

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
        self.area_check_btn.clicked.connect(self.check_area)  # エリア確認ボタンのシグナル接続
        
        # 設定ボタンのシグナル接続
        self.settings_btn.clicked.connect(self.show_settings)

    def format_phone_number_without_hyphen(self):
        """電話番号の自動フォーマット処理（ハイフンなし）"""
        sender = self.sender()
        text = sender.text().replace('-', '')  # ハイフンを除去
        
        # 数字以外を除去
        text = ''.join(filter(str.isdigit, text))
        
        # 11桁以内に制限
        text = text[:11]
        
        sender.setText(text)

    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # 設定が保存された場合
            settings = dialog.get_settings()
            self.SPREADSHEET_ID = settings['spreadsheet_id']
            self.format_template = settings['format_template']
            self.order_person = settings['order_person']
            
            # 受注者名を入力欄に設定
            self.order_person_input.setText(self.order_person)
            
            # スプレッドシートの再接続
            try:
                self.setup_google_sheets()
                QMessageBox.information(self, "成功", "設定を保存し、スプレッドシートに再接続しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"スプレッドシートへの接続に失敗しました。\n{str(e)}")

    def toggle_mobile_input(self, text):
        """携帯電話番号入力フィールドの有効/無効を切り替え"""
        if text == "なし":
            self.mobile_input.clear()
            self.mobile_input.setEnabled(False)
        else:
            self.mobile_input.setEnabled(True)

    def toggle_clipboard_monitor(self):
        """クリップボード監視の開始/停止を切り替え"""
        if self.clipboard_monitor_btn.isChecked():
            self.clipboard_timer.start(1000)  # 1秒ごとにチェック
            QMessageBox.information(self, "クリップボード監視", "クリップボード監視を開始しました。\n他のアプリからコピーした情報を自動で取得します。")
        else:
            self.clipboard_timer.stop()
            QMessageBox.information(self, "クリップボード監視", "クリップボード監視を停止しました。")

    def check_clipboard(self):
        """クリップボードの内容をチェックして適切なフィールドに自動入力"""
        current_text = self.clipboard.text()
        
        # クリップボードの内容が変更された場合のみ処理
        if current_text != self.last_clipboard_text:
            self.last_clipboard_text = current_text
            
            # クリップボードの内容を解析
            self.analyze_clipboard_content(current_text)

    def analyze_clipboard_content(self, text):
        """クリップボードの内容を解析して適切なフィールドに入力"""
        # 郵便番号のパターンマッチ（例：123-4567）
        if len(text) == 8 and text[3] == '-' and text.replace('-', '').isdigit():
            self.postal_code_input.setText(text)  # 通常の郵便番号フィールド
            self.list_postal_code_input.setText(text)  # リスト郵便番号フィールド
            return
            
        # 電話番号のパターンマッチ（例：03-1234-5678）
        if len(text) <= 13 and text.replace('-', '').isdigit():
            if len(text.replace('-', '')) == 11:  # 携帯電話の場合
                self.mobile_input.setText(text)
            else:
                self.list_phone_input.setText(text)
            return
            
        # 住所らしき文字列（漢字とカタカナが含まれる長い文字列）
        if len(text) > 10 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            self.address_input.setText(text)  # 通常の住所フィールド
            self.list_address_input.setText(text)  # リスト住所フィールド
            return
            
        # カタカナのみの文字列（フリガナとして扱う）
        if all(ord(c) >= 0x30A0 and ord(c) <= 0x30FF for c in text.replace(' ', '')):
            self.list_furigana_input.setText(text)  # リストフリガナに入力
            return
            
        # その他の文字列（名前として扱う）
        if len(text) <= 20 and any(ord(c) >= 0x4E00 and ord(c) <= 0x9FFF for c in text):
            self.list_name_input.setText(text)  # リスト名に入力
            return

    def check_area(self):
        """提供エリア確認機能"""
        # 入力値の取得
        prefecture = self.prefecture_combo.currentText()
        city = self.city_input.text().strip()
        town = self.town_input.text().strip()
        block = self.block_input.text().strip()
        
        # 入力チェック
        if not prefecture or not city:
            QMessageBox.warning(self, "入力エラー", "都道府県と市区町村は必須入力です。")
            return
        
        # 確認中メッセージを表示
        self.area_result_label.setText("エリア確認中...")
        self.area_result_label.setStyleSheet("color: blue;")
        
        # ログ出力を標準出力に表示するための設定
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler()  # 標準出力にログを表示
            ]
        )
        
        # エリア確認処理を別スレッドで実行
        self.area_check_thread = threading.Thread(
            target=self.check_area_thread,
            args=(prefecture, city, town, block)
        )
        self.area_check_thread.daemon = True
        self.area_check_thread.start()
    
    def check_area_thread(self, prefecture, city, town, block):
        """エリア確認処理を実行するスレッド"""
        try:
            # エリア確認クラスのインスタンス化
            area_checker = AreaChecker()
            
            # エリア確認実行
            result = area_checker.check_area(prefecture, city, town, block)
            
            # 結果をシグナルで送信
            self.area_check_signals.result_signal.emit(result)
        except Exception as e:
            # エラーが発生した場合
            error_result = {
                'status': 'ERROR',
                'message': f'エラーが発生しました: {str(e)}'
            }
            self.area_check_signals.result_signal.emit(error_result)

    def show_area_result(self, status, message):
        """
        エリア確認結果を表示
        
        Args:
            status (str): 結果のステータス（'OK', 'NG', 'ERROR'）
            message (str): 表示するメッセージ
        """
        # ボタンを再有効化
        self.area_check_btn.setEnabled(True)
        self.area_check_btn.setText("エリア確認")
        
        # 結果に応じてスタイルを設定
        if status == 'OK':
            style = "background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb;"
            self.judgment_combo.setCurrentText("OK")
            title = "提供可能"
        elif status == 'NG':
            style = "background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;"
            self.judgment_combo.setCurrentText("NG")
            title = "提供不可"
        else:  # ERROR
            style = "background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba;"
            title = "エラー"
        
        # メッセージを整形
        if status == 'ERROR':
            formatted_message = f"{title}:\n{message}\n\n再度お試しいただくか、別の郵便番号で確認してください。"
        else:
            formatted_message = f"{title}:\n{message}"
        
        self.area_result_label.setStyleSheet(f"QLabel {{ padding: 10px; border-radius: 3px; {style} }}")
        self.area_result_label.setText(formatted_message)
        self.area_result_label.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 