"""
メインウィンドウモジュール

このモジュールは、アプリケーションのメインウィンドウを提供します。
"""

import sys
import logging
import datetime
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
                              QSizePolicy, QProgressBar, QListView)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices

from version import VERSION, GITHUB_OWNER, GITHUB_REPO, APP_NAME

from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
import time
from typing import Dict, Any, List, Optional, Union, Tuple

from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, 
                             QTextEdit, QComboBox, QWidget, 
                             QMessageBox, QApplication, QDialog,
                             QStatusBar, QSizePolicy, QSpacerItem,
                             QTabWidget, QRadioButton, QGroupBox,
                             QScrollArea, QSplitter, QToolTip, QMenuBar)
from PySide6.QtCore import Qt, QObject, QTimer, Signal, Slot, QMetaObject, Q_ARG, QPoint, QEvent, QThread
from PySide6.QtGui import QFont, QIntValidator, QCloseEvent, QTextOption, QShowEvent, QIcon

from ui.main_window_functions import MainWindowFunctions
from services.oneclick import OneClickService
from services.phone_button_monitor import PhoneButtonMonitor
from utils.format_utils import format_phone_number, format_phone_number_without_hyphen, format_postal_code
from ui.easy_mode_dialogs import AddressInfoDialog, ListInfoDialog, OrdererInputDialog, OrderInfoDialog
from ui.easy_mode_dialogs import DIALOG_BACK, DIALOG_NEXT, DIALOG_CANCEL
from ui.easy_mode_dialogs import convert_to_half_width
from ui.settings_dialog import SettingsDialog
from ui.mode_selection_dialog import ModeSelectionDialog
from utils.string_utils import validate_name, validate_furigana, convert_to_half_width_except_space
from utils.furigana_utils import convert_to_furigana
from ui.update_dialog import UpdateDialog
from services.area_search import search_service_area


class CustomComboBox(QComboBox):
    """スクロールでの値変更を防止するカスタムコンボボックス"""
    def wheelEvent(self, event):
        """ホイールイベントを無視"""
        event.ignore()

class NoWheelComboBox(QComboBox):
    """スクロールイベントを無視するQComboBox"""
    def wheelEvent(self, event):
        event.ignore()

class MainWindow(QMainWindow, MainWindowFunctions):
    """メインウィンドウクラス"""
    
    def set_font_size(self, size):
        """
        フォントサイズを設定する
        
        Args:
            size (int): 設定するフォントサイズ
        """
        try:
            # フォントサイズを設定
            font = QFont()
            font.setPointSize(size)
            
            # 各ウィジェットにフォントを適用
            self.setFont(font)
            
            # プレビューエリアのフォントサイズを設定
            if hasattr(self, 'preview_text'):
                self.preview_text.setFont(font)
            
            logging.info(f"フォントサイズを {size} に設定しました")
            
        except Exception as e:
            logging.error(f"フォントサイズの設定中にエラーが発生しました: {e}")
    
    def setup_logging(self):
        """
        ログ設定を行う
        """
        try:
            # ログディレクトリの作成
            log_dir = "logs"
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            # ログファイル名の生成（タイムスタンプ付き）
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(log_dir, f"app_{timestamp}.log")
            
            # ログ設定
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler()
                ]
            )
            
            logging.info("ログ設定を完了しました")
            
        except Exception as e:
            print(f"ログ設定中にエラーが発生しました: {e}")
    
    def __init__(self):
        """
        メインウィンドウの初期化
        """
        super().__init__()
        
        # バージョン情報の設定
        self.version = "1.0.0"
        
        # モード変更フラグ（設定ダイアログ用）
        self.mode_changed = False
        self.new_mode = None
        
        # ログ設定
        self.setup_logging()
        
        # 設定ファイルのパスを設定
        if getattr(sys, 'frozen', False):
            # exeファイルとして実行されている場合
            self.settings_file = os.path.join(os.path.dirname(sys.executable), 'settings.json')
        else:
            # 通常のPythonスクリプトとして実行されている場合
            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
        
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # 設定を読み込む
        self.settings = {}
        
        # 設定ファイルが存在しない場合は新規作成
        if not os.path.exists(self.settings_file):
            logging.info("設定ファイルが存在しないため、新規作成します")
            self.save_mode_settings('simple', True)
        
        # 設定を読み込む
        self.load_settings()
        
        # アクティブな検索スレッドを保持するリスト
        self.active_search_threads = []
        
        # モード設定
        self.current_mode = self.settings.get('mode', 'simple')
        logging.info(f"現在のモード: {self.current_mode}")
        
        # ウィンドウの基本設定
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setGeometry(100, 100, 800, 600)
        
        # 生年月日入力用のコンボボックスを初期化
        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["令和", "平成", "昭和", "大正", "明治"])
        
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 151)])
        
        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        
        self.day_combo = NoWheelComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの作成
        self.main_layout = QVBoxLayout(main_widget)
        
        # モード選択ダイアログの表示（設定に基づいて表示を制御）
        if self.settings.get('show_mode_selection', True):
            self.show_mode_selection()
        
        # 選択されたモードに基づいてUIを初期化
        if self.current_mode == 'simple':
            self.init_simple_mode()
        else:
            self.init_easy_mode()
        
        # 電話ボタン監視の初期化
        self.phone_monitor = PhoneButtonMonitor(self)
        self.phone_monitor.start_monitoring()
        
        # フォントサイズの設定
        font_size = self.settings.get('font_size', 10)
        self.set_font_size(font_size)
    
    def check_and_show_mode_selection(self):
        """
        モード選択ダイアログの表示を確認し、必要に応じて表示する
        """
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # モード設定が存在しない場合、または次回以降表示する設定の場合
                    if 'mode' not in settings or settings.get('show_mode_selection', True):
                        self.show_mode_selection_dialog()
                    else:
                        self.current_mode = settings.get('mode', 'simple')
            else:
                # 設定ファイルが存在しない場合は、必ずモード選択ダイアログを表示
                self.show_mode_selection_dialog()
                # 設定ファイルを作成
                self.save_mode_settings('simple', True)
        except Exception as e:
            logging.error(f"モード設定の読み込み中にエラーが発生しました: {e}")
            # エラーが発生した場合は、モード選択ダイアログを表示
            self.show_mode_selection_dialog()
    
    def show_mode_selection(self):
        """
        モード選択ダイアログを表示する
        設定ファイルのshow_mode_selectionの値に基づいて表示を制御する
        """
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # show_mode_selectionがFalseの場合は表示しない
                    if not settings.get('show_mode_selection', True):
                        self.current_mode = settings.get('mode', 'simple')
                        return
            # 設定ファイルが存在しない場合、またはshow_mode_selectionがTrueの場合は表示
            self.show_mode_selection_dialog()
            # 設定ファイルが存在しない場合は作成
            if not os.path.exists(self.settings_file):
                self.save_mode_settings('simple', True)
        except Exception as e:
            logging.error(f"モード設定の読み込み中にエラーが発生しました: {e}")
            # エラーが発生した場合は、モード選択ダイアログを表示
            self.show_mode_selection_dialog()
    
    def show_mode_selection_dialog(self):
        """
        モード選択ダイアログを表示し、選択結果を保存する
        """
        dialog = ModeSelectionDialog(self)
        if dialog.exec():
            # 選択されたモードを保存
            self.current_mode = dialog.get_selected_mode()
            self.save_mode_settings(self.current_mode, dialog.should_show_again())
            logging.info(f"モードを {self.current_mode} に設定しました")
        else:
            # キャンセルされた場合は、デフォルトでシンプルモードを使用
            self.current_mode = 'simple'
            self.save_mode_settings(self.current_mode, True)
            logging.info("モード選択がキャンセルされました。シンプルモードを使用します。")
    
    def save_mode_settings(self, mode, show_again):
        """
        モード設定を保存する
        
        Args:
            mode: 選択されたモード（'simple'または'easy'）
            show_again: 次回から表示するかどうか
        """
        try:
            # 設定ファイルの読み込み
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            
            # モード設定を更新
            settings['mode'] = mode
            settings['show_mode_selection'] = show_again
            
            # 設定ファイルに保存
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            logging.info(f"モード設定を保存しました: mode={mode}, show_mode_selection={show_again}")
        except Exception as e:
            logging.error(f"モード設定の保存中にエラーが発生しました: {e}")
    
    def init_simple_mode(self):
        """通常モードのUIを初期化"""
        logging.info("通常モードの初期化を開始")
        
        # 設定に基づいてウィンドウタイトルを設定
        self.setWindowTitle("コールセンター業務効率化ツール - 通常モード")
        self.setMinimumSize(600, 400)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        
        # 設定ファイルのパスを確認
        if not hasattr(self, 'settings_file'):
            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
            logging.info(f"設定ファイルのパスを設定: {self.settings_file}")
        
        logging.info(f"設定ファイルの存在確認: {os.path.exists(self.settings_file)}")
        
        # 設定を読み込む
        if not hasattr(self, 'settings'):
            self.settings = {}
        
        # format_templateを設定
        if not hasattr(self, 'format_template') or not self.format_template:
            logging.info("format_templateを設定します")
            self.load_settings()
            if hasattr(self, 'settings') and 'format_template' in self.settings:
                self.format_template = self.settings['format_template']
                logging.info(f"format_templateを設定しました: {self.format_template[:100]}...")
            else:
                logging.error("format_templateの設定に失敗しました")
                QMessageBox.warning(self, "エラー", "テンプレートの設定に失敗しました。")
                return
        
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
    
    def init_easy_mode(self):
        """誘導モードのUIを初期化"""
        # 設定に基づいてウィンドウタイトルを設定
        self.setWindowTitle("コールセンター業務効率化ツール - 誘導モード")
        self.setMinimumSize(400, 300)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        
        # プレビューエリア
        preview_group = QGroupBox("プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        self.create_preview_area(preview_layout)
        main_layout.addWidget(preview_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 開始ボタン
        self.start_button = QPushButton("開始")
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }
        """)
        self.start_button.clicked.connect(self.start_easy_mode)
        button_layout.addWidget(self.start_button)
        
        # 設定ボタン
        self.settings_button = QPushButton("設定")
        self.settings_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        self.settings_button.clicked.connect(self.show_settings)
        button_layout.addWidget(self.settings_button)
        
        main_layout.addLayout(button_layout)
        
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
        
        self.init_menu()
    
    def start_easy_mode(self):
        """誘導モードを開始"""
        try:
            logging.info("誘導モードを開始")
            
            # 提供判定結果をリセット
            self.judgment_result_label.setText("提供エリア: 未検索")
            self.judgment_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    background-color: #f8f8f8;
                }
            """)
            
            # CTIデータを取得
            cti_data = self.cti_service.get_all_fields_data()
            if not cti_data:
                QMessageBox.warning(self, "警告", "CTIデータの取得に失敗しました。")
                return
            
            # 顧客名の処理（苗字と名前の間のスペースを全角に）
            customer_name = cti_data.customer_name
            if customer_name:
                customer_name = customer_name.replace(' ', '　')  # 半角スペースを全角に
                customer_name = convert_to_half_width_except_space(customer_name)
            
            # 住所の処理（ハイフンを半角に）
            address = cti_data.address
            if address:
                address = address.replace('－', '-')  # 全角ハイフンを半角に
                address = address.replace('ー', '-')  # 長音記号を半角ハイフンに
                address = address.replace('−', '-')  # 別種の全角ハイフンを半角に
                address = address.replace(' ', '　')  # 半角スペースを全角に
                address = convert_to_half_width_except_space(address)
            
            # データの初期化と設定
            self.address_data = {
                'postal_code': convert_to_half_width(cti_data.postal_code) if cti_data.postal_code else "",
                'address': address if address else ""
            }
            
            # 顧客名のフリガナを取得して設定
            customer_furigana = ""
            if customer_name:
                # フリガナ変換APIを使用
                customer_furigana = convert_to_furigana(customer_name)
            
            self.list_data = {
                'list_name': customer_name if customer_name else "",
                'list_furigana': customer_furigana,  # 自動生成したフリガナを設定
                'list_phone': convert_to_half_width(cti_data.phone) if cti_data.phone else "",
                'list_postal_code': convert_to_half_width(cti_data.postal_code) if cti_data.postal_code else "",
                'list_address': address if address else ""
            }
            
            self.orderer_data = {
                'operator': '',  # 対応者名は空で初期化
                'available_time': '',  # 出やすい時間帯は空で初期化
                'contractor': customer_name if customer_name else "",  # 変換済みの顧客名を使用
                'furigana': customer_furigana,  # 自動生成したフリガナを設定
                'birth_date': '1926/1/1',  # 誕生日の初期値を設定
                'order_person': '',  # 受注者名は空で初期化
                'employee_number': '',  # 社番は空で初期化
                'fee': '2500円～3000円',  # デフォルト値を設定
                'net_usage': 'なし',  # デフォルト値を設定
                'family_approval': 'なし',  # デフォルト値を設定
                'other_number': 'なし',  # デフォルト値を設定
                'phone_device': 'プッシュホン',  # デフォルト値を設定
                'forbidden_line': 'なし',  # デフォルト値を設定
                'nd': '',  # NDは空で初期化
                'relationship': ''  # 関係性は空で初期化
            }
            
            self.order_data = {
                'current_line': 'アナログ',  # デフォルト値を設定
                'order_date': f"{datetime.datetime.now().month}/{datetime.datetime.now().day}",
                'judgment': 'OK'  # デフォルト値を設定
            }
            
            # プレビューテキストを生成
            preview_text = self.generate_preview_text()
            if preview_text:
                self.preview_text.setText(preview_text)
            
            # 受注者入力項目ダイアログを表示
            dialog = OrdererInputDialog(self, self.orderer_data)
            
            # 提供判定処理を開始（非同期で実行）
            from PySide6.QtCore import QTimer
            QTimer.singleShot(0, self.start_service_area_search)
            
            # ダイアログの結果を処理
            result = dialog.exec()
            
            # 作成中止が選択された場合
            if result == DIALOG_CANCEL:
                logging.info("作成中止が選択されました")
                self.preview_text.clear()
                self.statusBar().showMessage("作成中止")
                return
            
            # 受注者情報を保存
            self.orderer_data = dialog.get_saved_data()
            
            # プレビューテキストが既に設定されている場合は何もしない
            # （作成ボタンクリック時にすでにプレビューテキストが設定されている）
            
        except Exception as e:
            logging.error(f"誘導モードの開始中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"誘導モードの開始中にエラーが発生しました: {e}")
    
    def start_service_area_search(self):
        """提供判定処理を開始"""
        try:
            postal_code = self.address_data.get('postal_code', '')
            address = self.address_data.get('address', '')
            
            if not postal_code or not address:
                logging.warning("郵便番号または住所が空のため、提供判定を行いません")
                self.update_judgment_result("未検索")
                return
            
            # 提供判定中の表示に更新
            self.update_judgment_result("検索中...")
            
            # 非同期で検索を実行
            from PySide6.QtCore import QThread, Signal
            
            class SearchThread(QThread):
                finished = Signal(dict)
                error = Signal(str)
                
                def __init__(self, postal_code, address):
                    super().__init__()
                    self.postal_code = postal_code
                    self.address = address
                
                def run(self):
                    try:
                        result = search_service_area(self.postal_code, self.address)
                        self.finished.emit(result)
                    except Exception as e:
                        self.error.emit(str(e))
            
            # 検索スレッドを作成して開始
            self.search_thread = SearchThread(postal_code, address)
            self.search_thread.finished.connect(self.handle_search_result)
            self.search_thread.error.connect(self.handle_search_error)
            self.search_thread.start()
            
            logging.info(f"提供エリア検索を開始しました: postal_code={postal_code}, address={address}")
            
        except Exception as e:
            logging.error(f"提供判定処理の開始中にエラー: {e}", exc_info=True)
            self.update_judgment_result("検索エラー")
    
    def handle_search_result(self, result):
        """検索結果を処理"""
        try:
            if result.get("status") == "available":
                self.update_judgment_result("提供可能")
            elif result.get("status") == "unavailable":
                self.update_judgment_result("提供エリア外")
            else:
                self.update_judgment_result("判定失敗")
            
            logging.info(f"提供エリア検索が完了しました: {result}")
            
        except Exception as e:
            logging.error(f"検索結果の処理中にエラー: {e}", exc_info=True)
            self.update_judgment_result("検索エラー")
    
    def handle_search_error(self, error_message):
        """検索エラーを処理"""
        try:
            logging.error(f"提供エリア検索中にエラー: {error_message}")
            self.update_judgment_result("検索エラー")
            
        except Exception as e:
            logging.error(f"エラー処理中に別のエラー: {e}", exc_info=True)

    def show_address_dialog(self):
        """住所情報ダイアログを表示"""
        try:
            # 以前のダイアログにはshow_address_dialogは保持しますが、別途管理するので
            # active_search_threadsでスレッドを管理するため、スレッド停止処理は削除
            
            # 新しいダイアログを作成
            dialog = AddressInfoDialog(self, self.address_data)
            self.address_dialog = dialog  # ダイアログへの参照を保持
            result = dialog.exec()
            
            # 現在のダイアログのデータを保存
            self.address_data = dialog.get_saved_data()
            
            # スレッドはダイアログを超えて動き続けるよう、ここではstopしない
            # スレッドの管理はactive_search_threadsで行う
            
            if result == QDialog.DialogCode.Accepted:
                self.show_list_dialog()
                
        except Exception as e:
            logging.error(f"住所情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"住所情報ダイアログの表示中にエラーが発生しました: {e}")

    @Slot(str)
    def update_judgment_result(self, result):
        """提供判定結果をメイン画面に反映する"""
        try:
            # 同じメソッドが複数回呼び出されるのを防ぐために結果をログに記録
            logging.info(f"★★★ メイン画面のupdate_judgment_result呼び出し: {result} ★★★")
            
            # judgment_result_labelが存在することを確認
            if not hasattr(self, 'judgment_result_label'):
                # 画面レイアウトに合わせて自動的に作成（なければ）
                logging.info("judgment_result_labelが見つからないため作成します")
                self.init_judgment_result_label()
            
            # 判定結果に応じてスタイルを変更
            if result == "検索中...":
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #FFA500;
                        border-radius: 4px;
                        background-color: #FFF3E0;
                        color: #E65100;
                    }
                """
            elif result == "検索エラー":
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #f44336;
                        border-radius: 4px;
                        background-color: #FFEBEE;
                        color: #B71C1C;
                    }
                """
            elif result == "提供エリア外":
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #FF9800;
                        border-radius: 4px;
                        background-color: #FFF3E0;
                        color: #E65100;
                    }
                """
            else:  # "提供可能"など
                style = """
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #4CAF50;
                        border-radius: 4px;
                        background-color: #E8F5E9;
                        color: #2E7D32;
                    }
                """
            
            # メイン画面の提供判定結果ラベルを更新
            self.judgment_result_label.setText(f"提供エリア: {result}")
            self.judgment_result_label.setStyleSheet(style)
            self.judgment_result_label.setVisible(True)  # 必ず表示
            logging.info(f"★★★ 提供判定結果ラベルを更新しました: {result} ★★★")
            
            # judgment_comboの値も更新
            try:
                if hasattr(self, 'judgment_combo'):
                    if result == "提供可能":
                        self.judgment_combo.setCurrentText("OK")
                        logging.info("judgment_comboを'OK'に設定しました")
                    elif result == "提供エリア外":
                        self.judgment_combo.setCurrentText("NG")
                        logging.info("judgment_comboを'NG'に設定しました")
            except Exception as combo_error:
                logging.error(f"judgment_comboの更新でエラー: {combo_error}")
            
            # プレビューも更新
            try:
                if hasattr(self, 'generate_preview_text'):
                    self.generate_preview_text()
                    logging.info("プレビューを更新しました")
            except Exception as preview_error:
                logging.error(f"プレビュー更新でエラー: {preview_error}")
            
            # UIが確実に更新されるようにイベントを処理
            QApplication.processEvents()
            
            # 結果をログに記録
            logging.info(f"★★★ 提供判定結果の更新が完了しました: {result} ★★★")
            
        except Exception as e:
            logging.error(f"提供判定結果の更新中にエラー: {e}", exc_info=True)
            try:
                if hasattr(self, 'judgment_result_label'):
                    self.judgment_result_label.setText("提供エリア: 更新エラー")
                    self.judgment_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #f44336;
                            border-radius: 4px;
                            background-color: #FFEBEE;
                            color: #B71C1C;
                        }
                    """)
            except Exception as inner_e:
                logging.error(f"エラー処理中に別のエラー: {inner_e}")
    
    def init_judgment_result_label(self):
        """判定結果表示ラベルを初期化する"""
        try:
            logging.info("判定結果ラベルを初期化します")
            # プレビューエリアを取得
            preview_area = None
            
            # プレビューエリアを探す
            for child in self.findChildren(QWidget):
                if hasattr(child, 'objectName') and child.objectName() == "preview_area":
                    preview_area = child
                    break
            
            if not preview_area and hasattr(self, 'preview_area'):
                preview_area = self.preview_area
            
            if not preview_area:
                # プレビューエリアが見つからない場合は直接メインウィンドウに追加
                logging.info("プレビューエリアが見つからないため、メインウィンドウに直接追加します")
                self.judgment_result_label = QLabel("提供エリア: 未検索", self)
                self.judgment_result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                        background-color: #f8f9fa;
                    }
                """)
                self.judgment_result_label.move(50, 50)
                self.judgment_result_label.resize(200, 30)
                self.judgment_result_label.show()
            else:
                # プレビューエリアに追加
                layout = preview_area.layout()
                if not layout:
                    layout = QVBoxLayout(preview_area)
                    preview_area.setLayout(layout)
                
                self.judgment_result_label = QLabel("提供エリア: 未検索")
                self.judgment_result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                        background-color: #f8f9fa;
                    }
                """)
                # レイアウトの先頭に追加
                layout.insertWidget(0, self.judgment_result_label)
            
            logging.info("判定結果ラベルの初期化が完了しました")
        except Exception as e:
            logging.error(f"判定結果ラベルの初期化中にエラー: {e}", exc_info=True)

    def show_list_dialog(self):
        """リスト情報ダイアログを表示"""
        try:
            dialog = ListInfoDialog(self, self.list_data)
            result = dialog.exec()
            
            # 現在のダイアログのデータを保存
            self.list_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Accepted:
                self.show_orderer_dialog()
            else:
                # 戻るボタンが押された場合、前のダイアログを表示
                self.show_address_dialog()
                
        except Exception as e:
            logging.error(f"リスト情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"リスト情報ダイアログの表示中にエラーが発生しました: {e}")

    def show_orderer_dialog(self):
        """受注者情報ダイアログを表示"""
        try:
            dialog = OrdererInputDialog(self, self.orderer_data)
            result = dialog.exec()
            
            # 現在のダイアログのデータを保存
            self.orderer_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Accepted:
                self.show_order_dialog()
            else:
                # 戻るボタンが押された場合、前のダイアログを表示
                self.show_list_dialog()
                
        except Exception as e:
            logging.error(f"受注者情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"受注者情報ダイアログの表示中にエラーが発生しました: {e}")

    def show_order_dialog(self):
        """受注情報ダイアログを表示"""
        try:
            dialog = OrderInfoDialog(self, self.order_data)
            result = dialog.exec()
            
            # 現在のダイアログのデータを保存
            self.order_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Rejected:
                # 戻るボタンが押された場合、前のダイアログを表示
                self.show_orderer_dialog()
                
        except Exception as e:
            logging.error(f"受注情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"受注情報ダイアログの表示中にエラーが発生しました: {e}")
    
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
        """入力フォームを作成します"""
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
        
        # 生年月日入力グループ
        birth_date_group = QGroupBox("生年月日")
        birth_date_layout = QHBoxLayout()
        
        # 元号選択
        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["令和", "平成", "昭和", "大正", "明治"])
        self.era_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.era_combo)
        
        # 年選択
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1, 151)])
        self.year_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.year_combo)
        birth_date_layout.addWidget(QLabel("年"))
        
        # 月選択
        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.month_combo)
        birth_date_layout.addWidget(QLabel("月"))
        
        # 日選択
        self.day_combo = NoWheelComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.day_combo)
        birth_date_layout.addWidget(QLabel("日"))
        
        birth_date_group.setLayout(birth_date_layout)
        input_layout.addWidget(birth_date_group)
        
        # 受注者名
        input_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        input_layout.addWidget(self.order_person_input)
        
        # 料金認識を追加（移動）
        input_layout.addWidget(QLabel("料金認識"))
        fee_layout = QHBoxLayout()
        self.fee_combo = NoWheelComboBox()
        self.fee_combo.addItems(["2500円～3000円", "3500円～4000円"])
        self.fee_combo.currentTextChanged.connect(self.on_fee_combo_changed)
        fee_layout.addWidget(self.fee_combo)
        self.fee_input = QLineEdit()
        self.fee_input.setPlaceholderText("手動入力")
        self.fee_input.textChanged.connect(self.reset_background_color)
        fee_layout.addWidget(self.fee_input)
        input_layout.addLayout(fee_layout)
        
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
        self.progress_bar.setRange(0, 100)  # 0-100%の範囲に設定
        self.progress_bar.setValue(0)  # 初期値を0%に設定
        self.progress_bar.setFixedHeight(10)  # 高さを10ピクセルに設定
        self.progress_bar.setTextVisible(True)  # テキストを表示
        self.progress_bar.setFormat("%p%")  # パーセント表示
        
        # アニメーションの設定
        self.progress_animation = QPropertyAnimation(self.progress_bar, b"value")
        self.progress_animation.setDuration(200)  # 200ミリ秒でアニメーション
        self.progress_animation.setEasingCurve(QEasingCurve.InOutQuad)  # イージング効果を追加
        
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #E3F2FD;
                border-radius: 5px;
                text-align: center;
                font-size: 10px;
                padding: 2px;
            }
            QProgressBar::chunk {
                background-color: #3498DB;
                border-radius: 5px;
                width: 10px; /* チャンクの最小幅を設定 */
                margin: 0px;
            }
            QProgressBar::chunk:hover {
                background-color: #2980B9;
            }
        """)

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
        self.list_address_input = QLineEdit()
        list_layout.addWidget(self.list_address_input)
        
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
        try:
            # 誘導モードの場合のみ、提供判定結果を表示するエリアを追加
            if self.current_mode != 'simple':
                # 提供エリア検索結果表示用のラベル
                self.judgment_result_label = QLabel("提供エリア: 未検索")
                self.judgment_result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #ddd;
                        border-radius: 4px;
                        background-color: #f8f8f8;
                    }
                """)
                parent_layout.addWidget(self.judgment_result_label)
            
            # プレビューテキストエリア
            self.preview_text = QTextEdit()
            self.preview_text.setReadOnly(True)
            self.preview_text.setMinimumHeight(300)
            parent_layout.addWidget(self.preview_text)
            
            # 通常モードの場合のみ、プレビュー更新ボタンを追加
            if self.current_mode == 'normal':
                # プレビュー更新ボタン
                self.update_preview_btn = QPushButton("プレビュー更新")
                self.update_preview_btn.setStyleSheet("""
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
                # プレビュー更新ボタンのシグナル接続
                self.update_preview_btn.clicked.connect(self.generate_preview_text)
                parent_layout.addWidget(self.update_preview_btn)
            
        except Exception as e:
            logging.error(f"プレビューエリア作成中にエラー: {e}")
    
    def setup_signals(self):
        """シグナルの設定"""
        if self.current_mode == 'simple':
            # シンプルモード用のシグナル設定
            # 自動フォーマット用のシグナル
            self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
            self.postal_code_input.textChanged.connect(self.format_postal_code)
            self.postal_code_input.textChanged.connect(self.convert_to_half_width)
            self.list_postal_code_input.textChanged.connect(self.format_postal_code)
            self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
            self.address_input.textChanged.connect(self.convert_to_half_width)
            self.list_address_input.textChanged.connect(self.convert_to_half_width)
            self.era_combo.currentTextChanged.connect(self.update_year_combo)
            
            # 名前とフリガナのバリデーション用のシグナル
            self.contractor_input.textChanged.connect(self.validate_contractor_name)
            self.furigana_input.textChanged.connect(self.validate_furigana_input)
            self.list_name_input.textChanged.connect(self.validate_list_name)
            self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
            
            # フリガナ自動変換のシグナル
            self.contractor_input.textChanged.connect(self.auto_generate_furigana)
            self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
            self.address_input.textChanged.connect(self.auto_generate_address_furigana)
            
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
            
            # ボタンのシグナル接続
            self.area_search_btn.clicked.connect(self.search_service_area)
            self.map_btn.clicked.connect(self.open_street_view)
        else:
            # 誘導モード用のシグナル設定
            # プレビュー更新ボタンのシグナル接続
            if hasattr(self, 'update_preview_btn'):
                self.update_preview_btn.clicked.connect(self.update_preview)
    
    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # ダイアログがOKで閉じられた場合、設定を再読み込み
            self.load_settings()
            # フォントサイズを適用
            self.apply_font_size()
            # ウィジェットを更新
            self.update()
            # 全てのウィジェットを再描画
            for widget in self.findChildren(QWidget):
                if isinstance(widget, QListView):
                    widget.viewport().update()  # QListViewの場合はviewport()を更新
                else:
                    widget.update()
            logging.info("設定を更新しました")
            
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
                # 住所のハイフンとスペースの処理
                converted_address = data.address.replace('－', '-')  # 全角ハイフンを半角に
                converted_address = converted_address.replace('ー', '-')  # 長音記号を半角ハイフンに
                converted_address = converted_address.replace('−', '-')  # 別種の全角ハイフンを半角に
                converted_address = converted_address.replace(' ', '　')  # 半角スペースを全角に
                converted_address = convert_to_half_width_except_space(converted_address)
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
        try:
            # すべてのアクティブな検索スレッドを停止
            if hasattr(self, 'active_search_threads'):
                for thread in self.active_search_threads:
                    if thread and thread.isRunning():
                        logging.info("アクティブな検索スレッドを停止します")
                        thread.stop()
                self.active_search_threads.clear()
            
            # 電話ボタン監視を停止
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.stop_monitoring()
                
            event.accept()
        except Exception as e:
            logging.error(f"アプリケーション終了処理中にエラー: {e}")
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
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        # 携帯電話番号入力エリアの参照を削除
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
        # 受注者名はクリアしない（保持する）
        # self.order_person_input.clear()
        # 料金認識はクリアしない（保持する）
        # self.fee_input.clear()
        
        # 他番号、電話機、禁止回線には初期値を設定
        self.other_number_input.setText("なし")
        self.phone_device_input.setText("プッシュホン")
        self.forbidden_line_input.setText("なし")
        
        # NDと備考（名義人との関係性）をクリア
        self.nd_input.clear()
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
        menubar.clear()
        
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
        """
        アップデート設定ダイアログを表示する
        """
        dialog = UpdateDialog(self)
        dialog.settings_file = self.settings_file  # 設定ファイルのパスを渡す
        dialog.exec()
        
    def show_about_dialog(self):
        """
        バージョン情報ダイアログを表示する
        """
        msg = f"{APP_NAME} v{VERSION}\n\n"
        msg += "ライセンス: MIT License"
        QMessageBox.information(self, "バージョン情報", msg)

    def check_for_updates(self):
        """
        アップデートをチェック
        """
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

    def get_template(self):
        """フォーマットテンプレートを取得する"""
        try:
            if not hasattr(self, 'format_template') or not self.format_template:
                if hasattr(self, 'settings') and 'format_template' in self.settings:
                    self.format_template = self.settings['format_template']
                else:
                    logging.error("format_templateを設定から読み込めません")
                    QMessageBox.warning(self, "エラー", "テンプレートが設定されていません。\n設定画面でテンプレートを設定してください。")
                    return None
            return self.format_template
        except Exception as e:
            logging.error(f"テンプレート取得中にエラー: {e}")
            return None

    def reconstruct_ui(self):
        """
        モード切り替え時にUIを再構築する
        """
        try:
            logging.info("UIの再構築を開始します")
            
            # 現在のモードに基づいてUIを再構築
            if self.current_mode == 'simple':
                self.init_simple_mode()
            else:
                self.init_easy_mode()
            
            # フォントサイズを再適用
            font_size = self.settings.get('font_size', 10)
            self.set_font_size(font_size)
            
            logging.info("UIの再構築が完了しました")
        except Exception as e:
            logging.error(f"UIの再構築中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"UIの再構築中にエラーが発生しました: {str(e)}")

    def update_search_progress(self, message):
        """検索の進捗状況を更新する"""
        try:
            # メッセージからパーセンテージを抽出
            import re
            match = re.search(r'\((\d+)%\)', message)
            if match:
                new_value = int(match.group(1))
                current_value = self.progress_bar.value()
                
                # アニメーションの設定
                self.progress_animation.setStartValue(current_value)
                self.progress_animation.setEndValue(new_value)
                self.progress_animation.start()
                
            # プログレスバーとメッセージを表示
            self.progress_bar.setVisible(True)
            self.area_result_label.setText(message)
            self.area_result_label.setStyleSheet("color: #666666;")
            
        except Exception as e:
            logging.error(f"進捗更新中にエラー: {str(e)}")
            self.area_result_label.setText(message)

    def on_fee_combo_changed(self, text):
        """
        料金認識のコンボボックスが変更された時の処理
        
        Args:
            text (str): 選択されたテキスト
        """
        self.fee_input.setText(text)
        self.reset_background_color()

    def check_birth_date_age(self):
        """
        生年月日から年齢を計算し、80歳以上の場合に赤く表示する
        """
        try:
            # 現在の日付を取得
            now = datetime.datetime.now()
            current_year = now.year
            current_month = now.month
            current_day = now.day
            
            # 生年月日の情報を取得
            era = self.era_combo.currentText()
            year = int(self.year_combo.currentText())
            month = int(self.month_combo.currentText())
            day = int(self.day_combo.currentText())
            
            # 和暦を西暦に変換
            if era == "昭和":
                year = year + 1925
            elif era == "平成":
                year = year + 1988
            elif era == "令和":
                year = year + 2018
            elif era == "大正":
                year = year + 1911
            elif era == "明治":
                year = year + 1867
            # 西暦の場合はそのまま
            
            # 年齢を計算
            age = current_year - year
            
            # 誕生日がまだ来ていない場合は年齢を1つ減らす
            if (month > current_month) or (month == current_month and day > current_day):
                age -= 1
            
            # 80歳以上かどうかをチェック
            is_over_80 = age >= 80
            
            # 背景色を設定
            if is_over_80:
                style = "background-color: #FFEBEE;"  # 赤系の背景色
            else:
                style = ""  # デフォルトの背景色
            
            # 各コンボボックスにスタイルを適用
            self.era_combo.setStyleSheet(style)
            self.year_combo.setStyleSheet(style)
            self.month_combo.setStyleSheet(style)
            self.day_combo.setStyleSheet(style)
            
            # 80歳以上の場合にログを出力
            if is_over_80:
                logging.info(f"80歳以上の顧客が検出されました: {age}歳")
            
        except Exception as e:
            logging.error(f"年齢チェック中にエラー: {e}")


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
        self._progress_steps = [
            {"message": "住所情報を解析中...", "weight": 5},
            {"message": "NTT西日本のサイトにアクセス中...", "weight": 10},
            {"message": "郵便番号を入力中...", "weight": 15},
            {"message": "住所を選択中...", "weight": 20},
            {"message": "番地を入力中...", "weight": 20},
            {"message": "号を入力中...", "weight": 20},
            {"message": "提供可否を判定中...", "weight": 10}
        ]
        self._current_step = 0
        self._total_weight = sum(step["weight"] for step in self._progress_steps)
        self._accumulated_progress = 0
    
    def cancel(self):
        """検索をキャンセルする"""
        self._is_cancelled = True
    
    def _update_progress(self, message=None):
        """
        進捗状況を更新する
        
        Args:
            message (str, optional): カスタムメッセージ。指定がない場合は定義済みメッセージを使用
        """
        try:
            if message is None and self._current_step < len(self._progress_steps):
                step_info = self._progress_steps[self._current_step]
                message = step_info["message"]
                # 現在のステップの重みに基づいて進捗を計算
                self._accumulated_progress += step_info["weight"]
            elif message:
                # メッセージに含まれるパーセンテージを抽出
                import re
                percent_match = re.search(r'(\d+)%', message)
                if percent_match:
                    self._accumulated_progress = int(percent_match.group(1))
                else:
                    # メッセージにパーセンテージが含まれていない場合は、次のステップに進む
                    if self._current_step < len(self._progress_steps):
                        self._accumulated_progress += self._progress_steps[self._current_step]["weight"]
            
            # 進捗率を計算（最大95%まで）
            progress_percent = min(int((self._accumulated_progress / self._total_weight) * 95), 95)
            
            # 進捗メッセージを生成
            if "%" not in message:
                message = f"{message} ({progress_percent}%)"
            
            self._current_step += 1
            self.progress.emit(message)
            
        except Exception as e:
            logging.error(f"進捗更新中にエラー: {e}")
            self.progress.emit(f"{message} (進捗更新エラー)")
    
    def run(self):
        """提供エリア検索を実行し、結果をシグナルで通知する"""
        try:
            # 進捗状況を通知するコールバック関数を定義
            def progress_callback(message):
                if self._is_cancelled:
                    raise CancellationError("検索がキャンセルされました")
                self._update_progress(message)

            # 検索を実行
            self._update_progress()  # 初期進捗を表示
            result = search_service_area(
                self.postal_code,
                self.address,
                progress_callback=progress_callback
            )
            
            if self._is_cancelled:
                raise CancellationError("検索がキャンセルされました")
            
            # 検索完了時に100%を表示
            if result.get("status") == "available":
                self.progress.emit("提供可能です (100%)")
            elif result.get("status") == "unavailable":
                self.progress.emit("提供不可です (100%)")
            else:
                self.progress.emit("検索が完了しました (100%)")
            self.finished.emit(result)
            
        except CancellationError as e:
            logging.info("検索がキャンセルされました")
            self.progress.emit("検索がキャンセルされました (0%)")
            self.finished.emit({
                "status": "cancelled",
                "message": "検索がキャンセルされました"
            })
        except Exception as e:
            logging.error(f"検索処理中にエラーが発生: {str(e)}")
            self.progress.emit("エラーが発生しました (0%)")
            self.finished.emit({
                "status": "error",
                "message": f"検索処理中にエラーが発生: {str(e)}"
            })

class CancellationError(Exception):
    """検索キャンセル時に発生する例外"""
    pass

    def save_input_data(self, input_data):
        """
        入力データを保存する
        
        Args:
            input_data (dict): 保存する入力データ
        """
        try:
            # 保存先ディレクトリの作成
            save_dir = "input_data"
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            
            # ファイル名の生成（タイムスタンプ付き）
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"input_data_{timestamp}.json"
            filepath = os.path.join(save_dir, filename)
            
            # データの保存
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(input_data, f, ensure_ascii=False, indent=4)
            
            logging.info(f"入力データを保存しました: {filepath}")
            QMessageBox.information(self, "完了", "入力データの保存が完了しました。")
            
        except Exception as e:
            logging.error(f"入力データの保存中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"入力データの保存中にエラーが発生しました: {e}")

    @Slot()
    def generate_preview_text(self):
        """プレビューテキストを生成する"""
        try:
            logging.info("プレビューテキストの生成を開始")
            
            # フォーマットテンプレートの取得
            format_template = self.settings.get('format_template', '')
            logging.info(f"フォーマットテンプレート: {format_template}")
            
            # テンプレートが空の場合はエラー
            if not format_template:
                logging.error("フォーマットテンプレートが設定されていません")
                QMessageBox.warning(self, "警告", "フォーマットテンプレートが設定されていません。\n設定画面でテンプレートを設定してください。")
                return None
            
            # データの初期化
            data = {}
            
            # 各入力フィールドからデータを取得し、末尾のスペースを削除
            if hasattr(self, 'operator_input'):
                data['operator'] = self.operator_input.text().rstrip()
            if hasattr(self, 'available_time_input'):
                data['available_time'] = self.available_time_input.text().rstrip()
            if hasattr(self, 'contractor_input'):
                data['contractor'] = self.contractor_input.text().rstrip()
            if hasattr(self, 'furigana_input'):
                data['furigana'] = self.furigana_input.text().rstrip()
            if hasattr(self, 'postal_code_input'):
                data['postal_code'] = self.postal_code_input.text().rstrip()
            if hasattr(self, 'address_input'):
                data['address'] = self.address_input.text().rstrip()
            if hasattr(self, 'list_name_input'):
                data['list_name'] = self.list_name_input.text().rstrip()
            if hasattr(self, 'list_furigana_input'):
                data['list_furigana'] = self.list_furigana_input.text().rstrip()
            if hasattr(self, 'list_phone_input'):
                data['list_phone'] = self.list_phone_input.text().rstrip()
            if hasattr(self, 'list_postal_code_input'):
                data['list_postal_code'] = self.list_postal_code_input.text().rstrip()
            if hasattr(self, 'list_address_input'):
                data['list_address'] = self.list_address_input.text().rstrip()
            if hasattr(self, 'order_person_input'):
                data['order_person'] = self.order_person_input.text().rstrip()
            if hasattr(self, 'fee_input'):
                data['fee'] = self.fee_input.text().rstrip()
            if hasattr(self, 'nd_input'):
                data['nd'] = self.nd_input.text().rstrip()
            if hasattr(self, 'relationship_input'):
                data['relationship'] = self.relationship_input.text().rstrip()
            if hasattr(self, 'phone_device_input'):
                data['phone_device'] = self.phone_device_input.text().rstrip()
            if hasattr(self, 'forbidden_line_input'):
                data['forbidden_line'] = self.forbidden_line_input.text().rstrip()
            
            # コンボボックスからデータを取得
            if hasattr(self, 'current_line_combo'):
                data['current_line'] = self.current_line_combo.currentText().rstrip()
            if hasattr(self, 'order_date_input'):
                data['order_date'] = self.order_date_input.text().rstrip()
            if hasattr(self, 'judgment_combo'):
                data['judgment'] = self.judgment_combo.currentText().rstrip()
            
            # データが空の場合はエラー
            if not data:
                logging.error("プレビュー生成に必要なデータが取得できません")
                return None
            
            # テンプレートの置換
            preview_text = format_template
            for key, value in data.items():
                placeholder = f"{{{key}}}"
                preview_text = preview_text.replace(placeholder, str(value or ''))
                logging.debug(f"プレースホルダー {placeholder} を {value} に置換")
            
            logging.info("プレビューテキストの生成が完了")
            
            # プレビューテキストを設定
            if hasattr(self, 'preview_text'):
                self.preview_text.setText(preview_text)
            
            return preview_text
            
        except Exception as e:
            logging.error(f"プレビュー生成中にエラー: {e}", exc_info=True)
            return None

    def load_settings(self):
        """設定を読み込む"""
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                    logging.info("設定ファイルを読み込みました")
                    logging.info(f"設定内容: {self.settings}")
            else:
                self.settings = {}
                logging.warning("設定ファイルが存在しません")
            
            # フォーマットテンプレートの確認
            if 'format_template' not in self.settings or not self.settings['format_template']:
                logging.error("フォーマットテンプレートが設定されていません")
                QMessageBox.warning(self, "警告", "フォーマットテンプレートが設定されていません。\n設定画面でテンプレートを設定してください。")
                return
            
            logging.info(f"フォーマットテンプレート: {self.settings['format_template']}")
            
            # フォントサイズの設定
            font_size = self.settings.get('font_size', 10)
            logging.info(f"フォントサイズを {font_size} に設定しました")
            
            # 電話ボタン監視の設定
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.update_settings()
            
        except Exception as e:
            logging.error(f"設定の読み込み中にエラーが発生しました: {e}", exc_info=True)
            self.settings = {}
            QMessageBox.critical(self, "エラー", f"設定の読み込み中にエラーが発生しました: {e}")

