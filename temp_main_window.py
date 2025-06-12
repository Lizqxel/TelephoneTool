"""
メインウィンドウモジュール

こ�Eモジュールは、アプリケーションのメインウィンドウを提供します、E"""

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
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread, QPropertyAnimation, QEasingCurve, QRect, QPoint
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices, QPalette, QColor

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
from services.cti_status_monitor import CTIStatusMonitor
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
    """スクロールでの値変更を防止するカスタムコンボ�EチE��ス"""
    def wheelEvent(self, event):
        """ホイールイベントを無要E""
        event.ignore()

class NoWheelComboBox(QComboBox):
    """スクロールイベントを無視するQComboBox"""
    def wheelEvent(self, event):
        event.ignore()

class MainWindow(QMainWindow, MainWindowFunctions):
    """メインウィンドウクラス"""
    
    # カスタムシグナル�E�CTI自動�E琁E��
    trigger_auto_search = Signal()
    
    def set_font_size(self, size):
        """
        フォントサイズを設定すめE        
        Args:
            size (int): 設定するフォントサイズ
        """
        try:
            # フォントサイズを設宁E            font = QFont()
            font.setPointSize(size)
            
            # 吁E��ィジェチE��にフォントを適用
            self.setFont(font)
            
            # プレビューエリアのフォントサイズを設宁E            if hasattr(self, 'preview_text'):
                self.preview_text.setFont(font)
            
            logging.info(f"フォントサイズめE{size} に設定しました")
            
        except Exception as e:
            logging.error(f"フォントサイズの設定中にエラーが発生しました: {e}")
    
    def setup_logging(self):
        """
        ログ設定を行う
        """
        try:
            # ログチE��レクトリの作�E
            log_dir = "logs"
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            # ログファイル名�E生�E�E�タイムスタンプ付き�E�E            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(log_dir, f"app_{timestamp}.log")
            
            # ログ設宁E            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler()
                ]
            )
            
            logging.info("ログ設定を完亁E��ました")
            
        except Exception as e:
            print(f"ログ設定中にエラーが発生しました: {e}")
    
    def __init__(self):
        """
        メインウィンドウの初期匁E        """
        super().__init__()
        
        # バ�Eジョン惁E��の設宁E        self.version = "1.0.0"
        
        # モード変更フラグ�E�設定ダイアログ用�E�E        self.mode_changed = False
        self.new_mode = None
        
        # ログ設宁E        self.setup_logging()
        
        # 設定ファイルのパスを設宁E        if getattr(sys, 'frozen', False):
            # exeファイルとして実行されてぁE��場吁E            self.settings_file = os.path.join(os.path.dirname(sys.executable), 'settings.json')
        else:
            # 通常のPythonスクリプトとして実行されてぁE��場吁E            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
        
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # 設定を読み込む
        self.settings = {}
        
        # 設定ファイルが存在しなぁE��合�E新規作�E
        if not os.path.exists(self.settings_file):
            logging.info("設定ファイルが存在しなぁE��め、新規作�EしまぁE)
            self.save_mode_settings('simple', True)
        
        # 設定を読み込む
        self.load_settings()
        
        # アクチE��ブな検索スレチE��を保持するリスチE        self.active_search_threads = []
        
        # モード設宁E        self.current_mode = self.settings.get('mode', 'simple')
        logging.info(f"現在のモーチE {self.current_mode}")
        
        # ウィンドウの基本設宁E        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setGeometry(100, 100, 800, 600)
        
        # 生年月日入力用のコンボ�EチE��スを�E期化
        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["西暦", "平戁E, "昭咁E])
        
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1926, datetime.datetime.now().year + 1)])
        
        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        
        self.day_combo = NoWheelComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        
        # メインウィジェチE��の設宁E        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト�E作�E
        self.main_layout = QVBoxLayout(main_widget)
        
        # モード選択ダイアログの表示�E�設定に基づぁE��表示を制御�E�E        if self.settings.get('show_mode_selection', True):
            self.show_mode_selection()
        
        # 選択されたモードに基づぁE��UIを�E期化
        if self.current_mode == 'simple':
            self.init_simple_mode()
        else:
            self.init_easy_mode()
        
        # CTI状態監視�E初期化と開始（設定に基づぁE��制御�E�E        cti_monitoring_enabled = self.settings.get('enable_cti_monitoring', True)
        logging.info(f"CTI監視設宁E {cti_monitoring_enabled}")
        
        if cti_monitoring_enabled:
            if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                self.cti_status_monitor = CTIStatusMonitor(self.on_cti_dialing_to_talking)
                self.cti_status_monitor.start_monitoring()
                logging.info("CTI状態監視を開始しました")
                
                # CTI自動�E琁E��のシグナル・スロチE��接続（重褁E��続を防ぐ！E                if not self.trigger_auto_search.isSignalConnected(self.trigger_auto_search, self.auto_search_service_area):
                    self.trigger_auto_search.connect(self.auto_search_service_area)
        else:
            logging.info("CTI監視が設定で無効になってぁE��ぁE)
            self.cti_status_monitor = None
        
        # 自動�E琁E�E重褁E��行防止用フラグ
        if not hasattr(self, 'is_auto_processing'):
            self.is_auto_processing = False
            self.last_auto_processing_time = 0
        
        # フォントサイズの設宁E        font_size = self.settings.get('font_size', 10)
        self.set_font_size(font_size)
    
    def check_and_show_mode_selection(self):
        """
        モード選択ダイアログの表示を確認し、忁E��に応じて表示する
        """
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # モード設定が存在しなぁE��合、また�E次回以降表示する設定�E場吁E                    if 'mode' not in settings or settings.get('show_mode_selection', True):
                        self.show_mode_selection_dialog()
                    else:
                        self.current_mode = settings.get('mode', 'simple')
            else:
                # 設定ファイルが存在しなぁE��合�E、忁E��モード選択ダイアログを表示
                self.show_mode_selection_dialog()
                # 設定ファイルを作�E
                self.save_mode_settings('simple', True)
        except Exception as e:
            logging.error(f"モード設定�E読み込み中にエラーが発生しました: {e}")
            # エラーが発生した場合�E、モード選択ダイアログを表示
            self.show_mode_selection_dialog()
    
    def show_mode_selection(self):
        """
        モード選択ダイアログを表示する
        設定ファイルのshow_mode_selectionの値に基づぁE��表示を制御する
        """
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # show_mode_selectionがFalseの場合�E表示しなぁE                    if not settings.get('show_mode_selection', True):
                        self.current_mode = settings.get('mode', 'simple')
                        return
            # 設定ファイルが存在しなぁE��合、また�Eshow_mode_selectionがTrueの場合�E表示
            self.show_mode_selection_dialog()
            # 設定ファイルが存在しなぁE��合�E作�E
            if not os.path.exists(self.settings_file):
                self.save_mode_settings('simple', True)
        except Exception as e:
            logging.error(f"モード設定�E読み込み中にエラーが発生しました: {e}")
            # エラーが発生した場合�E、モード選択ダイアログを表示
            self.show_mode_selection_dialog()
    
    def show_mode_selection_dialog(self):
        """
        モード選択ダイアログを表示し、E��択結果を保存すめE        """
        dialog = ModeSelectionDialog(self)
        if dialog.exec():
            # 選択されたモードを保孁E            self.current_mode = dialog.get_selected_mode()
            self.save_mode_settings(self.current_mode, dialog.should_show_again())
            logging.info(f"モードを {self.current_mode} に設定しました")
        else:
            # キャンセルされた場合�E、デフォルトでシンプルモードを使用
            self.current_mode = 'simple'
            self.save_mode_settings(self.current_mode, True)
            logging.info("モード選択がキャンセルされました。シンプルモードを使用します、E)
    
    def save_mode_settings(self, mode, show_again):
        """
        モード設定を保存すめE        
        Args:
            mode: 選択されたモード！Esimple'また�E'easy'�E�E            show_again: 次回から表示するかどぁE��
        """
        try:
            # 初期設定ファイル生�EかどぁE��をチェチE��
            is_initial_setup = not os.path.exists(self.settings_file)
            
            # 設定ファイルの読み込み
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            
            # モード設定を更新
            settings['mode'] = mode
            settings['show_mode_selection'] = show_again
            
            # 初期設定ファイル生�E時�EチE��ォルト値を設宁E            if is_initial_setup:
                logging.info("初期設定ファイルを生成します！ETI監視設定を含む�E�E)
                # チE��ォルト�EフォーマットテンプレーチE                default_format = """対応老E��お客様�E名前�E�：{operator}
工事希望日
☁E�EめE��ぁE��間帯�E�{available_time} 
☁E��話取次�E�アナログ→�E電話
☁E��話OP�E�E☁E��緁E契紁E��E書類名義)�E�{contractor}
フリガナ：{furigana}
生年月日�E�{birth_date}
郵便番号�E�{postal_code}
住所�E�{address}
リスト名�E�{list_name}
リスト名フリガナ：{list_furigana}
電話番号�E�{list_phone}
リスト郵便番号�E�{list_postal_code}
リスト住所�E�{list_address}
現状回線：{current_line}
受注日�E�{order_date}
受注老E��{order_person}
提供判定：{judgment}

料��認識：{fee}
ネット利用�E�{net_usage}
家族亁E���E�{family_approval}

他番号�E�{other_number}
電話機：{phone_device}
禁止回線：{forbidden_line}
ND�E�{nd}

備老E��{relationship}
お客様が今使ってぁE��回線：アナログ
案�E料�߁E�E500冁E"""
                
                # 初期設定�EチE��ォルト値を設宁E                settings.update({
                    'format_template': default_format,
                    'font_size': 9,
                    'delay_seconds': 0,
                    'browser_settings': {
                        'headless': False,
                        'disable_images': True,
                        'show_popup': True,
                        'auto_close': True,
                        'page_load_timeout': 30,
                        'script_timeout': 30
                    },
                    # CTI監視設定�EチE��ォルト値�E�オンに設定！E                    'enable_cti_monitoring': True,
                    'enable_auto_cti_processing': True,
                    'cti_monitor_interval': 0.2,
                    'cti_auto_processing_cooldown': 3.0
                })
            
            # 設定ファイルに保孁E            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            logging.info(f"モード設定を保存しました: mode={mode}, show_mode_selection={show_again}")
            if is_initial_setup:
                logging.info("CTI監視設定を有効にして初期設定ファイルを生成しました")
                
        except Exception as e:
            logging.error(f"モード設定�E保存中にエラーが発生しました: {e}")
    
    def init_simple_mode(self):
        """通常モード�EUIを�E期化"""
        logging.info("通常モード�E初期化を開姁E)
        
        # 設定に基づぁE��ウィンドウタイトルを設宁E        self.setWindowTitle("コールセンター業務効玁E��チE�Eル - 通常モーチE)
        self.setMinimumSize(600, 400)
        
        # メインウィジェチE��の設宁E        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト�E設宁E        main_layout = QVBoxLayout(main_widget)
        
        # 設定ファイルのパスを確誁E        if not hasattr(self, 'settings_file'):
            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
            logging.info(f"設定ファイルのパスを設宁E {self.settings_file}")
        
        logging.info(f"設定ファイルの存在確誁E {os.path.exists(self.settings_file)}")
        
        # 設定を読み込む
        if not hasattr(self, 'settings'):
            self.settings = {}
        
        # format_templateを設宁E        if not hasattr(self, 'format_template') or not self.format_template:
            logging.info("format_templateを設定しまぁE)
            self.load_settings()
            if hasattr(self, 'settings') and 'format_template' in self.settings:
                self.format_template = self.settings['format_template']
                logging.info(f"format_templateを設定しました: {self.format_template[:100]}...")
            else:
                logging.error("format_templateの設定に失敗しました")
                QMessageBox.warning(self, "エラー", "チE��プレート�E設定に失敗しました、E)
                return
        
        # トップバーの作�E
        self.create_top_bar(main_layout)
        
        # スプリチE��ーの作�E
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
        splitter.setHandleWidth(2)  # スプリチE��ーハンドルの幁E��設宁E        
        # 入力フォームエリア�E�左側�E�をスクロール可能に
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        self.create_input_form(form_layout)
        
        # スクロールエリアの作�E
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
        
        # プレビューエリア�E�右側�E�E        preview_group = QGroupBox("プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        self.create_preview_area(preview_layout)
        
        # スプリチE��ーにウィジェチE��を追加
        splitter.addWidget(scroll_area)
        splitter.addWidget(preview_group)
        
        # 初期のサイズ比率を設定！E:3�E�E        splitter.setSizes([700, 300])
        
        # スプリチE��ーをメインレイアウトに追加
        main_layout.addWidget(splitter)
        
        # シグナルの設宁E        self.setup_signals()
        
        # Google Sheetsの設宁E        self.setup_google_sheets()
        
        # フォントサイズの適用
        self.apply_font_size()
        
        # CTI連携サービスの初期匁E        self.cti_service = OneClickService()
        
        # CTI状態監視�E初期化と開姁E        self.cti_status_monitor = CTIStatusMonitor(self.on_cti_dialing_to_talking)
        self.cti_status_monitor.start_monitoring()
        
        
        # CTI自動�E琁E��のシグナル・スロチE��接綁E        self.trigger_auto_search.connect(self.auto_search_service_area)
        
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
        
        # カウントダウン更新用のタイマ�E
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        
        self.init_menu()
        
        # 起動時にアチE�EチE�EトをチェチE��
        QTimer.singleShot(0, self.check_for_updates)
    
    def init_easy_mode(self):
        """誘導モード�EUIを�E期化"""
        # 設定に基づぁE��ウィンドウタイトルを設宁E        self.setWindowTitle("コールセンター業務効玁E��チE�Eル - 誘導モーチE)
        self.setMinimumSize(400, 300)
        
        # メインウィジェチE��の設宁E        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト�E設宁E        main_layout = QVBoxLayout(main_widget)
        
        # プレビューエリア
        preview_group = QGroupBox("プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        self.create_preview_area(preview_layout)
        main_layout.addWidget(preview_group)
        
        # ボタンエリア
        button_layout = QHBoxLayout()
        
        # 開始�Eタン
        self.start_button = QPushButton("開姁E)
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
        
        # 設定�Eタン
        self.settings_button = QPushButton("設宁E)
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
        
        # シグナルの設宁E        self.setup_signals()
        
        # Google Sheetsの設宁E        self.setup_google_sheets()
        
        # フォントサイズの適用
        self.apply_font_size()
        
        # CTI連携サービスの初期匁E        self.cti_service = OneClickService()
        
        # CTI状態監視�E初期化と開姁E        self.cti_status_monitor = CTIStatusMonitor(self.on_cti_dialing_to_talking)
        self.cti_status_monitor.start_monitoring()
        
        
        # CTI自動�E琁E��のシグナル・スロチE��接綁E        self.trigger_auto_search.connect(self.auto_search_service_area)

        self.init_menu()
    
    def start_easy_mode(self):
        """誘導モードを開姁E""
        try:
            logging.info("誘導モードを開姁E)
            
            # 提供判定結果をリセチE��
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
            
            # CTIチE�Eタを取征E            cti_data = self.cti_service.get_all_fields_data()
            if not cti_data:
                QMessageBox.warning(self, "警呁E, "CTIチE�Eタの取得に失敗しました、E)
                return
            
            # 顧客名�E処琁E��苗字と名前の間�Eスペ�Eスを�E角に�E�E            customer_name = cti_data.customer_name
            if customer_name:
                customer_name = customer_name.replace(' ', '　')  # 半角スペ�Eスを�E角に
                customer_name = convert_to_half_width_except_space(customer_name)
            
            # 住所の処琁E��ハイフンを半角に�E�E            address = cti_data.address
            if address:
                address = address.replace('�E�E, '-')  # 全角ハイフンを半角に
                address = address.replace('ー', '-')  # 長音記号を半角ハイフンに
                address = address.replace('∁E, '-')  # 別種の全角ハイフンを半角に
                address = address.replace(' ', '　')  # 半角スペ�Eスを�E角に
                address = convert_to_half_width_except_space(address)
            
            # チE�Eタの初期化と設宁E            self.address_data = {
                'postal_code': convert_to_half_width(cti_data.postal_code) if cti_data.postal_code else "",
                'address': address if address else ""
            }
            
            # 顧客名�Eフリガナを取得して設宁E            customer_furigana = ""
            if customer_name:
                # フリガナ変換APIを使用
                customer_furigana = convert_to_furigana(customer_name)
            
            self.list_data = {
                'list_name': customer_name if customer_name else "",
                'list_furigana': customer_furigana,  # 自動生成したフリガナを設宁E                'list_phone': convert_to_half_width(cti_data.phone) if cti_data.phone else "",
                'list_postal_code': convert_to_half_width(cti_data.postal_code) if cti_data.postal_code else "",
                'list_address': address if address else ""
            }
            
            self.orderer_data = {
                'operator': '',  # 対応老E��は空で初期匁E                'available_time': '',  # 出めE��ぁE��間帯は空で初期匁E                'contractor': customer_name if customer_name else "",  # 変換済みの顧客名を使用
                'furigana': customer_furigana,  # 自動生成したフリガナを設宁E                'birth_date': '1926/1/1',  # 誕生日の初期値を設宁E                'order_person': '',  # 受注老E��は空で初期匁E                'employee_number': '',  # 社番は空で初期匁E                'fee': '2500冁E��E000冁E,  # チE��ォルト値を設宁E                'net_usage': 'なぁE,  # チE��ォルト値を設宁E                'family_approval': 'なぁE,  # チE��ォルト値を設宁E                'other_number': 'なぁE,  # チE��ォルト値を設宁E                'phone_device': 'プッシュホン',  # チE��ォルト値を設宁E                'forbidden_line': 'なぁE,  # チE��ォルト値を設宁E                'nd': '',  # NDは空で初期匁E                'relationship': ''  # 関係性は空で初期匁E            }
            
            self.order_data = {
                'current_line': 'アナログ',  # チE��ォルト値を設宁E                'order_date': f"{datetime.datetime.now().month}/{datetime.datetime.now().day}",
                'judgment': 'OK'  # チE��ォルト値を設宁E            }
            
            # プレビューチE��ストを生�E
            preview_text = self.generate_preview_text()
            if preview_text:
                self.preview_text.setText(preview_text)
            
            # 受注老E�E力頁E��ダイアログを表示
            dialog = OrdererInputDialog(self, self.orderer_data)
            
            # 提供判定�E琁E��開始（非同期で実行！E            from PySide6.QtCore import QTimer
            QTimer.singleShot(0, self.start_service_area_search)
            
            # ダイアログの結果を�E琁E            result = dialog.exec()
            
            # 作�E中止が選択された場吁E            if result == DIALOG_CANCEL:
                logging.info("作�E中止が選択されました")
                self.preview_text.clear()
                self.statusBar().showMessage("作�E中止")
                return
            
            # 受注老E��報を保孁E            self.orderer_data = dialog.get_saved_data()
            
            # プレビューチE��ストが既に設定されてぁE��場合�E何もしなぁE            # �E�作�EボタンクリチE��時にすでにプレビューチE��ストが設定されてぁE���E�E            
        except Exception as e:
            logging.error(f"誘導モード�E開始中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"誘導モード�E開始中にエラーが発生しました: {e}")
    
    def start_service_area_search(self):
        """提供判定�E琁E��開姁E""
        try:
            postal_code = self.address_data.get('postal_code', '')
            address = self.address_data.get('address', '')
            
            if not postal_code or not address:
                logging.warning("郵便番号また�E住所が空のため、提供判定を行いません")
                self.update_judgment_result("未検索")
                return
            
            # 提供判定中の表示に更新
            self.update_judgment_result("検索中...")
            
            # 非同期で検索を実衁E            from PySide6.QtCore import QThread, Signal
            
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
            
            # 検索スレチE��を作�Eして開姁E            self.search_thread = SearchThread(postal_code, address)
            self.search_thread.finished.connect(self.handle_search_result)
            self.search_thread.error.connect(self.handle_search_error)
            self.search_thread.start()
            
            logging.info(f"提供エリア検索を開始しました: postal_code={postal_code}, address={address}")
            
        except Exception as e:
            logging.error(f"提供判定�E琁E�E開始中にエラー: {e}", exc_info=True)
            self.update_judgment_result("検索エラー")
    
    def handle_search_result(self, result):
        """検索結果を�E琁E""
        try:
            status = result.get("status")
            if status == "available":
                self.update_judgment_result("提供可能")
            elif status == "unavailable":
                self.update_judgment_result("提供エリア夁E)
            elif status == "apartment":
                # 雁E��住宁E�E場合�E明示皁E��表示
                self.update_judgment_result("雁E��住宁E��アパ�Eト�Eマンション等！E)
            else:
                self.update_judgment_result("判定失敁E)
            
            logging.info(f"提供エリア検索が完亁E��ました: {result}")
            
        except Exception as e:
            logging.error(f"検索結果の処琁E��にエラー: {e}", exc_info=True)
            self.update_judgment_result("検索エラー")
    
    def handle_search_error(self, error_message):
        """検索エラーを�E琁E""
        try:
            logging.error(f"提供エリア検索中にエラー: {error_message}")
            self.update_judgment_result("検索エラー")
            
        except Exception as e:
            logging.error(f"エラー処琁E��に別のエラー: {e}", exc_info=True)

    def show_address_dialog(self):
        """住所惁E��ダイアログを表示"""
        try:
            # 以前�Eダイアログにはshow_address_dialogは保持しますが、別途管琁E��る�Eで
            # active_search_threadsでスレチE��を管琁E��るため、スレチE��停止処琁E�E削除
            
            # 新しいダイアログを作�E
            dialog = AddressInfoDialog(self, self.address_data)
            self.address_dialog = dialog  # ダイアログへの参�Eを保持
            result = dialog.exec()
            
            # 現在のダイアログのチE�Eタを保孁E            self.address_data = dialog.get_saved_data()
            
            # スレチE��はダイアログを趁E��て動き続けるよぁE��ここではstopしなぁE            # スレチE��の管琁E�Eactive_search_threadsで行う
            
            if result == QDialog.DialogCode.Accepted:
                self.show_list_dialog()
                
        except Exception as e:
            logging.error(f"住所惁E��ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"住所惁E��ダイアログの表示中にエラーが発生しました: {e}")

    @Slot(str)
    def update_judgment_result(self, result):
        """提供判定結果をメイン画面に反映する"""
        try:
            # 同じメソチE��が褁E��回呼び出される�Eを防ぐために結果をログに記録
            logging.info(f"☁E�E☁Eメイン画面のupdate_judgment_result呼び出ぁE {result} ☁E�E☁E)
            
            # judgment_result_labelが存在することを確誁E            if not hasattr(self, 'judgment_result_label'):
                # 画面レイアウトに合わせて自動的に作�E�E�なければ�E�E                logging.info("judgment_result_labelが見つからなぁE��め作�EしまぁE)
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
            elif result == "提供エリア夁E:
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
            self.judgment_result_label.setVisible(True)  # 忁E��表示
            logging.info(f"☁E�E☁E提供判定結果ラベルを更新しました: {result} ☁E�E☁E)
            
            # judgment_comboの値も更新
            try:
                if hasattr(self, 'judgment_combo'):
                    if result == "提供可能":
                        self.judgment_combo.setCurrentText("OK")
                        logging.info("judgment_comboめEOK'に設定しました")
                    elif result == "提供エリア夁E:
                        self.judgment_combo.setCurrentText("NG")
                        logging.info("judgment_comboめENG'に設定しました")
            except Exception as combo_error:
                logging.error(f"judgment_comboの更新でエラー: {combo_error}")
            
            # プレビューも更新
            try:
                if hasattr(self, 'generate_preview_text'):
                    self.generate_preview_text()
                    logging.info("プレビューを更新しました")
            except Exception as preview_error:
                logging.error(f"プレビュー更新でエラー: {preview_error}")
            
            # UIが確実に更新されるよぁE��イベントを処琁E            QApplication.processEvents()
            
            # 結果をログに記録
            logging.info(f"☁E�E☁E提供判定結果の更新が完亁E��ました: {result} ☁E�E☁E)
            
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
                logging.error(f"エラー処琁E��に別のエラー: {inner_e}")
    
    def init_judgment_result_label(self):
        """判定結果表示ラベルを�E期化する"""
        try:
            logging.info("判定結果ラベルを�E期化しまぁE)
            # プレビューエリアを取征E            preview_area = None
            
            # プレビューエリアを探ぁE            for child in self.findChildren(QWidget):
                if hasattr(child, 'objectName') and child.objectName() == "preview_area":
                    preview_area = child
                    break
            
            if not preview_area and hasattr(self, 'preview_area'):
                preview_area = self.preview_area
            
            if not preview_area:
                # プレビューエリアが見つからなぁE��合�E直接メインウィンドウに追加
                logging.info("プレビューエリアが見つからなぁE��め、メインウィンドウに直接追加しまぁE)
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
                # レイアウト�E先頭に追加
                layout.insertWidget(0, self.judgment_result_label)
            
            logging.info("判定結果ラベルの初期化が完亁E��ました")
        except Exception as e:
            logging.error(f"判定結果ラベルの初期化中にエラー: {e}", exc_info=True)

    def show_list_dialog(self):
        """リスト情報ダイアログを表示"""
        try:
            dialog = ListInfoDialog(self, self.list_data)
            result = dialog.exec()
            
            # 現在のダイアログのチE�Eタを保孁E            self.list_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Accepted:
                self.show_orderer_dialog()
            else:
                # 戻る�Eタンが押された場合、前のダイアログを表示
                self.show_address_dialog()
                
        except Exception as e:
            logging.error(f"リスト情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"リスト情報ダイアログの表示中にエラーが発生しました: {e}")

    def show_orderer_dialog(self):
        """受注老E��報ダイアログを表示"""
        try:
            dialog = OrdererInputDialog(self, self.orderer_data)
            result = dialog.exec()
            
            # 現在のダイアログのチE�Eタを保孁E            self.orderer_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Accepted:
                self.show_order_dialog()
            else:
                # 戻る�Eタンが押された場合、前のダイアログを表示
                self.show_list_dialog()
                
        except Exception as e:
            logging.error(f"受注老E��報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"受注老E��報ダイアログの表示中にエラーが発生しました: {e}")

    def show_order_dialog(self):
        """受注惁E��ダイアログを表示"""
        try:
            dialog = OrderInfoDialog(self, self.order_data)
            result = dialog.exec()
            
            # 現在のダイアログのチE�Eタを保孁E            self.order_data = dialog.get_saved_data()
            
            if result == QDialog.DialogCode.Rejected:
                # 戻る�Eタンが押された場合、前のダイアログを表示
                self.show_orderer_dialog()
                
        except Exception as e:
            logging.error(f"受注惁E��ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"受注惁E��ダイアログの表示中にエラーが発生しました: {e}")
    
    def create_top_bar(self, parent_layout):
        """トップバーを作�E"""
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
        
        # ワンクリチE��取得�Eタン�E�名称変更�E�顧客惁E��取得！E        self.oneclick_btn = QPushButton("顧客惁E��取征E)
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
        
        # 既存�Eボタン
        self.clear_btn = QPushButton("入力クリア")
        self.cti_copy_btn = QPushButton("営コメ作�E")
        self.screenshot_btn = QPushButton("提供判定�EスクリーンショチE��確誁E)
        self.spreadsheet_btn = QPushButton("スプレチE��シート転記（未実裁E��E)
        self.settings_btn = QPushButton("設宁E)
        
        # ボタンのスタイル設宁E        button_style = """
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
        
        # 吁E�Eタンのサイズポリシーを設宁E        buttons = [self.clear_btn, self.cti_copy_btn, 
                  self.screenshot_btn, self.spreadsheet_btn, self.settings_btn]
        
        for btn in buttons:
            btn.setStyleSheet(button_style)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        # ボタンの接綁E        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.copy_cti_to_clipboard)
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)
        self.settings_btn.clicked.connect(self.show_settings)
        
        # ボタンをレイアウトに追加
        for btn in buttons:
            top_bar_layout.addWidget(btn)
        
        parent_layout.addWidget(top_bar)
    
    def create_input_form(self, parent_layout):
        """入力フォームを作�EしまぁE""
        # 受注老E�E力頁E��セクション�E�新しく追加�E�E        input_group = QGroupBox("受注老E�E力頁E��")
        input_layout = QVBoxLayout()
        
        # 対応老E��
        input_layout.addWidget(QLabel("対応老E��"))
        self.operator_input = QLineEdit()
        input_layout.addWidget(self.operator_input)
        
        # 出めE��ぁE��間帯�E�携帯番号入力！E        input_layout.addWidget(QLabel("出めE��ぁE��間帯�E�携帯番号�E�E))
        
        # 携帯番号パターン選抁E        self.mobile_pattern_combo = CustomComboBox()
        self.mobile_pattern_combo.addItems(["①携帯ありで番号がわかる", "②携帯なぁE, "③携帯ありで番号がわからなぁE])
        self.mobile_pattern_combo.currentTextChanged.connect(self.on_mobile_pattern_changed)
        input_layout.addWidget(self.mobile_pattern_combo)
        
        # 携帯番号入力欁E��Eつの枠�E�E        self.mobile_number_widget = QWidget()
        mobile_number_layout = QHBoxLayout(self.mobile_number_widget)
        mobile_number_layout.setContentsMargins(0, 0, 0, 0)
        
        self.mobile_part1_input = QLineEdit()
        self.mobile_part1_input.setMaxLength(3)
        self.mobile_part1_input.setPlaceholderText("090")
        self.mobile_part1_input.textChanged.connect(self.format_mobile_number_part)
        mobile_number_layout.addWidget(self.mobile_part1_input)
        
        mobile_number_layout.addWidget(QLabel("-"))
        
        self.mobile_part2_input = QLineEdit()
        self.mobile_part2_input.setMaxLength(4)
        self.mobile_part2_input.setPlaceholderText("1234")
        self.mobile_part2_input.textChanged.connect(self.format_mobile_number_part)
        mobile_number_layout.addWidget(self.mobile_part2_input)
        
        mobile_number_layout.addWidget(QLabel("-"))
        
        self.mobile_part3_input = QLineEdit()
        self.mobile_part3_input.setMaxLength(4)
        self.mobile_part3_input.setPlaceholderText("5678")
        self.mobile_part3_input.textChanged.connect(self.format_mobile_number_part)
        mobile_number_layout.addWidget(self.mobile_part3_input)
        
        input_layout.addWidget(self.mobile_number_widget)
        
        # 従来の出めE��ぁE��間帯入力欁E��互換性のため保持、E��表示�E�E        self.available_time_input = QLineEdit()
        self.available_time_input.hide()
        
        # 初期状態�E設宁E        self.mobile_pattern_combo.setCurrentText("②携帯なぁE)
        self.mobile_number_widget.hide()
        self.available_time_input.setText("携帯なぁE)
        
        # 契紁E��E��
        input_layout.addWidget(QLabel("契紁E��E��"))
        self.contractor_input = QLineEdit()
        input_layout.addWidget(self.contractor_input)
        
        # フリガチE        furigana_layout = QHBoxLayout()
        furigana_layout.addWidget(QLabel("フリガチE))
        self.furigana_mode_combo = CustomComboBox()
        self.furigana_mode_combo.addItems(["自勁E, "手動"])
        furigana_layout.addWidget(self.furigana_mode_combo)
        input_layout.addLayout(furigana_layout)
        self.furigana_input = QLineEdit()
        input_layout.addWidget(self.furigana_input)
        
        # 生年月日入力グルーチE        birth_date_group = QGroupBox("生年月日")
        birth_date_layout = QHBoxLayout()
        
        # 允E��選抁E        self.era_combo = NoWheelComboBox()
        self.era_combo.addItems(["西暦", "平戁E, "昭咁E])
        self.era_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.era_combo)
        
        # 年選抁E        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1926, datetime.datetime.now().year + 1)])
        self.year_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.year_combo)
        birth_date_layout.addWidget(QLabel("年"))
        
        # 月選抁E        self.month_combo = NoWheelComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.month_combo)
        birth_date_layout.addWidget(QLabel("朁E))
        
        # 日選抁E        self.day_combo = NoWheelComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.day_combo)
        birth_date_layout.addWidget(QLabel("日"))
        
        birth_date_group.setLayout(birth_date_layout)
        input_layout.addWidget(birth_date_group)
        
        # 受注老E��
        input_layout.addWidget(QLabel("受注老E��"))
        self.order_person_input = QLineEdit()
        input_layout.addWidget(self.order_person_input)
        
        # 料��認識を追加�E�移動！E        input_layout.addWidget(QLabel("料��認譁E))
        fee_layout = QHBoxLayout()
        self.fee_combo = NoWheelComboBox()
        self.fee_combo.addItems(["2500冁E��E000冁E, "3500冁E��E000冁E])
        self.fee_combo.currentTextChanged.connect(self.on_fee_combo_changed)
        fee_layout.addWidget(self.fee_combo)
        self.fee_input = QLineEdit()
        self.fee_input.setPlaceholderText("手動入劁E)
        self.fee_input.textChanged.connect(self.reset_background_color)
        fee_layout.addWidget(self.fee_input)
        input_layout.addLayout(fee_layout)
        
        # ネット利用
        input_layout.addWidget(QLabel("ネット利用"))
        self.net_usage_combo = CustomComboBox()
        self.net_usage_combo.addItems(["なぁE, "あり"])
        input_layout.addWidget(self.net_usage_combo)
        
        # 家族亁E��
        input_layout.addWidget(QLabel("家族亁E��"))
        self.family_approval_combo = CustomComboBox()
        self.family_approval_combo.addItems(["ok", "なぁE])
        input_layout.addWidget(self.family_approval_combo)
        
        # 他番号
        input_layout.addWidget(QLabel("他番号"))
        self.other_number_input = QLineEdit()
        self.other_number_input.setText("なぁE)
        input_layout.addWidget(self.other_number_input)
        
        # 電話橁E        input_layout.addWidget(QLabel("電話橁E))
        self.phone_device_input = QLineEdit()
        self.phone_device_input.setText("プッシュホン")
        input_layout.addWidget(self.phone_device_input)
        
        # 禁止回緁E        input_layout.addWidget(QLabel("禁止回緁E))
        self.forbidden_line_input = QLineEdit()
        self.forbidden_line_input.setText("なぁE)
        input_layout.addWidget(self.forbidden_line_input)
        
        # ND
        input_layout.addWidget(QLabel("ND"))
        self.nd_input = QLineEdit()
        input_layout.addWidget(self.nd_input)
        
        # リストとの関係性�E�表示を「名義人の○○」�E形式に変更�E�E        relationship_layout = QHBoxLayout()
        relationship_layout.addWidget(QLabel("備老E��E))
        self.relationship_input = QLineEdit()
        self.relationship_input.setPlaceholderText("名義人の...")
        relationship_layout.addWidget(self.relationship_input)
        input_layout.addLayout(relationship_layout)
        
        input_group.setLayout(input_layout)
        parent_layout.addWidget(input_group)
        
        # 住所惁E��セクション
        address_group = QGroupBox("住所惁E��")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        address_layout.addWidget(self.address_input)
        
        # 住所フリガチE        address_layout.addWidget(QLabel("住所フリガチE))
        self.address_furigana_input = QLineEdit()
        address_layout.addWidget(self.address_furigana_input)
        
        # マップアイコンボタン
        self.map_btn = QPushButton()
        self.map_btn.setFixedSize(24, 24)
        
        # アプリケーションの実行ディレクトリからの絶対パスを設宁E        app_dir = os.path.dirname(os.path.abspath(__file__))
        root_dir = os.path.dirname(app_dir)  # uiフォルダの親チE��レクトリ
        map_icon_path = os.path.join(root_dir, "map.png")
        
        # アイコンが存在する場合�Eみ設宁E        if os.path.exists(map_icon_path):
            self.map_btn.setIcon(QIcon(map_icon_path))
        else:
            # アイコンが見つからなぁE��合�E代替チE��ストを設宁E            self.map_btn.setText("🗺�E�E)
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

        # プログレスバ�E�E��E期状態では非表示�E�E        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)  # 0-100%の篁E��に設宁E        self.progress_bar.setValue(0)  # 初期値めE%に設宁E        self.progress_bar.setFixedHeight(10)  # 高さめE0ピクセルに設宁E        self.progress_bar.setTextVisible(True)  # チE��ストを表示
        self.progress_bar.setFormat("%p%")  # パ�Eセント表示
        
        # アニメーションの設宁E        self.progress_animation = QPropertyAnimation(self.progress_bar, b"value")
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
                width: 10px; /* チャンクの最小幁E��設宁E*/
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
        
        # リストフリガチE        list_furigana_layout = QHBoxLayout()
        list_furigana_layout.addWidget(QLabel("リストフリガチE))
        self.list_furigana_mode_combo = CustomComboBox()
        self.list_furigana_mode_combo.addItems(["自勁E, "手動"])
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
        
        # 受注惁E��セクション
        order_group = QGroupBox("受注惁E��")
        order_layout = QVBoxLayout()
        
        # 現状回緁E        order_layout.addWidget(QLabel("現状回緁E))
        self.current_line_combo = CustomComboBox()
        self.current_line_combo.addItems(["アナログ"])
        order_layout.addWidget(self.current_line_combo)
        
        # 受注日�E�本日自動�E力！E        order_layout.addWidget(QLabel("受注日"))
        self.order_date_input = QLineEdit()
        # 0埋めなし�E朁E日フォーマットを生�E
        now = datetime.datetime.now()
        month = str(now.month)  # 0埋めなし�E朁E        day = str(now.day)      # 0埋めなし�E日
        self.order_date_input.setText(f"{month}/{day}")
        self.order_date_input.setReadOnly(True)
        order_layout.addWidget(self.order_date_input)
        
        # 提供判宁E        order_layout.addWidget(QLabel("提供判宁E))
        self.judgment_combo = CustomComboBox()
        self.judgment_combo.addItems(["OK", "NG"])
        order_layout.addWidget(self.judgment_combo)
        
        order_group.setLayout(order_layout)
        parent_layout.addWidget(order_group)
    
    def create_preview_area(self, parent_layout):
        """プレビューエリアを作�E"""
        try:
            # 誘導モード�E場合�Eみ、提供判定結果を表示するエリアを追加
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
            
            # プレビューチE��ストエリア
            self.preview_text = QTextEdit()
            self.preview_text.setReadOnly(True)
            self.preview_text.setMinimumHeight(300)
            parent_layout.addWidget(self.preview_text)
            
            # 通常モード�E場合�Eみ、�Eレビュー更新ボタンを追加
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
                # プレビュー更新ボタンのシグナル接綁E                self.update_preview_btn.clicked.connect(self.generate_preview_text)
                parent_layout.addWidget(self.update_preview_btn)
            
        except Exception as e:
            logging.error(f"プレビューエリア作�E中にエラー: {e}")
    
    def setup_signals(self):
        """シグナルの設宁E""
        if self.current_mode == 'simple':
            # シンプルモード用のシグナル設宁E            # 自動フォーマット用のシグナル
            self.list_phone_input.textChanged.connect(self.format_phone_number_without_hyphen)
            self.postal_code_input.textChanged.connect(self.format_postal_code)
            self.postal_code_input.textChanged.connect(self.convert_to_half_width)
            self.list_postal_code_input.textChanged.connect(self.format_postal_code)
            self.list_postal_code_input.textChanged.connect(self.convert_to_half_width)
            self.address_input.textChanged.connect(self.convert_to_half_width)
            self.list_address_input.textChanged.connect(self.convert_to_half_width)
            self.era_combo.currentTextChanged.connect(self.update_year_combo)
            
            # 名前とフリガナ�EバリチE�Eション用のシグナル
            self.contractor_input.textChanged.connect(self.validate_contractor_name)
            self.furigana_input.textChanged.connect(self.validate_furigana_input)
            self.list_name_input.textChanged.connect(self.validate_list_name)
            self.list_furigana_input.textChanged.connect(self.validate_list_furigana)
            
            # フリガナ�E動変換のシグナル
            self.contractor_input.textChanged.connect(self.auto_generate_furigana)
            self.list_name_input.textChanged.connect(self.auto_generate_list_furigana)
            self.address_input.textChanged.connect(self.auto_generate_address_furigana)
            
            # 入力時に背景色をリセチE��するシグナル
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
            
            # ボタンのシグナル接綁E            self.area_search_btn.clicked.connect(self.search_service_area)
            self.map_btn.clicked.connect(self.open_street_view)
        else:
            # 誘導モード用のシグナル設宁E            # プレビュー更新ボタンのシグナル接綁E            if hasattr(self, 'update_preview_btn'):
                self.update_preview_btn.clicked.connect(self.update_preview)
    
    def show_settings(self):
        """設定ダイアログを表示"""
        # 変更前�ECTI監視設定を保孁E        old_cti_monitoring = self.settings.get('enable_cti_monitoring', True)
        
        dialog = SettingsDialog(self)
        if dialog.exec():
            # ダイアログがOKで閉じられた場合、設定を再読み込み
            self.load_settings()
            
            # CTI監視設定が変更された場合�E処琁E            new_cti_monitoring = self.settings.get('enable_cti_monitoring', True)
            if old_cti_monitoring != new_cti_monitoring:
                logging.info(f"CTI監視設定が変更されました: {old_cti_monitoring} ↁE{new_cti_monitoring}")
                
                if new_cti_monitoring:
                    # CTI監視を有効にする
                    if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                        self.cti_status_monitor = CTIStatusMonitor(self.on_cti_dialing_to_talking)
                        self.cti_status_monitor.start_monitoring()
                        logging.info("CTI状態監視を開始しました")
                        
                        # CTI自動�E琁E��のシグナル・スロチE��接綁E                        if not self.trigger_auto_search.isSignalConnected(self.trigger_auto_search, self.auto_search_service_area):
                            self.trigger_auto_search.connect(self.auto_search_service_area)
                    elif hasattr(self.cti_status_monitor, 'start_monitoring'):
                        self.cti_status_monitor.start_monitoring()
                        logging.info("CTI状態監視を再開しました")
                else:
                    # CTI監視を無効にする
                    if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                        self.cti_status_monitor.stop_monitoring()
                        logging.info("CTI状態監視を停止しました")
            
            # 既存�ECTI監視サービスの設定を更新
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                if hasattr(self.cti_status_monitor, 'update_settings'):
                    self.cti_status_monitor.update_settings()
                    logging.info("CTI監視サービスの設定を更新しました")
            
            # フォントサイズを適用
            self.apply_font_size()
            # ウィジェチE��を更新
            self.update()
            # 全てのウィジェチE��を�E描画
            for widget in self.findChildren(QWidget):
                if isinstance(widget, QListView):
                    widget.viewport().update()  # QListViewの場合�Eviewport()を更新
                else:
                    widget.update()
            logging.info("設定を更新しました")
    
    def update_countdown(self):
        """カウントダウン表示を更新"""
        try:
                if remaining_time > 0:
                    self.countdown_label.setText(f"惁E��取得まで: {int(remaining_time)}私E)
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
        CTIチE�Eタをフォームに反映しまぁE        
        Args:
            data: CTIから取得したデータ
        """
        try:
            # 顧客吁E            if data.customer_name:
                # 半角スペ�Eスを�E角スペ�Eスに変換
                converted_customer_name = data.customer_name.replace(' ', '　')
                converted_customer_name = convert_to_half_width_except_space(converted_customer_name)
                self.list_name_input.setText(converted_customer_name)
                self.contractor_input.setText(converted_customer_name)
            
            # 住所
            if data.address:
                # 住所のハイフンとスペ�Eスの処琁E                converted_address = data.address.replace('�E�E, '-')  # 全角ハイフンを半角に
                converted_address = converted_address.replace('ー', '-')  # 長音記号を半角ハイフンに
                converted_address = converted_address.replace('∁E, '-')  # 別種の全角ハイフンを半角に
                converted_address = converted_address.replace(' ', '　')  # 半角スペ�Eスを�E角に
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
                
            # プレビューを更新しなぁE��営業コメントを自動作�EしなぁE��E            # self.update_preview()
            
            # 成功メチE��ージ
            self.statusBar().showMessage("チE�Eタを取得しました", 5000)
            
        except Exception as e:
            logging.error(f"フォーム更新中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"フォームの更新中にエラーが発生しました: {e}")
            
    def fetch_cti_data(self):
        """CTIチE�Eタを取征E""
        try:
            # カウントダウン表示を非表示
            self.countdown_label.hide()
            self.countdown_timer.stop()
            
            # CTIチE�Eタの取得�E琁E            data = self.cti_service.get_all_fields_data()
            if data:
                # メインスレチE��でUIを更新
                QApplication.instance().postEvent(self, QEvent(QEvent.User))
                self.update_form_with_data(data)
                logging.info("CTIチE�Eタの取得に成功しました")
            else:
                logging.warning("CTIチE�Eタの取得に失敗しました")
        except Exception as e:
            logging.error(f"CTIチE�Eタの取得中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"CTIチE�Eタの取得中にエラーが発生しました: {e}")
            
    def event(self, event):
        """イベントハンドラ"""
        if event.type() == QEvent.User:
            # メインスレチE��でUIを更新
            self.update_form_with_data(self.cti_service.get_all_fields_data())
            return True
        return super().event(event)

    def validate_contractor_name(self, text):
        """
        契紁E��E��の入力を検証します、E        全角文字�Eみを許可し、半角文字が含まれてぁE��場合�E警告を表示します、E        
        Args:
            text (str): 入力されたチE��スチE        """
        import unicodedata
        
        # 空斁E���Eの場合�E検証をスキチE�E
        if not text:
            return
        
        # 半角文字が含まれてぁE��かチェチE��
        has_half_width = any(unicodedata.east_asian_width(char) in ['Na', 'H'] for char in text)
        
        if has_half_width:
            self.statusBar().showMessage("契紁E��E��は全角文字で入力してください", 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_furigana_input(self, text):
        """
        フリガナ�E入力を検証します、E        カタカナと長音記号のみを許可し、それ以外�E斁E��が含まれてぁE��場合�E警告を表示します、E        
        Args:
            text (str): 入力されたチE��スチE        """
        import re
        
        # 空斁E���Eの場合�E検証をスキチE�E
        if not text:
            return
        
        # カタカナと長音記号のみを許可する正規表現パターン
        katakana_pattern = r'^[ァ-ヶーヽヾ]+$'
        
        if not re.match(katakana_pattern, text):
            self.statusBar().showMessage("フリガナ�E全角カタカナで入力してください", 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_list_name(self, text):
        """
        リスト名の入力を検証します、E        半角英数字とハイフンのみを許可し、それ以外�E斁E��が含まれてぁE��場合�E警告を表示します、E        
        Args:
            text (str): 入力されたチE��スチE        """
        import re
        
        # 空斁E���Eの場合�E検証をスキチE�E
        if not text:
            return
        
        # 半角英数字とハイフンのみを許可する正規表現パターン
        pattern = r'^[A-Za-z0-9\-_]+$'
        
        if not re.match(pattern, text):
            self.statusBar().showMessage("リスト名は半角英数字とハイフンのみ使用できまぁE, 5000)
            # 背景色変更を削除
        else:
            # 背景色変更を削除
            self.statusBar().clearMessage()

    def validate_list_furigana(self):
        """リストフリガナ�EバリチE�Eション"""
        text = self.list_furigana_input.text()
        if not validate_furigana(text):
            # 背景色変更を削除
            QToolTip.showText(
                self.list_furigana_input.mapToGlobal(QPoint(0, 0)),
                "フリガナに数字や不適刁E��斁E��を含めることはできません",
                self.list_furigana_input
            )
        else:
            # 背景色変更を削除
            QToolTip.hideText()

    def reset_background_color(self):
        """
        フィールド�E背景色をリセチE��する
        
        入力�E有無に関わらず、対応する未入力警告�E背景色をリセチE��します、E        """
        sender = self.sender()
        if sender:
            sender.setStyleSheet("")

    def closeEvent(self, event):
        """ウィンドウを閉じる際�E処琁E""
        try:
            # すべてのアクチE��ブな検索スレチE��を停止
            if hasattr(self, 'active_search_threads'):
                for thread in self.active_search_threads:
                    if thread and thread.isRunning():
                        logging.info("アクチE��ブな検索スレチE��を停止しまぁE)
                        thread.stop()
                self.active_search_threads.clear()
            
            # 電話ボタン監視を停止
                
            # CTI状態監視を停止
            if hasattr(self, 'cti_status_monitor'):
                self.cti_status_monitor.stop_monitoring()
                
            event.accept()
        except Exception as e:
            logging.error(f"アプリケーション終亁E�E琁E��にエラー: {e}")
            event.accept()

    def update_preview(self):
        """プレビューを更新"""
        try:
            # 直接プレビューチE��ストを生�Eして設定！EEventを使わなぁE��E            preview_text = self.generate_preview_text()
            if preview_text and hasattr(self, 'preview_text'):
                self.preview_text.setText(preview_text)
        except Exception as e:
            logging.error(f"プレビュー更新中にエラー: {e}")

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # チE��スト�E力フィールド�Eクリア
        self.operator_input.clear()
        # 携帯電話番号入力エリアの参�Eを削除
        self.available_time_input.clear()  # 出めE��ぁE��間帯をクリア
        
        # 新しい携帯番号入力欁E�Eクリア
        self.mobile_part1_input.clear()
        self.mobile_part2_input.clear()
        self.mobile_part3_input.clear()
        self.mobile_pattern_combo.setCurrentText("②携帯なぁE)
        self.mobile_number_widget.hide()
        self.available_time_input.setText("携帯なぁE)
        
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
        # 受注老E��はクリアしなぁE��保持する�E�E        # self.order_person_input.clear()
        # 料��認識�EクリアしなぁE��保持する�E�E        # self.fee_input.clear()
        
        # 他番号、E��話機、禁止回線には初期値を設宁E        self.other_number_input.setText("なぁE)
        self.phone_device_input.setText("プッシュホン")
        self.forbidden_line_input.setText("なぁE)
        
        # NDと備老E��名義人との関係性�E�をクリア
        self.nd_input.clear()
        self.relationship_input.clear()
        # コンボ�EチE��スをデフォルト値に
        self.era_combo.setCurrentIndex(0)
        self.year_combo.setCurrentIndex(0)
        self.month_combo.setCurrentIndex(0)
        self.day_combo.setCurrentIndex(0)
        self.current_line_combo.setCurrentIndex(0)
        self.judgment_combo.setCurrentIndex(0)
        self.net_usage_combo.setCurrentIndex(0)
        self.family_approval_combo.setCurrentIndex(0)  # okがインチE��クス0になめE        # 結果ラベルをクリア
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
        # スクリーンショチE��ボタンをクリア
        self.update_screenshot_button()
        # プレビューもクリア
        self.preview_text.clear()

    def init_menu(self):
        """メニューバ�Eの初期匁E""
        menubar = self.menuBar()
        menubar.clear()
        
        # ファイルメニュー
        file_menu = menubar.addMenu("ファイル")
        
        # 終亁E        exit_action = file_menu.addAction("終亁E)
        exit_action.triggered.connect(self.close)
        
        # ヘルプメニュー
        help_menu = menubar.addMenu("ヘルチE)
        
        # アチE�EチE�Eト�E確誁E        update_action = help_menu.addAction("アチE�EチE�Eト�E確誁E)
        update_action.triggered.connect(self.show_update_dialog)
        
        # バ�Eジョン惁E��
        about_action = help_menu.addAction("バ�Eジョン惁E��")
        about_action.triggered.connect(self.show_about_dialog)
        
        # バ�Eジョン表示ラベル
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
        アチE�EチE�Eト設定ダイアログを表示する
        """
        dialog = UpdateDialog(self)
        dialog.settings_file = self.settings_file  # 設定ファイルのパスを渡ぁE        dialog.exec()
        
    def show_about_dialog(self):
        """
        バ�Eジョン惁E��ダイアログを表示する
        """
        msg = f"{APP_NAME} v{VERSION}\n\n"
        msg += "ライセンス: MIT License"
        QMessageBox.information(self, "バ�Eジョン惁E��", msg)

    def check_for_updates(self):
        """
        アチE�EチE�EトをチェチE��
        """
        try:
            # GitHubのAPIを使用して最新リリースを取征E            url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            response = requests.get(url)
            response.raise_for_status()
            latest_release = response.json()
            
            latest_version = latest_release["tag_name"].lstrip("v")
            current_version = VERSION
            
            if latest_version > current_version:
                # 新しいバ�Eジョンが利用可能
                msg = f"新しいバ�Eジョン v{latest_version} が利用可能です、En"
                msg += f"現在のバ�Eジョン: v{current_version}\n\n"
                msg += "更新しますか�E�E
                
                reply = QMessageBox.question(self, "アチE�EチE�EチE, msg,
                                          QMessageBox.StandardButton.Yes |
                                          QMessageBox.StandardButton.No)
                
                if reply == QMessageBox.StandardButton.Yes:
                    # アチE�EチE�Eトダイアログを作�Eして更新を実衁E                    dialog = UpdateDialog(self)
                    dialog.settings_file = self.settings_file
                    dialog.download_and_apply_update(latest_release)
        except Exception as e:
            logging.error(f"アチE�EチE�EトチェチE��中にエラー: {e}")

    def show_screenshot(self):
        """スクリーンショチE��を表示する"""
        try:
            if hasattr(self, 'screenshot_path') and self.screenshot_path:
                screenshot_path = self.screenshot_path
            else:
                screenshot_path = "debug_screenshot.png"
            
            if not os.path.exists(screenshot_path):
                QMessageBox.warning(
                    self,
                    "エラー",
                    "スクリーンショチE��ファイルが見つかりません、E
                )
                return
            
            # QPixmapを使用して画像を表示
            from PySide6.QtGui import QPixmap
            from PySide6.QtWidgets import QLabel, QDialog, QVBoxLayout, QScrollArea
            from PySide6.QtCore import Qt
            
            dialog = QDialog(self)
            dialog.setWindowTitle("スクリーンショチE�� - 提供判定結果")
            dialog.setMinimumSize(800, 600)
            layout = QVBoxLayout(dialog)
            
            # スクロールエリアを作�E
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            
            # ラベルを作�Eしてピクスマップを設宁E            label = QLabel()
            pixmap = QPixmap(screenshot_path)
            
            # 画像�Eアスペクト比を維持しながらスケーリング
            scaled_pixmap = pixmap.scaled(
                800,  # 最大幁E                4000,  # 十�Eな高さ�E�スクロール可能�E�E                Qt.AspectRatioMode.KeepAspectRatio,  # アスペクト比を維持E                Qt.TransformationMode.SmoothTransformation  # スムーズな変換
            )
            
            label.setPixmap(scaled_pixmap)
            
            # スクロールエリアにラベルを設宁E            scroll_area.setWidget(label)
            layout.addWidget(scroll_area)
            
            dialog.setLayout(layout)
            dialog.exec()
            
        except Exception as e:
            logging.error(f"スクリーンショチE��表示エラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"スクリーンショチE��の表示中にエラーが発生しました: {str(e)}"
            )

    def search_service_area(self):
        """提供エリア検索を開姁E""
        postal_code = self.postal_code_input.text().strip()
        address = self.address_input.text().strip()
        
        if not postal_code or not address:
            QMessageBox.warning(self, "入力エラー", "郵便番号と住所を�E力してください、E)
            return
        
        try:
            # 既存�EスレチE��とワーカーをクリーンアチE�E
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
            
            # プログレスバ�Eを表示
            self.progress_bar.setVisible(True)
            
            # 検索スチE�Eタスを更新
            self.area_result_label.setText("提供エリア: 検索を開始しまぁE..")
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
            
            # ワーカーを作�E
            self.worker = ServiceAreaSearchWorker(postal_code, address)
            self.worker.finished.connect(self.on_search_completed)
            self.worker.progress.connect(self.update_search_progress)
            
            # スレチE��を作�Eして検索を開姁E            self.thread = QThread()
            self.worker.moveToThread(self.thread)
            self.thread.started.connect(self.worker.run)
            self.thread.finished.connect(self.thread.deleteLater)
            self.thread.start()
            
        except Exception as e:
            logging.error(f"検索の開始に失敁E {str(e)}")
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

        # バックエンド�E琁E�Eキャンセル
        if hasattr(self, 'worker'):
            self.worker.cancel()
            # キャンセル完亁E��征E��ため、�Eタンとプログレスバ�Eはそ�Eまま維持E
    def reset_search_button(self):
        """検索ボタンを�E期状態に戻ぁE""
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
        # 検索ボタンのクリチE��イベントを允E��戻ぁE        self.area_search_btn.clicked.disconnect()
        self.area_search_btn.clicked.connect(self.search_service_area)

    def on_search_completed(self, result):
        """検索完亁E��の処琁E""
        # プログレスバ�Eを非表示
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
            # キャンセル完亁E��に検索ボタンを�E期状態に戻ぁE            self.reset_search_button()
            return
        
        # キャンセル以外�E完亁E��の処琁E        self.reset_search_button()
        
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
            self.judgment_combo.setCurrentText("◁E)
        elif status == "unavailable":
            self.area_result_label.setText("提供エリア: 未提侁E)
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
            self.judgment_combo.setCurrentText("ÁE)
        elif status == "apartment":
            # 雁E��住宁E�E場合�E明示皁E��表示
            self.area_result_label.setText("提供エリア: 雁E��住宁E��アパ�Eト�Eマンション等！E)
            self.area_result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #FF9800;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #E65100;
                }
            """)
            self.judgment_combo.setCurrentText("◁E)
        else:
            self.area_result_label.setText("提供エリア: 判定失敁E)
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

        # スクリーンショチE��の更新
        if "screenshot" in result:
            self.update_screenshot_button(result["screenshot"])

        # 詳細惁E��の表示
        if "details" in result and result.get("show_popup", True):
            details = result["details"]
            details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
            QMessageBox.information(self, "検索結果", details_text)

    def cleanup_thread(self):
        """
        スレチE��のクリーンアチE�Eを行う
        """
        try:
            if self.thread and isinstance(self.thread, QThread):
                if self.thread.isRunning():
                    self.thread.quit()
                    self.thread.wait()
                self.thread.deleteLater()
                self.thread = None
        except Exception as e:
            logging.error(f"スレチE��のクリーンアチE�E中にエラー: {str(e)}")
            # エラーが発生しても、スレチE��をNoneに設定して続衁E            self.thread = None

    def get_template(self):
        """フォーマットテンプレートを取得すめE""
        try:
            if not hasattr(self, 'format_template') or not self.format_template:
                if hasattr(self, 'settings') and 'format_template' in self.settings:
                    self.format_template = self.settings['format_template']
                else:
                    logging.error("format_templateを設定から読み込めません")
                    QMessageBox.warning(self, "エラー", "チE��プレートが設定されてぁE��せん、En設定画面でチE��プレートを設定してください、E)
                    return None
            return self.format_template
        except Exception as e:
            logging.error(f"チE��プレート取得中にエラー: {e}")
            return None

    def reconstruct_ui(self):
        """
        モード�Eり替え時にUIを�E構築すめE        """
        try:
            logging.info("UIの再構築を開始しまぁE)
            
            # 現在のモードに基づぁE��UIを�E構篁E            if self.current_mode == 'simple':
                self.init_simple_mode()
            else:
                self.init_easy_mode()
            
            # フォントサイズを�E適用
            font_size = self.settings.get('font_size', 10)
            self.set_font_size(font_size)
            
            logging.info("UIの再構築が完亁E��ました")
        except Exception as e:
            logging.error(f"UIの再構築中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"UIの再構築中にエラーが発生しました: {str(e)}")

    def update_search_progress(self, message):
        """検索の進捗状況を更新する"""
        try:
            # メチE��ージからパ�EセンチE�Eジを抽出
            import re
            match = re.search(r'\((\d+)%\)', message)
            if match:
                new_value = int(match.group(1))
                current_value = self.progress_bar.value()
                
                # アニメーションの設宁E                self.progress_animation.setStartValue(current_value)
                self.progress_animation.setEndValue(new_value)
                self.progress_animation.start()
                
            # プログレスバ�EとメチE��ージを表示
            self.progress_bar.setVisible(True)
            self.area_result_label.setText(message)
            self.area_result_label.setStyleSheet("color: #666666;")
            
        except Exception as e:
            logging.error(f"進捗更新中にエラー: {str(e)}")
            self.area_result_label.setText(message)

    def on_fee_combo_changed(self, text):
        """
        料��認識�Eコンボ�EチE��スが変更された時の処琁E        
        Args:
            text (str): 選択されたチE��スチE        """
        self.fee_input.setText(text)
        self.reset_background_color()

    def check_birth_date_age(self):
        """
        生年月日から年齢を計算し、E0歳以上�E場合に赤く表示する
        """
        try:
            # 現在の日付を取征E            now = datetime.datetime.now()
            current_year = now.year
            current_month = now.month
            current_day = now.day
            
            # 生年月日の惁E��を取征E            era = self.era_combo.currentText()
            year = int(self.year_combo.currentText())
            month = int(self.month_combo.currentText())
            day = int(self.day_combo.currentText())
            
            # 和暦を西暦に変換
            if era == "昭咁E:
                year = year + 1925
            elif era == "平戁E:
                year = year + 1988
            # 西暦の場合�Eそ�Eまま
            
            # 年齢を計箁E            age = current_year - year
            
            # 誕生日がまだ来てぁE��ぁE��合�E年齢めEつ減らぁE            if (month > current_month) or (month == current_month and day > current_day):
                age -= 1
            
            # 80歳以上かどぁE��をチェチE��
            is_over_80 = age >= 80
            
            # 背景色を設宁E            if is_over_80:
                style = "background-color: #FFEBEE;"  # 赤系の背景色
            else:
                style = ""  # チE��ォルト�E背景色
            
            # 吁E��ンボ�EチE��スにスタイルを適用
            self.era_combo.setStyleSheet(style)
            self.year_combo.setStyleSheet(style)
            self.month_combo.setStyleSheet(style)
            self.day_combo.setStyleSheet(style)
            
            # 80歳以上�E場合にログを�E劁E            if is_over_80:
                logging.info(f"80歳以上�E顧客が検�Eされました: {age}歳")
            
        except Exception as e:
            logging.error(f"年齢チェチE��中にエラー: {e}")

    def on_mobile_pattern_changed(self, text):
        """
        携帯番号パターンが変更された時の処琁E        
        Args:
            text (str): 選択されたチE��スチE        """
        if text == "①携帯ありで番号がわかる":
            # 携帯番号入力欁E��表示
            self.mobile_number_widget.show()
            # 入力欁E��クリア
            self.mobile_part1_input.clear()
            self.mobile_part2_input.clear()
            self.mobile_part3_input.clear()
            # フォーカスを最初�E入力欁E��設宁E            self.mobile_part1_input.setFocus()
        else:
            # 携帯番号入力欁E��非表示
            self.mobile_number_widget.hide()
            # パターンに応じてavailable_time_inputを更新
            if text == "②携帯なぁE:
                self.available_time_input.setText("携帯なぁE)
            elif text == "③携帯ありで番号がわからなぁE:
                self.available_time_input.setText("携帯不�E")
        
        # パターン変更時�Eみプレビューを更新�E�リアルタイム更新は削除�E�E
    def format_mobile_number_part(self):
        """
        携帯番号の吁E��刁E��変更された時の処琁E        数字�Eみを許可し、�E動的に次の入力欁E��フォーカスを移勁E        """
        sender = self.sender()
        if not sender:
            return
            
        # 数字以外�E斁E��を削除
        text = sender.text()
        formatted_text = ''.join(filter(str.isdigit, text))
        
        # 全角数字を半角に変換
        formatted_text = formatted_text.translate(str.maketrans('�E�１２３４５６７８！E, '0123456789'))
        
        if formatted_text != text:
            sender.setText(formatted_text)
        
        # 自動フォーカス移勁E        if sender == self.mobile_part1_input and len(formatted_text) == 3:
            self.mobile_part2_input.setFocus()
        elif sender == self.mobile_part2_input and len(formatted_text) == 4:
            self.mobile_part3_input.setFocus()
        
        # 携帯番号が完�Eしたらavailable_time_inputを更新
        self.update_available_time_from_mobile_parts()
    
    def update_available_time_from_mobile_parts(self):
        """
        携帯番号の吁E��刁E��ら完�Eな携帯番号を絁E��立ててavailable_time_inputを更新
        """
        part1 = self.mobile_part1_input.text().strip()
        part2 = self.mobile_part2_input.text().strip()
        part3 = self.mobile_part3_input.text().strip()
        
        if part1 and part2 and part3:
            # 3つの部刁E��すべて入力されてぁE��場吁E            mobile_number = f"{part1}-{part2}-{part3}"
            self.available_time_input.setText(mobile_number)
        elif part1 or part2 or part3:
            # 一部だけ�E力されてぁE��場合�E空にする
            self.available_time_input.setText("")
        
        # リアルタイムプレビュー更新を削除�E�営業コメント作�Eボタンを押した時�Eみ更新�E�E
    def on_cti_dialing_to_talking(self):
        """
        CTI状態が「発信中」�E「通話中」に変化した時�E自動�E琁E        
        1. 顧客惁E��を�E動取征E        2. 提供判定検索を�E動実衁E        """
        try:
            import time
            current_time = time.time()
            
            # 重褁E��行防止チェチE��
            if hasattr(self, 'is_auto_processing') and self.is_auto_processing:
                logging.info("CTI自動�E琁E��既に実行中のため、E��褁E��行をスキチE�EしまぁE)
                return
                
            # 前回実行から短時間の場合�EスキチE�E
            if hasattr(self, 'last_auto_processing_time'):
                time_since_last = current_time - self.last_auto_processing_time
                if time_since_last < 3.0:  # 3秒以冁E�E重褁E��行を防ぁE                    logging.info(f"前回の自動�E琁E��ら{time_since_last:.2f}秒しか経過してぁE��ぁE��め、E��褁E��行をスキチE�EしまぁE)
                    return
            
            # 処琁E��フラグを設宁E            self.is_auto_processing = True
            self.last_auto_processing_time = current_time
            
            logging.info("CTI状態変化による自動�E琁E��開始しまぁE)
            
            # 1. 顧客惁E��取得を実行（既存�Efetch_cti_dataメソチE��を呼び出し！E            logging.info("1. 顧客惁E��の自動取得を開姁E)
            self.fetch_cti_data()
            
            # 2. 顧客惁E��取得が完亁E��てから提供判定検索を実衁E            # シグナルを使用してメインスレチE��で実行（スレチE��セーフ！E            import threading
            def delayed_trigger():
                try:
                    self.trigger_auto_search.emit()
                    logging.debug("提供判定検索のシグナルを送信しました")
                except Exception as e:
                    logging.error(f"シグナル送信中にエラー: {str(e)}")
                finally:
                    # 処琁E��亁E��にフラグをリセチE��
                    time.sleep(2.0)  # 2秒後にリセチE��
                    self.is_auto_processing = False
                    logging.debug("自動�E琁E��ラグをリセチE��しました")
                    
            timer = threading.Timer(1.0, delayed_trigger)
            timer.daemon = True
            timer.start()
            
        except Exception as e:
            logging.error(f"CTI自動�E琁E��にエラーが発甁E {str(e)}")
            # エラー時もフラグをリセチE��
            if hasattr(self, 'is_auto_processing'):
                self.is_auto_processing = False
    
    @Slot()
    def auto_search_service_area(self):
        """
        自動提供判定検索を実衁E        """
        try:
            logging.info("2. 提供判定検索の自動実行を開姁E)
            
            # 郵便番号と住所が�E力されてぁE��かチェチE��
            postal_code = ""
            address = ""
            
            # シンプルモードと誘導モードで異なる�E力フィールドを参�E
            if hasattr(self, 'postal_code_input'):
                postal_code = self.postal_code_input.text().strip()
            if hasattr(self, 'address_input'):
                address = self.address_input.text().strip()
                
            # 入力データが不足してぁE��場合�E処琁E            if not postal_code or not address:
                logging.warning("郵便番号また�E住所が未入力�Eため、提供判定検索をスキチE�Eしました")
                return
                
            # 既存�E検索メソチE��を呼び出ぁE            self.search_service_area()
            
            logging.info("CTI自動�E琁E��完亁E��ました")
            
        except Exception as e:
            logging.error(f"自動提供判定検索中にエラーが発甁E {str(e)}")
            
    def closeEvent(self, event):
        """ウィンドウを閉じる際�E処琁E""
        try:
            # すべてのアクチE��ブな検索スレチE��を停止
            if hasattr(self, 'active_search_threads'):
                for thread in self.active_search_threads:
                    if thread and thread.isRunning():
                        logging.info("アクチE��ブな検索スレチE��を停止しまぁE)
                        thread.stop()
                self.active_search_threads.clear()
            
            # 電話ボタン監視を停止
                
            # CTI状態監視を停止
            if hasattr(self, 'cti_status_monitor'):
                self.cti_status_monitor.stop_monitoring()
                
            event.accept()
        except Exception as e:
            logging.error(f"アプリケーション終亁E�E琁E��にエラー: {e}")
            event.accept()


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
            {"message": "住所惁E��を解析中...", "weight": 5},
            {"message": "NTT西日本のサイトにアクセス中...", "weight": 10},
            {"message": "郵便番号を�E力中...", "weight": 15},
            {"message": "住所を選択中...", "weight": 20},
            {"message": "番地を�E力中...", "weight": 20},
            {"message": "号を�E力中...", "weight": 20},
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
            message (str, optional): カスタムメチE��ージ。指定がなぁE��合�E定義済みメチE��ージを使用
        """
        try:
            if message is None and self._current_step < len(self._progress_steps):
                step_info = self._progress_steps[self._current_step]
                message = step_info["message"]
                # 現在のスチE��プ�E重みに基づぁE��進捗を計箁E                self._accumulated_progress += step_info["weight"]
            elif message:
                # メチE��ージに含まれるパ�EセンチE�Eジを抽出
                import re
                percent_match = re.search(r'(\d+)%', message)
                if percent_match:
                    self._accumulated_progress = int(percent_match.group(1))
                else:
                    # メチE��ージにパ�EセンチE�Eジが含まれてぁE��ぁE��合�E、次のスチE��プに進む
                    if self._current_step < len(self._progress_steps):
                        self._accumulated_progress += self._progress_steps[self._current_step]["weight"]
            
            # 進捗率を計算（最大95%まで�E�E            progress_percent = min(int((self._accumulated_progress / self._total_weight) * 95), 95)
            
            # 進捗メチE��ージを生戁E            if "%" not in message:
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

            # 検索を実衁E            self._update_progress()  # 初期進捗を表示
            result = search_service_area(
                self.postal_code,
                self.address,
                progress_callback=progress_callback
            )
            
            if self._is_cancelled:
                raise CancellationError("検索がキャンセルされました")
            
            # 検索完亁E��に100%を表示
            if result.get("status") == "available":
                self.progress.emit("提供可能でぁE(100%)")
            elif result.get("status") == "unavailable":
                self.progress.emit("提供不可でぁE(100%)")
            else:
                self.progress.emit("検索が完亁E��ました (100%)")
            self.finished.emit(result)
            
        except CancellationError as e:
            logging.info("検索がキャンセルされました")
            self.progress.emit("検索がキャンセルされました (0%)")
            self.finished.emit({
                "status": "cancelled",
                "message": "検索がキャンセルされました"
            })
        except Exception as e:
            logging.error(f"検索処琁E��にエラーが発甁E {str(e)}")
            self.progress.emit("エラーが発生しました (0%)")
            self.finished.emit({
                "status": "error",
                "message": f"検索処琁E��にエラーが発甁E {str(e)}"
            })

class CancellationError(Exception):
    """検索キャンセル時に発生する例夁E""
    pass

    def save_input_data(self, input_data):
        """
        入力データを保存すめE        
        Args:
            input_data (dict): 保存する�E力データ
        """
        try:
            # 保存�EチE��レクトリの作�E
            save_dir = "input_data"
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            
            # ファイル名�E生�E�E�タイムスタンプ付き�E�E            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"input_data_{timestamp}.json"
            filepath = os.path.join(save_dir, filename)
            
            # チE�Eタの保孁E            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(input_data, f, ensure_ascii=False, indent=4)
            
            logging.info(f"入力データを保存しました: {filepath}")
            QMessageBox.information(self, "完亁E, "入力データの保存が完亁E��ました、E)
            
        except Exception as e:
            logging.error(f"入力データの保存中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"入力データの保存中にエラーが発生しました: {e}")

    @Slot()
    def generate_preview_text(self):
        """プレビューチE��ストを生�Eする"""
        try:
            logging.info("プレビューチE��スト�E生�Eを開姁E)
            
            # フォーマットテンプレート�E取征E            format_template = self.settings.get('format_template', '')
            logging.info(f"フォーマットテンプレーチE {format_template}")
            
            # チE��プレートが空の場合�Eエラー
            if not format_template:
                logging.error("フォーマットテンプレートが設定されてぁE��せん")
                QMessageBox.warning(self, "警呁E, "フォーマットテンプレートが設定されてぁE��せん、En設定画面でチE��プレートを設定してください、E)
                return None
            
            # チE�Eタの初期匁E            data = {}
            
            # 吁E�E力フィールドからデータを取得し、末尾のスペ�Eスを削除
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
            
            # コンボ�EチE��スからチE�Eタを取征E            if hasattr(self, 'current_line_combo'):
                data['current_line'] = self.current_line_combo.currentText().rstrip()
            if hasattr(self, 'order_date_input'):
                data['order_date'] = self.order_date_input.text().rstrip()
            if hasattr(self, 'judgment_combo'):
                data['judgment'] = self.judgment_combo.currentText().rstrip()
            
            # チE�Eタが空の場合�Eエラー
            if not data:
                logging.error("プレビュー生�Eに忁E��なチE�Eタが取得できません")
                return None
            
            # チE��プレート�E置揁E            preview_text = format_template
            for key, value in data.items():
                placeholder = f"{{{key}}}"
                preview_text = preview_text.replace(placeholder, str(value or ''))
                logging.debug(f"プレースホルダー {placeholder} めE{value} に置揁E)
            
            logging.info("プレビューチE��スト�E生�Eが完亁E)
            
            # プレビューチE��ストを設宁E            if hasattr(self, 'preview_text'):
                self.preview_text.setText(preview_text)
            
            return preview_text
            
        except Exception as e:
            logging.error(f"プレビュー生�E中にエラー: {e}", exc_info=True)
            return None

    def load_settings(self):
        """設定を読み込む"""
        try:
            # 設定ファイルの読み込み
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                    logging.info("設定ファイルを読み込みました")
                    logging.info(f"設定�E容: {self.settings}")
            else:
                self.settings = {}
                logging.warning("設定ファイルが存在しません")
            
            # フォーマットテンプレート�E確誁E            if 'format_template' not in self.settings or not self.settings['format_template']:
                logging.error("フォーマットテンプレートが設定されてぁE��せん")
                QMessageBox.warning(self, "警呁E, "フォーマットテンプレートが設定されてぁE��せん、En設定画面でチE��プレートを設定してください、E)
                return
            
            logging.info(f"フォーマットテンプレーチE {self.settings['format_template']}")
            
            # フォントサイズの設宁E            font_size = self.settings.get('font_size', 10)
            logging.info(f"フォントサイズめE{font_size} に設定しました")
            
            
        except Exception as e:
            logging.error(f"設定�E読み込み中にエラーが発生しました: {e}", exc_info=True)
            self.settings = {}
            QMessageBox.critical(self, "エラー", f"設定�E読み込み中にエラーが発生しました: {e}")

