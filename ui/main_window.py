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
import threading
from urllib.parse import quote
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QLabel, QLineEdit, QComboBox, QPushButton,
                              QTextEdit, QGroupBox, QMessageBox, QScrollArea,
                              QApplication, QToolTip, QSplitter, QMenuBar, QMenu,
                              QSizePolicy, QProgressBar, QListView)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread, QPropertyAnimation, QEasingCurve, QRect, QPoint, QMetaObject, Q_ARG
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
from services.phone_button_monitor import PhoneButtonMonitor
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


class CancelWorker(QObject):
    """
    キャンセル処理を並列実行するワーカークラス
    TelephoneTeikyou-crossの高速キャンセル方式を適用
    """
    finished = Signal()
    progress = Signal(str)  # 進捗報告用シグナルを追加
    
    def __init__(self, worker_to_cancel=None):
        """
        ワーカーの初期化
        
        Args:
            worker_to_cancel: キャンセル対象のワーカーオブジェクト
        """
        super().__init__()
        self.worker_to_cancel = worker_to_cancel
        self.is_cancelled = False
    
    def cancel(self):
        """
        キャンセル処理を実行（cleanup_threadから呼び出し用）
        """
        self.is_cancelled = True
        if self.worker_to_cancel and hasattr(self.worker_to_cancel, 'cancel'):
            self.worker_to_cancel.cancel()
    
    def run(self):
        """
        軽量化されたキャンセル処理を実行
        TelephoneTeikyou-crossの高速方式を採用
        """
        try:
            logging.info("★★★ 並列キャンセル処理を開始します ★★★")
            self.progress.emit("キャンセル処理を実行中...")
            
            # ワーカーのキャンセル処理
            if self.worker_to_cancel and hasattr(self.worker_to_cancel, 'cancel'):
                logging.info("- ワーカーのキャンセル処理を実行")
                self.worker_to_cancel.cancel()
                
            # キャンセルフラグの設定
            try:
                from services.area_search import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
                logging.info("- エリア検索のキャンセルフラグを設定しました")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
                logging.info("- 東日本エリア検索のキャンセルフラグを設定しました")
            except ImportError:
                pass
            
            # 短時間の待機で完了確認（TelephoneTeikyou-crossと同様）
            import time
            time.sleep(0.5)  # 500ms待機
            
            self.progress.emit("キャンセル処理完了")
            logging.info("★★★ 並列キャンセル処理が完了しました ★★★")
            
        except Exception as e:
            logging.error(f"並列キャンセル処理中にエラー: {str(e)}")
        finally:
            self.is_cancelled = True
            self.finished.emit()


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
    
    # カスタムシグナル：CTI自動処理用
    trigger_auto_search = Signal()
    
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
        
        # キャンセル処理関連の初期化
        self.cancel_worker = None
        self.cancel_thread = None
        self.cancel_timer = None
        
        # ログ設定
        self.setup_logging()
        
        # 設定ファイルのパスを設定（exe直下 > CWD > ソース直下）
        if getattr(sys, 'frozen', False):
            self.settings_file = os.path.join(os.path.dirname(sys.executable), 'settings.json')
        else:
            # ソース直下を既定としつつ、存在チェック
            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
            try:
                cwd_candidate = os.path.join(os.getcwd(), 'settings.json')
                if os.path.exists(cwd_candidate):
                    self.settings_file = cwd_candidate
            except Exception:
                pass
        
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # 設定を読み込む
        self.settings = {}
        
        # 設定ファイルが存在しない場合は新規作成（format や CTI のデフォルトのみ）
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
        self.era_combo.addItems(["西暦", "平成", "昭和"])
        
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1926, datetime.datetime.now().year + 1)])
        
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
        
        # 電話ボタン監視の初期化と開始
        self.phone_monitor = PhoneButtonMonitor(self.fetch_cti_data)
        self.phone_monitor.start_monitoring()
        
        # CTI状態監視の初期化と開始（設定に基づいて制御）
        cti_monitoring_enabled = self.settings.get('enable_cti_monitoring', True)
        logging.info(f"CTI監視設定: {cti_monitoring_enabled}")
        
        if cti_monitoring_enabled:
            if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                self.cti_status_monitor = CTIStatusMonitor(
                    on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
                    on_call_ended_callback=self.on_cti_call_ended,
                    on_talking_started_callback=self.on_cti_talking_started,
                    on_cancel_processing_callback=self.on_cancel_processing_request
                )
                self.cti_status_monitor.start_monitoring()
                logging.info("CTI状態監視を開始しました")
                
                # CTI自動処理用のシグナル・スロット接続（重複接続を防ぐ）
                if not self.trigger_auto_search.isSignalConnected(self.trigger_auto_search, self.auto_search_service_area):
                    self.trigger_auto_search.connect(self.auto_search_service_area)
        else:
            logging.info("CTI監視が設定で無効になっています")
            self.cti_status_monitor = None
        
        # 自動処理の重複実行防止用フラグ
        if not hasattr(self, 'is_auto_processing'):
            self.is_auto_processing = False
            self.last_auto_processing_time = 0
        
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
            # 初期設定ファイル生成かどうかをチェック
            is_initial_setup = not os.path.exists(self.settings_file)
            
            # 設定ファイルの読み込み
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            
            # モード設定を更新
            settings['mode'] = mode
            settings['show_mode_selection'] = show_again
            
            # 初期設定ファイル生成時のデフォルト値を設定
            if is_initial_setup:
                logging.info("初期設定ファイルを生成します（CTI監視設定を含む）")
                # デフォルトのフォーマットテンプレート
                default_format = """対応者（お客様の名前）：{operator}
工事希望日
★出やすい時間帯：{available_time} 
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

他番号：{other_number}
電話機：{phone_device}
禁止回線：{forbidden_line}
ND：{nd}

備考：{relationship}
お客様が今使っている回線：アナログ
案内料金：2500円
"""
                
                # 初期設定のデフォルト値を設定
                settings.update({
                    'format_template': default_format,
                    'font_size': 9,
                    'delay_seconds': 0,
                    'browser_settings': {
                        'headless': True,
                        'disable_images': True,
                        'show_popup': True,
                        'auto_close': True,
                        'page_load_timeout': 30,
                        'script_timeout': 30
                    },
                    # CTI監視設定のデフォルト値（オンに設定）
                    'enable_cti_monitoring': True,
                    'enable_auto_cti_processing': True,
                    'cti_monitor_interval': 0.5,
                    'cti_auto_processing_cooldown': 3.0,
                    'call_duration_threshold': 0
                })
            
            # 設定ファイルに保存
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            logging.info(f"モード設定を保存しました: mode={mode}, show_mode_selection={show_again}")
            if is_initial_setup:
                logging.info("CTI監視設定を有効にして初期設定ファイルを生成しました")
                
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

        # 転記機能の可用性チェック（共有トークン未設定時は案内を出す）
        try:
            gf = (self.settings or {}).get('googleFormPosting') or {}
            if not gf or not gf.get('tokenValue'):
                logging.warning("共有トークン未設定のため、転記機能は使用できません。設定からトークンを追加してください。")
        except Exception:
            pass
        
        # フォントサイズの適用
        self.apply_font_size()
        
        # CTI連携サービスの初期化
        self.cti_service = OneClickService()
        
        # 電話ボタン監視の初期化と開始
        self.phone_monitor = PhoneButtonMonitor(self.fetch_cti_data)
        self.phone_monitor.start_monitoring()
        
        # CTI状態監視の初期化と開始
        self.cti_status_monitor = CTIStatusMonitor(
            on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
            on_call_ended_callback=self.on_cti_call_ended,
            on_talking_started_callback=self.on_cti_talking_started,
            on_cancel_processing_callback=self.on_cancel_processing_request
        )
        self.cti_status_monitor.start_monitoring()
        
        # CTI自動処理用のシグナル・スロット接続
        self.trigger_auto_search.connect(self.auto_search_service_area)
        
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
        
        # CTI状態監視の初期化と開始
        self.cti_status_monitor = CTIStatusMonitor(
            on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
            on_call_ended_callback=self.on_cti_call_ended,
            on_talking_started_callback=self.on_cti_talking_started,
            on_cancel_processing_callback=self.on_cancel_processing_request
        )
        self.cti_status_monitor.start_monitoring()
        
        # CTI自動処理用のシグナル・スロット接続
        self.trigger_auto_search.connect(self.auto_search_service_area)

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
            status = result.get("status")
            if status == "available":
                self.update_judgment_result("提供可能")
            elif status == "unavailable":
                self.update_judgment_result("提供エリア外")
            elif status == "apartment":
                # 集合住宅の場合は明示的に表示
                self.update_judgment_result("集合住宅（アパート・マンション等）")
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
        self.spreadsheet_btn = QPushButton("スプレッドシート転記")
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
        
        # 出やすい時間帯（携帯番号入力）
        input_layout.addWidget(QLabel("出やすい時間帯（携帯番号）"))
        
        # 携帯番号パターン選択
        self.mobile_pattern_combo = CustomComboBox()
        self.mobile_pattern_combo.addItems(["①携帯ありで番号がわかる", "②携帯なし", "③携帯ありで番号がわからない"])
        self.mobile_pattern_combo.currentTextChanged.connect(self.on_mobile_pattern_changed)
        input_layout.addWidget(self.mobile_pattern_combo)
        
        # 携帯番号入力欄（3つの枠）
        self.mobile_number_widget = QWidget()
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
        
        # 出やすい時間帯入力欄
        self.time_preference_widget = QWidget()
        time_preference_layout = QHBoxLayout(self.time_preference_widget)
        time_preference_layout.setContentsMargins(0, 0, 0, 0)
        
        time_preference_layout.addWidget(QLabel("出やすい時間帯："))
        self.time_preference_input = QLineEdit()
        self.time_preference_input.setPlaceholderText("例：午前中")
        self.time_preference_input.textChanged.connect(self.update_available_time_from_mobile_parts)
        time_preference_layout.addWidget(self.time_preference_input)
        
        input_layout.addWidget(self.time_preference_widget)
        self.time_preference_widget.hide()  # 初期状態では非表示
        
        # 従来の出やすい時間帯入力欄（互換性のため保持、非表示）
        self.available_time_input = QLineEdit()
        self.available_time_input.hide()
        
        # 初期状態の設定
        self.mobile_pattern_combo.setCurrentText("②携帯なし")
        self.mobile_number_widget.hide()
        self.available_time_input.setText("携帯なし")
        
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
        self.era_combo.addItems(["西暦", "平成", "昭和"])
        self.era_combo.currentTextChanged.connect(self.check_birth_date_age)
        birth_date_layout.addWidget(self.era_combo)
        
        # 年選択
        self.year_combo = NoWheelComboBox()
        self.year_combo.addItems([str(i) for i in range(1926, datetime.datetime.now().year + 1)])
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

        # 年齢情報表示用ラベル（生年月日入力欄の下に配置）
        self.age_info_label = QLabel("")
        self.age_info_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                padding-top: 2px;
            }
        """)
        self.age_info_label.hide()  # 初期状態は非表示
        input_layout.addWidget(self.age_info_label)
        
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

        # 結果表示とリロードボタンを横並びにするコンテナ
        result_row_container = QWidget()
        result_row_layout = QHBoxLayout(result_row_container)
        result_row_layout.setContentsMargins(0, 0, 0, 0)
        result_row_layout.setSpacing(5)

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
        result_row_layout.addWidget(self.area_result_label)

        # リロードボタン（初期状態では非表示）
        self.reload_btn = QPushButton("リロード")
        self.reload_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                font-size: 12px;
                border-radius: 4px;
                min-width: 60px;
                max-width: 80px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:pressed {
                background-color: #E65100;
            }
        """)
        self.reload_btn.clicked.connect(self.reload_application)
        self.reload_btn.hide()  # 初期状態では非表示
        result_row_layout.addWidget(self.reload_btn)

        # 再起動ボタン（初期状態では非表示）
        self.restart_btn = QPushButton("再起動")
        self.restart_btn.setStyleSheet("""
            QPushButton {
                background-color: #F44336;
                color: white;
                border: none;
                padding: 5px 10px;
                text-align: center;
                font-size: 12px;
                border-radius: 4px;
                min-width: 60px;
                max-width: 80px;
            }
            QPushButton:hover {
                background-color: #D32F2F;
            }
            QPushButton:pressed {
                background-color: #B71C1C;
            }
        """)
        self.restart_btn.clicked.connect(self.restart_application)
        self.restart_btn.hide()  # 初期状態では非表示
        result_row_layout.addWidget(self.restart_btn)

        area_result_layout.addWidget(result_row_container)

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
        # 変更前のCTI監視設定を保存
        old_cti_monitoring = self.settings.get('enable_cti_monitoring', True)
        
        dialog = SettingsDialog(self)
        if dialog.exec():
            # ダイアログがOKで閉じられた場合、設定を再読み込み
            self.load_settings()
            
            # CTI監視設定が変更された場合の処理
            new_cti_monitoring = self.settings.get('enable_cti_monitoring', True)
            if old_cti_monitoring != new_cti_monitoring:
                logging.info(f"CTI監視設定が変更されました: {old_cti_monitoring} → {new_cti_monitoring}")
                
                if new_cti_monitoring:
                    # CTI監視を有効にする
                    if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                        self.cti_status_monitor = CTIStatusMonitor(
                            on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
                            on_call_ended_callback=self.on_cti_call_ended,
                            on_talking_started_callback=self.on_cti_talking_started,
                            on_cancel_processing_callback=self.on_cancel_processing_request
                        )
                        self.cti_status_monitor.start_monitoring()
                        logging.info("CTI状態監視を開始しました")
                        
                        # CTI自動処理用のシグナル・スロット接続
                        if not self.trigger_auto_search.isSignalConnected(self.trigger_auto_search, self.auto_search_service_area):
                            self.trigger_auto_search.connect(self.auto_search_service_area)
                    elif hasattr(self.cti_status_monitor, 'start_monitoring'):
                        self.cti_status_monitor.start_monitoring()
                        logging.info("CTI状態監視を再開しました")
                else:
                    # CTI監視を無効にする
                    if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                        self.cti_status_monitor.stop_monitoring()
                        logging.info("CTI状態監視を停止しました")
            
            # 既存のCTI監視サービスの設定を更新
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                if hasattr(self.cti_status_monitor, 'update_settings'):
                    self.cti_status_monitor.update_settings()
                    logging.info("CTI監視サービスの設定を更新しました")
            
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
            logging.info("アプリケーション終了処理を開始します")
            
            # 検索スレッドとワーカーをクリーンアップ
            self.cleanup_thread()
            
            # キャンセルスレッドとワーカーをクリーンアップ
            if hasattr(self, 'cancel_worker') and self.cancel_worker:
                try:
                    self.cancel_worker.deleteLater()
                except:
                    pass
                self.cancel_worker = None
                
            if hasattr(self, 'cancel_thread') and self.cancel_thread:
                try:
                    if hasattr(self.cancel_thread, 'isRunning') and callable(self.cancel_thread.isRunning):
                        if self.cancel_thread.isRunning():
                            self.cancel_thread.quit()
                            self.cancel_thread.wait(2000)
                    self.cancel_thread.deleteLater()
                except Exception as e:
                    logging.error(f"キャンセルスレッドの終了処理エラー: {str(e)}")
                self.cancel_thread = None
            
            # キャンセルタイマーを停止
            if hasattr(self, 'cancel_timer') and self.cancel_timer:
                try:
                    self.cancel_timer.stop()
                    self.cancel_timer.deleteLater()
                except:
                    pass
                self.cancel_timer = None
            
            # すべてのアクティブな検索スレッドを停止
            if hasattr(self, 'active_search_threads'):
                for thread in self.active_search_threads:
                    try:
                        if thread and hasattr(thread, 'isRunning') and callable(thread.isRunning):
                            if thread.isRunning():
                                logging.info("アクティブな検索スレッドを停止します")
                                if hasattr(thread, 'stop'):
                                    thread.stop()
                                thread.quit()
                                thread.wait(1000)
                    except Exception as e:
                        logging.error(f"検索スレッドの停止エラー: {str(e)}")
                self.active_search_threads.clear()
            
            # 電話ボタン監視を停止
            if hasattr(self, 'phone_monitor'):
                try:
                    self.phone_monitor.stop_monitoring()
                except Exception as e:
                    logging.error(f"電話ボタン監視の停止エラー: {str(e)}")
                
            # CTI状態監視を停止
            if hasattr(self, 'cti_status_monitor'):
                try:
                    self.cti_status_monitor.stop_monitoring()
                except Exception as e:
                    logging.error(f"CTI状態監視の停止エラー: {str(e)}")
            
            logging.info("アプリケーション終了処理が完了しました")
            event.accept()
            
        except Exception as e:
            logging.error(f"アプリケーション終了処理中にエラー: {e}")
            event.accept()

    def update_preview(self):
        """プレビューを更新"""
        try:
            # 直接プレビューテキストを生成して設定（QEventを使わない）
            preview_text = self.generate_preview_text()
            if preview_text and hasattr(self, 'preview_text'):
                self.preview_text.setText(preview_text)
        except Exception as e:
            logging.error(f"プレビュー更新中にエラー: {e}")

    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        # テキスト入力フィールドのクリア
        self.operator_input.clear()
        # 携帯電話番号入力エリアの参照を削除
        self.available_time_input.clear()  # 出やすい時間帯をクリア
        
        # 新しい携帯番号入力欄のクリア
        self.mobile_part1_input.clear()
        self.mobile_part2_input.clear()
        self.mobile_part3_input.clear()
        self.mobile_pattern_combo.setCurrentText("②携帯なし")
        self.mobile_number_widget.hide()
        self.available_time_input.setText("携帯なし")
        
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
            # CTI自動処理中かどうかを判定
            is_auto_processing = hasattr(self, 'is_auto_processing') and self.is_auto_processing
            
            if is_auto_processing:
                # 自動処理ではメッセージボックスを表示しない
                logging.warning("CTI自動処理: 郵便番号または住所が未入力のため検索をスキップします")
                return
            else:
                # 手動実行時はメッセージボックスを表示
                QMessageBox.warning(self, "入力エラー", "郵便番号と住所を入力してください。")
                return
        
        try:
            # ★★★ 検索開始前にキャンセルフラグを必ずクリア ★★★
            try:
                from services.area_search import clear_cancel_flag
                clear_cancel_flag()
                logging.info("★★★ 検索開始前：西日本エリア検索のキャンセルフラグをクリアしました ★★★")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import clear_cancel_flag
                clear_cancel_flag()
                logging.info("★★★ 検索開始前：東日本エリア検索のキャンセルフラグをクリアしました ★★★")
            except ImportError:
                pass
            
            # 既存のスレッドとワーカーをクリーンアップ
            self.cleanup_thread()
            
            # CTI監視システムに検索開始を通知（処理フラグ設定）
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                try:
                    with self.cti_status_monitor.processing_lock:
                        self.cti_status_monitor.is_processing = True
                        # 提供判定開始時刻を記録（アクションボタンクリック検出用）
                        import time
                        if self.cti_status_monitor.talking_start_time == 0:
                            self.cti_status_monitor.talking_start_time = time.time()
                        logging.info("★★★ CTI監視システムに提供判定検索開始を通知しました ★★★")
                        logging.info(f"- 処理フラグ: {self.cti_status_monitor.is_processing}")
                        logging.info(f"- 検索開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                except Exception as cti_error:
                    logging.error(f"CTI監視システムの処理フラグ設定中にエラー: {str(cti_error)}")
            
            # 検索ボタンを即座にキャンセルボタンに変更
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
            
            # シグナル接続を安全に行う
            try:
                self.area_search_btn.clicked.disconnect()
            except:
                pass  # 既に接続されていない場合のエラーを無視
            
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
            
            logging.info("提供判定検索を開始しました（キャンセルボタンは即座に有効）")
            
        except Exception as e:
            logging.error(f"検索の開始に失敗: {str(e)}")
            self.reset_search_button()
            
            # 検索開始失敗時もCTI監視システムの処理フラグをリセット
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                try:
                    with self.cti_status_monitor.processing_lock:
                        self.cti_status_monitor.is_processing = False
                        self.cti_status_monitor.talking_start_time = 0
                        logging.info("検索開始失敗により、CTI監視システムの処理フラグをリセットしました")
                except Exception as cti_error:
                    logging.error(f"CTI監視システムの処理フラグリセット中にエラー: {str(cti_error)}")
            
            # CTI自動処理中かどうかを判定
            is_auto_processing = hasattr(self, 'is_auto_processing') and self.is_auto_processing
            
            if not is_auto_processing:
                # 手動実行時のみメッセージボックスを表示
                QMessageBox.critical(self, "エラー", f"検索の開始に失敗しました: {str(e)}")

    def cancel_search(self):
        """
        検索をキャンセルする
        TelephoneTeikyou-crossの高速キャンセル方式を適用
        """
        try:
            logging.info("=== 検索キャンセル処理開始 ===")
            
            # キャンセル中の状態をUIに即時反映
            if hasattr(self, 'area_search_btn'):
                self.area_search_btn.setEnabled(False)
                self.area_search_btn.setText("キャンセル中...")
            
            # 結果ラベルを即座に更新
            if hasattr(self, 'area_result_label'):
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
            
            # リロード/再起動ボタンを表示（キャンセル中状態で問題が発生した場合の対処用）
            if hasattr(self, 'reload_btn'):
                self.reload_btn.show()
            if hasattr(self, 'restart_btn'):
                self.restart_btn.show()
            
            # バックエンド処理のキャンセル
            if hasattr(self, 'worker') and self.worker:
                self.worker.cancel()
                
            # 並列キャンセル処理を開始（メインスレッドから実行）
            if threading.current_thread() == threading.main_thread():
                # メインスレッドから呼び出された場合
                self.start_parallel_cancel()
                # 2秒後にキャンセルが完了していない場合の自動リセット処理（高速化のため短縮）
                QTimer.singleShot(5000, self.check_cancel_timeout)
            else:
                # 別スレッドから呼び出された場合はシグナルを使用
                QMetaObject.invokeMethod(self, "start_parallel_cancel", Qt.QueuedConnection)
                QMetaObject.invokeMethod(self, "schedule_cancel_timeout", Qt.QueuedConnection)
                
        except Exception as e:
            logging.error(f"キャンセル処理中にエラー: {str(e)}")
            # エラー時もUI状態をリセット
            self.reset_search_button()
    
    @Slot()
    def schedule_cancel_timeout(self):
        """
        キャンセルタイムアウト処理をスケジュール（メインスレッド専用）
        """
        try:
            QTimer.singleShot(5000, self.check_cancel_timeout)  # 高速化のため2秒に短縮
        except Exception as e:
            logging.error(f"キャンセルタイムアウトスケジュール中にエラー: {str(e)}")
    
    @Slot()
    def start_parallel_cancel(self):
        """
        並列キャンセル処理を開始
        UIの応答性を維持するため、キャンセル処理を別スレッドで実行
        TelephoneTeikyou-crossの軽量方式を採用
        """
        try:
            # 既存のキャンセルスレッドがある場合は終了
            self.cleanup_cancel_thread()
            
            # キャンセルワーカーを作成
            self.cancel_worker = CancelWorker(worker_to_cancel=self.worker)
            
            # キャンセル用スレッドを作成
            self.cancel_thread = QThread()
            self.cancel_worker.moveToThread(self.cancel_thread)
            
            # シグナルとスロットを接続
            self.cancel_thread.started.connect(self.cancel_worker.run)
            self.cancel_worker.finished.connect(self.on_cancel_completed)
            self.cancel_worker.progress.connect(self.update_cancel_progress)
            self.cancel_worker.finished.connect(self.cancel_thread.quit)
            self.cancel_worker.finished.connect(self.cancel_worker.deleteLater)
            self.cancel_thread.finished.connect(self.cancel_thread.deleteLater)
            
            # キャンセルスレッドを開始
            self.cancel_thread.start()
            logging.info("並列キャンセル処理を開始しました")
            
        except Exception as e:
            logging.error(f"並列キャンセル処理の開始に失敗: {str(e)}")
            # フォールバック：同期的にキャンセル処理を実行
            self.fallback_cancel()
    
    def cleanup_cancel_thread(self):
        """
        キャンセル用スレッドとワーカーをクリーンアップ
        """
        try:
            # ワーカーをキャンセル
            if hasattr(self, 'cancel_worker') and self.cancel_worker:
                try:
                    self.cancel_worker.cancel()
                    logging.debug("ワーカーをキャンセルしました")
                except Exception as e:
                    logging.error(f"ワーカーキャンセル中にエラー: {str(e)}")
                
            # スレッドを終了
            if hasattr(self, 'cancel_thread') and self.cancel_thread:
                try:
                    if self.cancel_thread.isRunning():
                        self.cancel_thread.quit()
                        if not self.cancel_thread.wait(3000):  # 最大3秒待機（TelephoneTeikyou-crossと同じ）
                            logging.warning("スレッドが正常に終了しませんでした。強制終了します")
                            self.cancel_thread.terminate()
                        else:
                            logging.debug("スレッドを正常に終了しました")
                    else:
                        logging.debug("スレッドは既に停止しています")
                except Exception as e:
                    logging.error(f"スレッド終了中にエラー: {str(e)}")
                    
            # 参照をクリア
            self.cancel_worker = None
            self.cancel_thread = None
            
        except Exception as e:
            logging.error(f"スレッドクリーンアップ中にエラー: {str(e)}")
            # エラー時でも参照をクリア
            self.cancel_worker = None
            self.cancel_thread = None
    
    def fallback_cancel(self):
        """
        フォールバック用の同期キャンセル処理
        """
        try:
            logging.info("フォールバックキャンセル処理を実行")
            if hasattr(self, 'worker') and self.worker:
                self.worker.cancel()
            
            # キャンセルフラグの設定
            try:
                from services.area_search import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
            except ImportError:
                pass
                
            try:
                from services.area_search_east import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
            except ImportError:
                pass
            
            # メインスレッドでタイマーを開始
            QTimer.singleShot(500, self.on_cancel_completed)
        except Exception as e:
            logging.error(f"フォールバックキャンセル処理中にエラー: {str(e)}")
            # エラー時も強制的にキャンセル完了処理を実行
            try:
                self.on_cancel_completed()
            except Exception as complete_error:
                logging.error(f"強制キャンセル完了処理中にもエラー: {str(complete_error)}")
    
    def update_cancel_progress(self, message):
        """
        キャンセル進捗の更新
        
        Args:
            message (str): 進捗メッセージ
        """
        try:
            logging.info(f"キャンセル進捗: {message}")
            # プログレスバーがある場合は更新
            if hasattr(self, 'progress_bar') and self.progress_bar:
                self.progress_bar.setFormat(message)
        except Exception as e:
            logging.error(f"キャンセル進捗更新中にエラー: {str(e)}")
    
    @Slot()
    def check_cancel_timeout(self):
        """
        キャンセルタイムアウトをチェックし、必要に応じて強制リセット
        TelephoneTeikyou-crossの高速方式を適用
        """
        try:
            logging.info("キャンセルタイムアウトチェックを実行")
            
            # まだキャンセル処理中の場合は強制リセット
            if hasattr(self, 'area_search_btn') and self.area_search_btn.text() == "キャンセル中...":
                logging.warning("キャンセル処理が5秒以上継続：強制リセットを実行")
                self.reset_search_button()
                
                # キャンセルスレッドの強制クリーンアップ
                self.cleanup_cancel_thread()
                
                # 結果ラベルを更新
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: キャンセル完了（タイムアウト）")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #E74C3C;
                            border-radius: 4px;
                            background-color: #FADBD8;
                            color: #E74C3C;
                        }
                    """)
            else:
                logging.info("キャンセル処理は正常に完了済みです")
                
        except Exception as e:
            logging.error(f"キャンセルタイムアウトチェック中にエラー: {str(e)}")
            # エラー時も強制リセット
            try:
                self.reset_search_button()
            except Exception as reset_error:
                logging.error(f"強制リセット中にもエラー: {str(reset_error)}")

    def on_cancel_completed(self):
        """
        キャンセル処理完了時の処理
        TelephoneTeikyou-cross完全準拠の高速実装
        """
        try:
            logging.info("★★★ キャンセル処理が完了しました ★★★")
            
            # ★★★ キャンセル完了時にフラグを明示的にクリア ★★★
            try:
                from services.area_search import set_cancel_flag, clear_cancel_flag
                clear_cancel_flag()
                logging.info("キャンセル完了：西日本エリア検索のキャンセルフラグをクリアしました")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import set_cancel_flag, clear_cancel_flag
                clear_cancel_flag()
                logging.info("キャンセル完了：東日本エリア検索のキャンセルフラグをクリアしました")
            except ImportError:
                pass
            
            # キャンセルスレッドのクリーンアップ
            self.cleanup_cancel_thread()
            
            # UI状態をリセット
            self.reset_search_button()
            
            logging.info("キャンセル完了処理が終了しました")
            
        except Exception as e:
            logging.error(f"キャンセル完了処理中にエラー: {str(e)}")
            # エラー時も強制的にUI状態をリセット
            try:
                self.reset_search_button()
            except Exception as reset_error:
                logging.error(f"強制UI状態リセット中にもエラー: {str(reset_error)}")
    
    def on_cancel_timeout(self):
        """
        キャンセル処理がタイムアウトした場合の処理
        """
        try:
            logging.warning("★★★ キャンセル処理がタイムアウトしました - 強制リセットを実行します ★★★")
            
            # 強制的にスレッドとワーカーをクリーンアップ
            try:
                self.cleanup_thread()
            except Exception as cleanup_error:
                logging.error(f"強制クリーンアップ中にエラー: {str(cleanup_error)}")
            
            # キャンセルワーカーとスレッドも強制終了
            if hasattr(self, 'cancel_worker') and self.cancel_worker is not None:
                try:
                    self.cancel_worker.deleteLater()
                except:
                    pass
                self.cancel_worker = None
                
            if hasattr(self, 'cancel_thread') and self.cancel_thread is not None:
                try:
                    from PySide6.QtCore import QThread
                    if isinstance(self.cancel_thread, QThread):
                        if self.cancel_thread.isRunning():
                            logging.info("キャンセルスレッドを強制終了します")
                            self.cancel_thread.terminate()
                            self.cancel_thread.wait(1000)
                        self.cancel_thread.deleteLater()
                    else:
                        logging.warning(f"cancel_threadが不正な型です: {type(self.cancel_thread)}")
                except Exception as e:
                    logging.error(f"キャンセルスレッド終了中にエラー: {str(e)}")
                self.cancel_thread = None
            
            # 強制的にボタンを元に戻す
            self.reset_search_button()
            
            # タイマーをクリーンアップ
            if self.cancel_timer:
                try:
                    self.cancel_timer.deleteLater()
                except:
                    pass
                self.cancel_timer = None
            
            # 結果ラベルを更新
            if hasattr(self, 'area_result_label'):
                self.area_result_label.setText("提供エリア: キャンセル（タイムアウト）")
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
            
            logging.info("★★★ 強制リセットが完了しました ★★★")
            
        except Exception as e:
            logging.error(f"キャンセルタイムアウト処理中にエラー: {str(e)}")
            # 最後の手段として基本的なリセットのみ実行
            try:
                if hasattr(self, 'area_search_btn'):
                    self.area_search_btn.setText("提供エリア検索")
                    self.area_search_btn.setEnabled(True)
            except:
                pass

    def reset_search_button(self):
        """検索ボタンを初期状態に戻す"""
        # ★★★ ボタンリセット時にもキャンセルフラグをクリア ★★★
        try:
            from services.area_search import set_cancel_flag, clear_cancel_flag
            clear_cancel_flag()
            logging.info("ボタンリセット：西日本エリア検索のキャンセルフラグをクリアしました")
        except ImportError:
            pass
        
        try:
            from services.area_search_east import set_cancel_flag, clear_cancel_flag
            clear_cancel_flag()
            logging.info("ボタンリセット：東日本エリア検索のキャンセルフラグをクリアしました")
        except ImportError:
            pass
        
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
        try:
            self.area_search_btn.clicked.disconnect()
        except:
            pass  # 既に接続されていない場合のエラーを無視
        self.area_search_btn.clicked.connect(self.search_service_area)
        
        # リロード/再起動ボタンを非表示
        if hasattr(self, 'reload_btn'):
            self.reload_btn.hide()
        if hasattr(self, 'restart_btn'):
            self.restart_btn.hide()

    def on_search_completed(self, result):
        """検索完了時の処理"""
        try:
            # 検索完了時に進捗バーを100%にする
            self.update_search_progress("検索完了 (100%)")
            
            # ★★★ 検索完了時にキャンセルフラグを必ずクリア ★★★
            try:
                from services.area_search import set_cancel_flag, clear_cancel_flag
                clear_cancel_flag()
                logging.info("★★★ 検索完了：西日本エリア検索のキャンセルフラグをクリアしました ★★★")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import set_cancel_flag, clear_cancel_flag
                clear_cancel_flag()
                logging.info("★★★ 検索完了：東日本エリア検索のキャンセルフラグをクリアしました ★★★")
            except ImportError:
                pass
            
            # 検索スレッドをクリーンアップ
            self.cleanup_thread()
            self.reset_search_button()
            
            # 自動処理フラグをリセット
            self.is_auto_processing = False
            
            # CTI監視システムにも検索完了を通知（処理フラグリセット用）
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                try:
                    # CTI監視システムの処理フラグもリセット
                    with self.cti_status_monitor.processing_lock:
                        if self.cti_status_monitor.is_processing:
                            self.cti_status_monitor.is_processing = False
                            self.cti_status_monitor.talking_start_time = 0
                            logging.info("検索完了によりCTI監視システムの処理フラグもリセットしました")
                        else:
                            logging.debug("CTI監視システムの処理フラグは既にリセット済みです")
                except Exception as cti_error:
                    logging.error(f"CTI監視システムの処理フラグリセット中にエラー: {str(cti_error)}")
            
            # プログレスバーを非表示
            if hasattr(self, 'progress_bar'):
                self.progress_bar.setVisible(False)
            
            if result["status"] == "cancelled":
                # キャンセルされた場合
                logging.info("提供エリア検索がキャンセルされました")
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: キャンセルされました")
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
                return
            
            # 結果表示
            if result["status"] == "available":
                # 提供可能
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: 〇提供可能")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #28A745;
                            border-radius: 4px;
                            background-color: #D4EDDA;
                            color: #155724;
                        }
                    """)
                logging.info("提供判定結果: 提供可能")
                
            elif result["status"] == "unavailable":
                # 提供不可
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: ×提供不可")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #DC3545;
                            border-radius: 4px;
                            background-color: #F8D7DA;
                            color: #721C24;
                        }
                    """)
                logging.info("提供判定結果: 提供不可")
                
            elif result["status"] == "error":
                # エラー
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: エラーが発生しました")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #FFC107;
                            border-radius: 4px;
                            background-color: #FFF3CD;
                            color: #856404;
                        }
                    """)
                logging.error(f"提供判定結果: エラー - {result.get('message', '不明なエラー')}")
                QMessageBox.critical(self, "エラー", result.get("message", "不明なエラーが発生しました"))
                
            else:
                # その他・不明
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: 判定不能")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #6C757D;
                            border-radius: 4px;
                            background-color: #E2E3E5;
                            color: #383D41;
                        }
                    """)
                logging.warning(f"提供判定結果: 不明な状態 - {result}")
                
            # 詳細情報のポップアップ表示
            details = result.get("details", {})
            if result.get("show_popup", False) and details:
                details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
                QMessageBox.information(self, "提供判定結果", details_text)
                
            # スクリーンショットの保存（自動表示はしない）
            if "screenshot" in result and os.path.exists(result["screenshot"]):
                self.screenshot_path = result["screenshot"]
                logging.info(f"スクリーンショットを保存しました: {result['screenshot']}")
                    
        except Exception as e:
            logging.error(f"検索完了処理中にエラーが発生: {str(e)}")
            # エラー時も確実にフラグをリセット
            self.is_auto_processing = False
            self.cleanup_thread()
            self.reset_search_button()
                
            # エラー時もCTI監視システムのフラグをリセット
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                try:
                    with self.cti_status_monitor.processing_lock:
                        self.cti_status_monitor.is_processing = False
                        self.cti_status_monitor.talking_start_time = 0
                        logging.info("エラー時にCTI監視システムの処理フラグもリセットしました")
                except:
                    pass

    def cleanup_thread(self):
        """スレッドとワーカーをクリーンアップ"""
        try:
            # 進捗バーをリセット
            if hasattr(self, 'progress_bar'):
                self.progress_bar.setValue(0)
                self.progress_bar.setVisible(False)
            
            # ワーカーをキャンセル
            if hasattr(self, 'worker') and self.worker:
                try:
                    self.worker.cancel()
                    logging.debug("ワーカーをキャンセルしました")
                except Exception as e:
                    logging.error(f"ワーカーキャンセル中にエラー: {str(e)}")
                
            # スレッドを終了
            if hasattr(self, 'thread') and self.thread:
                try:
                    # QThreadオブジェクトかどうかを確認
                    from PySide6.QtCore import QThread
                    if isinstance(self.thread, QThread):
                        # 正しいQThreadオブジェクトの場合
                        if self.thread.isRunning():
                            self.thread.quit()
                            if not self.thread.wait(2000):  # 最大2秒待機
                                logging.warning("スレッドが正常に終了しませんでした - 強制終了します")
                                self.thread.terminate()
                                self.thread.wait(1000)  # 強制終了後の待機
                            else:
                                logging.debug("スレッドを正常に終了しました")
                    else:
                        logging.warning(f"スレッドオブジェクトが不正な型です: {type(self.thread)}")
                        # 不正な型の場合は参照のみクリア
                except Exception as e:
                    logging.error(f"スレッド終了処理中にエラー: {str(e)}")
                    # エラー時は参照のみクリア（強制操作は行わない）
                    
            # 参照をクリア
            self.worker = None
            self.thread = None
            
        except Exception as e:
            logging.error(f"スレッドクリーンアップ中にエラー: {str(e)}")
            # エラー時でも参照をクリア
            self.worker = None
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
        生年月日から年齢を計算し、80歳以上の場合は年齢をわかりやすく表示する
        （背景色の赤ハイライトは行わない）
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
            # 西暦の場合はそのまま
            
            # 年齢を計算
            age = current_year - year
            
            # 誕生日がまだ来ていない場合は年齢を1つ減らす
            if (month > current_month) or (month == current_month and day > current_day):
                age -= 1
            
            # 80歳以上かどうかをチェック
            is_over_80 = age >= 80

            # 以前の背景スタイルを必ずクリア（赤が残り続けないように）
            self.era_combo.setStyleSheet("")
            self.year_combo.setStyleSheet("")
            self.month_combo.setStyleSheet("")
            self.day_combo.setStyleSheet("")

            # 年齢情報ラベルの表示/非表示とスタイル
            if is_over_80:
                # 表示テキストとスタイル
                self.age_info_label.setText(f"年齢: {age}歳（80歳以上）")
                self.age_info_label.setStyleSheet("""
                    QLabel {
                        color: #C62828;  /* 赤系で注意喚起 */
                        font-weight: bold;
                        font-size: 12px;
                        padding-top: 2px;
                    }
                """)
                self.age_info_label.show()
                logging.info(f"80歳以上の顧客が検出されました: {age}歳")
            else:
                # 80歳未満は非表示（もしくは必要なら年齢を淡色表示に変更可能）
                self.age_info_label.hide()
            
        except Exception as e:
            logging.error(f"年齢チェック中にエラー: {e}")

    def on_mobile_pattern_changed(self, text):
        """
        携帯番号パターンが変更された時の処理
        
        Args:
            text (str): 選択されたテキスト
        """
        if text == "①携帯ありで番号がわかる":
            # 携帯番号入力欄と時間帯入力欄を表示
            self.mobile_number_widget.show()
            self.time_preference_widget.show()
            # 入力欄をクリア
            self.mobile_part1_input.clear()
            self.mobile_part2_input.clear()
            self.mobile_part3_input.clear()
            self.time_preference_input.clear()
            # フォーカスを最初の入力欄に設定
            self.mobile_part1_input.setFocus()
        else:
            # 携帯番号入力欄を非表示、時間帯入力欄は表示
            self.mobile_number_widget.hide()
            self.time_preference_widget.show()
            # 時間帯入力欄をクリア
            self.time_preference_input.clear()
            # パターンに応じてavailable_time_inputを更新
            if text == "②携帯なし":
                self.available_time_input.setText("携帯なし")
            elif text == "③携帯ありで番号がわからない":
                self.available_time_input.setText("携帯不明")
        

    def format_mobile_number_part(self):
        """
        携帯番号の各部分が変更された時の処理
        数字のみを許可し、自動的に次の入力欄にフォーカスを移動
        """
        sender = self.sender()
        if not sender:
            return
            
        # 数字以外の文字を削除
        text = sender.text()
        formatted_text = ''.join(filter(str.isdigit, text))
        
        # 全角数字を半角に変換
        formatted_text = formatted_text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
        
        if formatted_text != text:
            sender.setText(formatted_text)
        
        # 自動フォーカス移動
        if sender == self.mobile_part1_input and len(formatted_text) == 3:
            self.mobile_part2_input.setFocus()
        elif sender == self.mobile_part2_input and len(formatted_text) == 4:
            self.mobile_part3_input.setFocus()
        
        # 携帯番号が完成したらavailable_time_inputを更新
        self.update_available_time_from_mobile_parts()
    
    def update_available_time_from_mobile_parts(self):
        """
        携帯番号の各部分と時間帯から完全な情報を組み立ててavailable_time_inputを更新
        """
        part1 = self.mobile_part1_input.text().strip()
        part2 = self.mobile_part2_input.text().strip()
        part3 = self.mobile_part3_input.text().strip()
        time_pref = self.time_preference_input.text().strip()
        
        if self.mobile_pattern_combo.currentText() == "①携帯ありで番号がわかる":
            if part1 and part2 and part3:
                # 3つの部分がすべて入力されている場合
                mobile_number = f"{part1}-{part2}-{part3}"
                if time_pref:
                    self.available_time_input.setText(f"{time_pref}\n{mobile_number}")
                else:
                    self.available_time_input.setText(mobile_number)
            elif part1 or part2 or part3:
                # 一部だけ入力されている場合は空にする
                self.available_time_input.setText("")
        else:
            # 携帯なしまたは携帯不明の場合
            if time_pref:
                if self.mobile_pattern_combo.currentText() == "②携帯なし":
                    self.available_time_input.setText(f"{time_pref}\n携帯なし")
                else:  # 携帯ありで番号がわからない
                    self.available_time_input.setText(f"{time_pref}\n携帯不明")
            else:
                if self.mobile_pattern_combo.currentText() == "②携帯なし":
                    self.available_time_input.setText("携帯なし")
                else:  # 携帯ありで番号がわからない
                    self.available_time_input.setText("携帯不明")

    def on_cti_dialing_to_talking(self):
        """
        CTI状態が「発信中」→「通話中」に変化した時の自動処理
        
        1. 顧客情報を自動取得
        2. 提供判定検索を自動実行
        """
        try:
            import time
            current_time = time.time()
            
            # 重複実行防止チェック
            if hasattr(self, 'is_auto_processing') and self.is_auto_processing:
                logging.info("CTI自動処理が既に実行中のため、重複実行をスキップします")
                return
                
            # 前回実行から短時間の場合はスキップ
            if hasattr(self, 'last_auto_processing_time'):
                time_since_last = current_time - self.last_auto_processing_time
                if time_since_last < 3.0:  # 3秒以内の重複実行を防ぐ
                    logging.info(f"前回の自動処理から{time_since_last:.2f}秒しか経過していないため、重複実行をスキップします")
                    return
            
            # 処理中フラグを設定
            self.is_auto_processing = True
            self.last_auto_processing_time = current_time
            
            logging.info("CTI状態変化による自動処理を開始します")
            
            # 1. 顧客情報取得を実行（既存のfetch_cti_dataメソッドを呼び出し）
            logging.info("1. 顧客情報の自動取得を開始")
            self.fetch_cti_data()
            
            # 2. 顧客情報取得が完了してから提供判定検索を実行
            # シグナルを使用してメインスレッドで実行（スレッドセーフ）
            import threading
            def delayed_trigger():
                try:
                    self.trigger_auto_search.emit()
                    logging.debug("提供判定検索のシグナルを送信しました")
                except Exception as e:
                    logging.error(f"シグナル送信中にエラー: {str(e)}")
                    # エラー時はフラグをリセット
                    self.is_auto_processing = False
                    logging.debug("自動処理フラグをリセットしました")
                    
            timer = threading.Timer(1.0, delayed_trigger)
            timer.daemon = True
            timer.start()
            
        except Exception as e:
            logging.error(f"CTI自動処理中にエラーが発生: {str(e)}")
            # エラー時もフラグをリセット
            if hasattr(self, 'is_auto_processing'):
                self.is_auto_processing = False
    
    @Slot()
    def auto_search_service_area(self):
        """CTI自動処理から呼び出される提供エリア検索"""
        try:
            logging.info("2. 提供判定検索の自動実行を開始")
            
            # 郵便番号と住所の入力チェック
            postal_code = self.postal_code_input.text().strip()
            address = self.address_input.text().strip()
            
            if not postal_code or not address:
                logging.warning("CTI自動処理: 郵便番号または住所が未入力のため検索をスキップします")
                return
            
            # UIの提供エリア検索ボタンが存在し、クリック可能な場合は直接クリック
            if (hasattr(self, 'area_search_btn') and 
                self.area_search_btn.text() == "提供エリア検索" and 
                self.area_search_btn.isEnabled()):
                
                logging.info("- UIの提供エリア検索ボタンを直接クリックします")
                # UIの検索ボタンをプログラム的にクリック
                self.area_search_btn.click()
                logging.info("★★★ CTI自動処理: UIボタンクリックによる提供判定検索を開始しました ★★★")
                
            elif (hasattr(self, 'area_search_btn') and 
                  self.area_search_btn.text() == "キャンセル"):
                
                logging.info("- 既に検索処理中です（キャンセルボタン表示中）")
                logging.info("★★★ CTI自動処理: 既に検索実行中のため、重複実行をスキップします ★★★")
                
            elif (hasattr(self, 'area_search_btn') and 
                  self.area_search_btn.text() == "キャンセル中..."):
                
                logging.info("- 既にキャンセル処理中です")
                logging.info("★★★ CTI自動処理: キャンセル処理中のため、実行をスキップします ★★★")
                
            else:
                # UIボタンが利用できない場合のフォールバック（従来の処理）
                logging.warning("- UIの提供エリア検索ボタンが利用できません。直接検索処理を実行します")
                
                # 手動実行と同じ処理を使用（ボタン状態変更、キャンセル機能を含む）
                self.search_service_area()
                logging.info("★★★ CTI自動処理: 直接処理による提供判定検索を開始しました ★★★")
            
        except Exception as e:
            logging.error(f"自動検索処理中にエラーが発生: {str(e)}")
            if hasattr(self, 'area_search_btn'):
                self.reset_search_button()

    def on_cti_call_ended(self):
        """通話終了時（通話中→待ち受け中）のコールバック処理"""
        try:
            logging.info("★★★ 通話終了を検出: 2秒後に電話ボタン監視を再開します ★★★")
            if hasattr(self, 'phone_monitor'):
                # 2秒後に監視を再開するタイマーを設定
                QTimer.singleShot(2000, self.restart_phone_monitoring)
        except Exception as e:
            logging.error(f"通話終了時の処理でエラーが発生: {str(e)}")

    def restart_phone_monitoring(self):
        """電話ボタン監視を再開"""
        try:
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.start_monitoring()
                logging.info("電話ボタン監視を再開しました")
        except Exception as e:
            logging.error(f"電話ボタン監視の再開でエラーが発生: {str(e)}")

    def on_cti_talking_started(self):
        """通話中状態開始時のコールバック処理"""
        try:
            logging.info("★★★ 通話中状態を検出: 電話ボタン監視を停止します ★★★")
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.stop_monitoring()
        except Exception as e:
            logging.error(f"通話中状態開始時の処理でエラーが発生: {str(e)}")

    def apply_cti_settings(self):
        """CTI監視設定を適用する"""
        try:
            cti_settings = self.settings.get('cti_settings', {})
            
            # CTI監視の有効/無効を設定
            if cti_settings.get('enable_cti', True):
                if not self.cti_status_monitor.is_monitoring:
                    self.cti_status_monitor.start_monitoring()
                    logging.info("CTI監視を開始しました")
            else:
                if self.cti_status_monitor.is_monitoring:
                    self.cti_status_monitor.stop_monitoring()
                    logging.info("CTI監視を停止しました")
            
            # 自動処理の有効/無効を設定
            self.cti_status_monitor.enable_auto_processing = cti_settings.get('enable_auto_cti_processing', True)
            
            # 監視間隔を設定
            self.cti_status_monitor.monitor_interval = cti_settings.get('cti_monitor_interval', 0.2)
            
            # 通話時間の閾値を設定
            self.cti_status_monitor.call_duration_threshold = cti_settings.get('call_duration_threshold', 0)
            
            logging.info("CTI監視設定を適用しました")
            
        except Exception as e:
            logging.error(f"CTI監視設定の適用中にエラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"CTI監視設定の適用中にエラーが発生しました: {str(e)}")

    def init_cti_monitoring(self):
        """CTI監視機能を初期化"""
        try:
            # CTI状態監視を初期化
            self.cti_status_monitor = CTIStatusMonitor(
                on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
                on_call_ended_callback=self.on_cti_call_ended,
                on_talking_started_callback=self.on_cti_talking_started,
                on_cancel_processing_callback=self.on_cancel_processing_request
            )
            
            # CTI監視の有効状態をログ出力
            logging.info(f"CTI監視: {self.cti_status_monitor.enable_auto_processing}")
            
            # フォントサイズを設定
            self.set_font_size(self.font_size)
            
            if self.cti_status_monitor.enable_auto_processing:
                # CTI状態監視を開始
                self.cti_status_monitor.start_monitoring()
            
        except Exception as e:
            logging.error(f"CTI監視機能の初期化中にエラー: {str(e)}")
            # エラーが発生してもアプリケーションを継続

    def on_cancel_processing_request(self, button_name: str):
        """
        アクションボタンクリック時の処理キャンセル要求を処理
        
        Args:
            button_name: クリックされたボタンの名前
        """
        try:
            logging.info(f"★★★ 「{button_name}」ボタンクリックによる処理キャンセル要求を受信 ★★★")
            
            # 1. 即座にキャンセルフラグを設定（最優先）
            try:
                from services.area_search import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
                logging.info(f"- エリア検索のキャンセルフラグを設定しました")
            except ImportError:
                logging.warning("エリア検索モジュールのインポートに失敗しました")
            
            try:
                from services.area_search_east import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
                logging.info("- 東日本エリア検索のキャンセルフラグを設定しました")
            except ImportError:
                pass
            
            # 2. UIボタンの状態をログ出力
            if hasattr(self, 'area_search_btn'):
                current_button_text = self.area_search_btn.text()
                current_button_enabled = self.area_search_btn.isEnabled()
                logging.info(f"- 現在のUIボタン状態: '{current_button_text}' (有効: {current_button_enabled})")
                
                # 3. 既にキャンセル処理中の場合は重複実行を防ぐ
                if current_button_text == "キャンセル中...":
                    logging.info("- 既にキャンセル処理中です。重複実行をスキップします")
                    return
            else:
                logging.warning("- area_search_btnが存在しません")
                current_button_text = "ボタンなし"
            
            # 検索処理が実行中かどうかを判定（ワーカーとスレッドの存在で判定）
            worker_running = hasattr(self, 'worker') and self.worker is not None
            thread_running = (hasattr(self, 'thread') and self.thread is not None and 
                             hasattr(self.thread, 'isRunning') and self.thread.isRunning())
            
            is_search_running = worker_running or thread_running
            
            logging.info(f"- ワーカー実行中: {worker_running}")
            logging.info(f"- スレッド実行中: {thread_running}")
            logging.info(f"- 検索処理実行中: {is_search_running}")
            
            if is_search_running:
                logging.info(f"- 検索処理が実行中です。UIキャンセル処理を実行します")
                
                # 4. 即座にボタン状態を変更して重複実行を防ぐ
                if hasattr(self, 'area_search_btn'):
                    self.area_search_btn.setText("キャンセル中...")
                    self.area_search_btn.setEnabled(False)
                
                # 5. メインスレッドで安全にキャンセル処理を実行
                if hasattr(self, 'cancel_search'):
                    logging.info(f"- cancel_searchメソッドを実行します")
                    
                    # メインスレッドかどうかを確認
                    if QThread.currentThread() == QApplication.instance().thread():
                        # 少し遅延を入れて確実にキャンセル処理を実行
                        QTimer.singleShot(50, self.cancel_search)
                    else:
                        # 別スレッドから呼ばれた場合はQueuedConnectionで実行
                        QMetaObject.invokeMethod(self, "cancel_search", Qt.QueuedConnection)
                    
                    logging.info(f"★★★ 「{button_name}」ボタン: UIキャンセル処理を開始しました ★★★")
                else:
                    # cancel_searchが存在しない場合の直接処理
                    logging.warning("cancel_searchメソッドが存在しません。直接キャンセル処理を実行します")
                    
                    # ワーカーのキャンセル
                    if hasattr(self, 'worker') and self.worker:
                        logging.info(f"- ワーカーをキャンセルします")
                        self.worker.cancel()
                        logging.info(f"- 実行中のワーカーをキャンセルしました")
                    
                    # UI状態を「キャンセル中」に設定
                    if hasattr(self, 'area_search_btn'):
                        self.area_search_btn.setEnabled(False)
                        self.area_search_btn.setText("キャンセル中...")
                        logging.info(f"- 検索ボタンを「キャンセル中」に設定しました")
                    
                    if hasattr(self, 'area_result_label'):
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
                        logging.info(f"- 結果表示を「キャンセル中」に設定しました")
                    
                    logging.info(f"★★★ 「{button_name}」ボタン: 直接キャンセル処理を実行しました ★★★")
                    
            else:
                logging.info(f"- 検索処理が実行されていません")
                logging.info(f"★★★ 「{button_name}」ボタン: キャンセル対象の処理が実行中ではありません ★★★")
            
        except Exception as e:
            logging.error(f"処理キャンセル要求の処理中にエラー: {str(e)}")
            logging.error(f"エラー詳細: {type(e).__name__}: {str(e)}")
            
            # エラー時も基本的なキャンセルを実行
            try:
                # エラー時もキャンセルフラグを設定
                from services.area_search import set_cancel_flag, clear_cancel_flag
                set_cancel_flag()
                logging.info("エラー時にキャンセルフラグを設定しました")
                
                # エラー時も強制的にcancel_search実行を試行
                if hasattr(self, 'cancel_search'):
                    self.cancel_search()
                    logging.info("エラー時に強制UIキャンセル処理を実行しました")
            except Exception as fallback_error:
                logging.error(f"エラー時のフォールバック処理も失敗: {str(fallback_error)}")
                pass

    def reload_application(self):
        """
        アプリケーションをリロード（初期状態に戻す）
        """
        try:
            logging.info("★★★ アプリケーションのリロードを開始します ★★★")
            
            # 確認ダイアログを表示
            reply = QMessageBox.question(
                self, 
                "リロード確認", 
                "アプリケーションを初期状態に戻しますか？\n（入力中のデータは失われます）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                # 実行中の処理をキャンセル
                if hasattr(self, 'cancel_search'):
                    self.cancel_search()
                
                # CTI監視を停止
                if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                    self.cti_status_monitor.stop_monitoring()
                
                # 電話監視を停止
                if hasattr(self, 'phone_monitor') and self.phone_monitor:
                    self.phone_monitor.stop_monitoring()
                
                # 全ての入力フィールドをクリア
                self.clear_all_inputs()
                
                # 提供エリア検索結果をリセット
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: 未検索")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #ddd;
                            border-radius: 4px;
                            background-color: #f8f8f8;
                        }
                    """)
                
                # 検索ボタンをリセット
                if hasattr(self, 'reset_search_button'):
                    self.reset_search_button()
                
                # リロード/再起動ボタンを非表示
                if hasattr(self, 'reload_btn'):
                    self.reload_btn.hide()
                if hasattr(self, 'restart_btn'):
                    self.restart_btn.hide()
                
                # プログレスバーを非表示
                if hasattr(self, 'progress_bar'):
                    self.progress_bar.hide()
                    self.progress_bar.setValue(0)
                
                # CTI監視を再開
                if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                    self.cti_status_monitor.start_monitoring()
                
                # 電話監視を再開
                if hasattr(self, 'phone_monitor') and self.phone_monitor:
                    self.phone_monitor.start_monitoring()
                
                logging.info("★★★ アプリケーションのリロードが完了しました ★★★")
                QMessageBox.information(self, "リロード完了", "アプリケーションが初期状態に戻りました。")
                
        except Exception as e:
            logging.error(f"アプリケーションリロード中にエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"リロード中にエラーが発生しました: {str(e)}")

    def restart_application(self):
        """
        アプリケーションを再起動
        """
        try:
            logging.info("★★★ アプリケーションの再起動を開始します ★★★")
            
            # 確認ダイアログを表示
            reply = QMessageBox.question(
                self, 
                "再起動確認", 
                "アプリケーションを再起動しますか？\n（入力中のデータは失われます）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                # 実行中の処理をキャンセル
                if hasattr(self, 'cancel_search'):
                    self.cancel_search()
                
                # CTI監視を停止
                if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                    self.cti_status_monitor.stop_monitoring()
                
                # 電話監視を停止
                if hasattr(self, 'phone_monitor') and self.phone_monitor:
                    self.phone_monitor.stop_monitoring()
                
                logging.info("アプリケーションを再起動します...")
                
                # 実行ファイルのパスを正しく取得
                import sys
                import os
                import subprocess
                
                if getattr(sys, 'frozen', False):
                    # PyInstallerで作成されたexeファイルの場合
                    executable_path = sys.executable
                    logging.info(f"PyInstaller環境での再起動: {executable_path}")
                    
                    # 新しいプロセスを起動（引数なしで起動）
                    try:
                        subprocess.Popen([executable_path], 
                                       cwd=os.path.dirname(executable_path),
                                       creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == 'nt' else 0)
                        logging.info("新しいプロセスを起動しました")
                    except Exception as start_error:
                        logging.error(f"新しいプロセス起動エラー: {start_error}")
                        QMessageBox.critical(self, "エラー", f"新しいプロセスの起動に失敗しました: {str(start_error)}")
                        return
                else:
                    # 通常のPythonスクリプトとして実行されている場合
                    python_executable = sys.executable
                    script_path = sys.argv[0]
                    
                    # 新しいプロセスでアプリケーションを起動
                    if os.name == 'nt':  # Windows
                        subprocess.Popen([python_executable, script_path])
                    else:  # Unix/Linux/Mac
                        os.execv(python_executable, [python_executable] + sys.argv)
                    
                    logging.info("通常のPython環境での再起動")
                
                # 現在のアプリケーションを終了
                logging.info("現在のプロセスを終了します")
                QApplication.quit()
                
        except Exception as e:
            logging.error(f"アプリケーション再起動中にエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"再起動中にエラーが発生しました: {str(e)}")

    def check_cancel_timeout(self):
        """
        キャンセル処理のタイムアウトをチェックし、必要に応じて強制リセット
        """
        try:
            # 現在のボタン状態をチェック
            if (hasattr(self, 'area_search_btn') and 
                self.area_search_btn.text() == "キャンセル中..." and 
                not self.area_search_btn.isEnabled()):
                
                logging.warning("★★★ キャンセル処理がタイムアウトしました - 強制リセットを実行します ★★★")
                
                # 強制的にリセット
                self.reset_search_button()
                
                # 結果表示を更新
                if hasattr(self, 'area_result_label'):
                    self.area_result_label.setText("提供エリア: キャンセルタイムアウト（強制リセット済み）")
                    self.area_result_label.setStyleSheet("""
                        QLabel {
                            font-size: 14px;
                            padding: 5px;
                            border: 1px solid #DC3545;
                            border-radius: 4px;
                            background-color: #F8D7DA;
                            color: #721C24;
                        }
                    """)
                
                # プログレスバーを非表示
                if hasattr(self, 'progress_bar'):
                    self.progress_bar.hide()
                    self.progress_bar.setValue(0)
                
                # ワーカーとスレッドを強制終了
                if hasattr(self, 'worker'):
                    self.worker = None
                if hasattr(self, 'thread') and self.thread:
                    if self.thread.isRunning():
                        self.thread.quit()
                        self.thread.wait(1000)  # 1秒待機
                    self.thread = None
                
                # CTI監視システムの処理フラグもリセット
                if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                    try:
                        with self.cti_status_monitor.processing_lock:
                            self.cti_status_monitor.is_processing = False
                            self.cti_status_monitor.talking_start_time = 0
                            logging.info("タイムアウト時にCTI監視システムの処理フラグもリセットしました")
                    except Exception as cti_error:
                        logging.error(f"CTI監視システムのリセット中にエラー: {str(cti_error)}")
                
                # 自動処理フラグもリセット
                if hasattr(self, 'is_auto_processing'):
                    self.is_auto_processing = False
                
                logging.info("キャンセルタイムアウトによる強制リセットが完了しました")
                
        except Exception as e:
            logging.error(f"キャンセルタイムアウトチェック中にエラー: {str(e)}")


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
        
        # area_searchモジュールのキャンセルフラグを設定
        try:
            from services.area_search import set_cancel_flag, clear_cancel_flag
            set_cancel_flag()
            logging.info("ServiceAreaSearchWorker: area_searchのキャンセルフラグを設定しました")
        except ImportError as e:
            logging.warning(f"ServiceAreaSearchWorker: area_searchのキャンセルフラグ設定に失敗: {e}")
        
        # area_search_eastモジュールのキャンセルフラグを設定
        try:
            from services.area_search_east import set_cancel_flag, clear_cancel_flag
            set_cancel_flag()
            logging.info("ServiceAreaSearchWorker: area_search_eastのキャンセルフラグを設定しました")
        except ImportError as e:
            logging.warning(f"ServiceAreaSearchWorker: area_search_eastのキャンセルフラグ設定に失敗: {e}")
    
    def _update_progress(self, message=None):
        """
        進捗状況を更新する
        
        Args:
            message (str, optional): カスタムメッセージ。指定がない場合は定義済みメッセージを使用
        """
        try:
            # メッセージに既にパーセンテージが含まれている場合はそのまま使用
            if message and "%" in message:
                self.progress.emit(message)
                return
            
            # メッセージごとの進捗マッピング
            progress_map = {
                "住所情報を解析中...": 5,
                "ブラウザを起動中...": 15,
                "NTT西日本サイトにアクセス中...": 25,
                "郵便番号を入力中...": 35,
                "基本住所の候補を検索中...": 45,
                "番地を入力中...": 60,
                "号を入力中...": 75,
                "検索結果を確認中...": 85,
                "集合住宅と判定しました。スクリーンショットを保存します。": 90
            }
            
            if message in progress_map:
                # 特定のメッセージの場合は対応する進捗率を使用
                progress_percent = progress_map[message]
                message = f"{message} ({progress_percent}%)"
            elif message and "%" not in message:
                # その他のメッセージの場合は現在のステップベース
                progress_percent = min(int((self._accumulated_progress / self._total_weight) * 90), 90)
                message = f"{message} ({progress_percent}%)"
            elif message is None and self._current_step < len(self._progress_steps):
                # 定義済みステップの場合
                step_info = self._progress_steps[self._current_step]
                message = step_info["message"]
                self._accumulated_progress += step_info["weight"]
                progress_percent = min(int((self._accumulated_progress / self._total_weight) * 90), 90)
                message = f"{message} ({progress_percent}%)"
                self._current_step += 1
            else:
                message = "処理中... (0%)"
            
            self.progress.emit(message)
            
        except Exception as e:
            logging.error(f"進捗更新中にエラー: {e}")
            self.progress.emit(f"処理中... (エラー)")
    
    def run(self):
        """提供エリア検索を実行し、結果をシグナルで通知する"""
        try:
            # 開始前にキャンセルフラグをチェック
            try:
                from services.area_search import is_cancelled
                if is_cancelled():
                    logging.info("検索開始前にキャンセルが検出されました")
                    raise CancellationError("検索がキャンセルされました")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import is_cancelled
                if is_cancelled():
                    logging.info("検索開始前にキャンセルが検出されました（東日本）")
                    raise CancellationError("検索がキャンセルされました")
            except ImportError:
                pass
            
            # ワーカー自体のキャンセルフラグもチェック
            if self._is_cancelled:
                logging.info("ワーカーレベルでキャンセルが検出されました")
                raise CancellationError("検索がキャンセルされました")
            
            # キャンセルフラグを初期化（まだキャンセルされていない場合のみ）
            try:
                from services.area_search import is_cancelled, clear_cancel_flag
                if not is_cancelled():
                    clear_cancel_flag()
            except ImportError:
                pass
            
            try:
                from services.area_search_east import is_cancelled, clear_cancel_flag
                if not is_cancelled():
                    clear_cancel_flag()
            except ImportError:
                pass
                
            # 進捗状況を通知するコールバック関数を定義
            def progress_callback(message):
                if self._is_cancelled:
                    raise CancellationError("検索がキャンセルされました")
                self._update_progress(message)

            # 検索を実行
            result = search_service_area(
                self.postal_code,
                self.address,
                progress_callback=progress_callback
            )
            
            # 結果処理前の最終キャンセルチェック（競合状態回避）
            if self._is_cancelled:
                logging.info("検索結果処理前にキャンセルが検出されました")
                raise CancellationError("検索がキャンセルされました")
            
            # 外部キャンセルフラグもチェック
            try:
                from services.area_search import is_cancelled
                if is_cancelled():
                    logging.info("検索結果処理前に外部キャンセルが検出されました")
                    raise CancellationError("検索がキャンセルされました")
            except ImportError:
                pass
            
            try:
                from services.area_search_east import is_cancelled
                if is_cancelled():
                    logging.info("検索結果処理前に東日本キャンセルが検出されました")
                    raise CancellationError("検索がキャンセルされました")
            except ImportError:
                pass
            
            logging.info(f"★★★ 検索結果を返却します: {result.get('status', 'unknown')} ★★★")
            
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
        """設定ファイルを読み込む"""
        try:
            # 初期設定を設定
            self.settings = {
                'font_size': 11,
                'cti_settings': {
                    'enable_cti': True,
                    'enable_auto_cti_processing': True,
                    'cti_monitor_interval': 0.2,
                    'call_duration_threshold': 0  # 通話時間の閾値（デフォルトは0秒）
                },
                'browser_settings': {
                    'headless': True,
                    'disable_images': True,
                    'show_popup': True,
                    'auto_close': True,
                    'page_load_timeout': 60,
                    'script_timeout': 60
                }
            }
            
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    # 既存の設定を更新
                    self.settings.update(loaded_settings)
            else:
                # デフォルト設定をファイルに保存
                self.save_settings()
            
            logging.info("設定ファイルを読み込みました")
            logging.info(f"設定内容: {self.settings}")
            
        except Exception as e:
            logging.error(f"設定の読み込み中にエラーが発生しました: {e}", exc_info=True)
            self.settings = {}
            QMessageBox.critical(self, "エラー", f"設定の読み込み中にエラーが発生しました: {e}")

