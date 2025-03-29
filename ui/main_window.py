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
                              QApplication, QToolTip, QSplitter, QMenuBar, QMenu,
                              QDialog)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices

from ui.settings_dialog import SettingsDialog
from ui.mode_selection_dialog import ModeSelectionDialog
from ui.easy_mode_dialogs import AddressInfoDialog, ListInfoDialog, OrdererInputDialog, OrderInfoDialog, DIALOG_BACK, DIALOG_NEXT, DIALOG_CANCEL
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
        
        # ログ設定
        self.setup_logging()
        
        # 設定ファイルのパスを設定
        self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # 設定を読み込む
        self.settings = {}
        self.load_settings()
        
        # モード設定
        self.current_mode = self.settings.get('mode', 'simple')
        logging.info(f"現在のモード: {self.current_mode}")
        
        # ウィンドウの基本設定
        self.setWindowTitle("電話ツール")
        self.setMinimumSize(800, 600)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの作成
        self.main_layout = QVBoxLayout(main_widget)
        
        # モード選択ダイアログの表示（設定に関係なく常に表示）
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
                # 設定ファイルが存在しない場合は、モード選択ダイアログを表示
                self.show_mode_selection_dialog()
        except Exception as e:
            logging.error(f"モード設定の読み込み中にエラーが発生しました: {e}")
            # エラーが発生した場合は、モード選択ダイアログを表示
            self.show_mode_selection_dialog()
    
    def show_mode_selection(self):
        """
        モード選択ダイアログを表示する
        """
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
            mode (str): 選択されたモード
            show_again (bool): 次回以降表示するかどうか
        """
        try:
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            
            # モード設定を更新
            settings['mode'] = mode
            settings['show_mode_selection'] = show_again
            
            # 設定を保存
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            
            logging.info(f"モード設定を保存しました: {mode}")
        except Exception as e:
            logging.error(f"モード設定の保存中にエラーが発生しました: {e}")
            # エラーが発生した場合は、デフォルトでシンプルモードを使用
            self.current_mode = 'simple'
    
    def init_simple_mode(self):
        """シンプルモードのUIを初期化"""
        # 設定に基づいてウィンドウタイトルを設定
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(600, 400)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの設定
        main_layout = QVBoxLayout(main_widget)
        
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
    
    def init_easy_mode(self):
        """使いやすいモードのUIを初期化"""
        # 設定に基づいてウィンドウタイトルを設定
        self.setWindowTitle("コールセンター業務効率化ツール - 使いやすいモード")
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
        """使いやすいモードを開始"""
        try:
            logging.info("使いやすいモードを開始")
            
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
            
            self.list_data = {
                'list_name': customer_name if customer_name else "",
                'list_furigana': convert_to_half_width(getattr(cti_data, 'customer_furigana', '')) if hasattr(cti_data, 'customer_furigana') else "",
                'list_phone': convert_to_half_width(cti_data.phone) if cti_data.phone else "",
                'list_postal_code': convert_to_half_width(cti_data.postal_code) if cti_data.postal_code else "",
                'list_address': address if address else ""
            }
            
            self.orderer_data = {
                'operator': '',  # 対応者名は空で初期化
                'available_time': '',  # 出やすい時間帯は空で初期化
                'contractor': customer_name if customer_name else "",  # 変換済みの顧客名を使用
                'furigana': convert_to_half_width(getattr(cti_data, 'customer_furigana', '')) if hasattr(cti_data, 'customer_furigana') else "",
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
                'order_date': datetime.datetime.now().strftime('%m/%d'),
                'judgment': 'OK'  # デフォルト値を設定
            }
            
            current_dialog = None
            while True:
                if current_dialog is None or isinstance(current_dialog, AddressInfoDialog):
                    # 住所情報ダイアログを表示
                    dialog = AddressInfoDialog(self, self.address_data)
                    result = dialog.exec()
                    
                    # 作成中止が選択された場合
                    if result == DIALOG_CANCEL:
                        logging.info("作成中止が選択されました")
                        self.preview_text.clear()
                        self.statusBar().showMessage("作成中止")
                        return
                    
                    # 住所情報を保存
                    self.address_data = dialog.get_saved_data()
                    current_dialog = dialog
                    
                    if result == DIALOG_NEXT:
                        current_dialog = ListInfoDialog(self, self.list_data)
                
                elif isinstance(current_dialog, ListInfoDialog):
                    # リスト情報ダイアログを表示
                    result = current_dialog.exec()
                    
                    # 作成中止が選択された場合
                    if result == DIALOG_CANCEL:
                        logging.info("作成中止が選択されました")
                        self.preview_text.clear()
                        self.statusBar().showMessage("作成中止")
                        return
                    
                    # 戻るボタンが押された場合
                    if result == DIALOG_BACK:
                        # リスト情報を保存
                        self.list_data = current_dialog.get_saved_data()
                        current_dialog = AddressInfoDialog(self, self.address_data)
                        continue
                    
                    # リスト情報を保存
                    self.list_data = current_dialog.get_saved_data()
                    current_dialog = OrdererInputDialog(self, self.orderer_data)
                
                elif isinstance(current_dialog, OrdererInputDialog):
                    # 受注者入力項目ダイアログを表示
                    result = current_dialog.exec()
                    
                    # 作成中止が選択された場合
                    if result == DIALOG_CANCEL:
                        logging.info("作成中止が選択されました")
                        self.preview_text.clear()
                        self.statusBar().showMessage("作成中止")
                        return
                    
                    # 戻るボタンが押された場合
                    if result == DIALOG_BACK:
                        # 受注者情報を保存
                        self.orderer_data = current_dialog.get_saved_data()
                        current_dialog = ListInfoDialog(self, self.list_data)
                        continue
                    
                    # 受注者情報を保存
                    self.orderer_data = current_dialog.get_saved_data()
                    current_dialog = OrderInfoDialog(self, self.order_data)
                
                elif isinstance(current_dialog, OrderInfoDialog):
                    # 受注情報入力ダイアログを表示
                    result = current_dialog.exec()
                    
                    # 作成中止が選択された場合
                    if result == DIALOG_CANCEL:
                        logging.info("作成中止が選択されました")
                        self.preview_text.clear()
                        self.statusBar().showMessage("作成中止")
                        return
                    
                    # 戻るボタンが押された場合
                    if result == DIALOG_BACK:
                        # 受注情報を保存
                        self.order_data = current_dialog.get_saved_data()
                        current_dialog = OrdererInputDialog(self, self.orderer_data)
                        continue
                    
                    # 受注情報を保存
                    self.order_data = current_dialog.get_saved_data()
                    break
            
        except Exception as e:
            logging.error(f"使いやすいモードの開始中にエラー: {e}", exc_info=True)
            QMessageBox.critical(self, "エラー", f"使いやすいモードの開始中にエラーが発生しました: {e}")

    def show_address_dialog(self):
        """住所情報ダイアログを表示"""
        try:
            dialog = AddressInfoDialog(self, self.address_data)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                # 次のダイアログへ進む前にデータを保存
                self.address_data = dialog.get_saved_data()
                self.show_list_dialog()
            else:
                logging.info("住所情報入力がキャンセルされました")
        except Exception as e:
            logging.error(f"住所情報ダイアログの表示中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"住所情報ダイアログの表示中にエラーが発生しました: {e}")

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
        top_bar.setFixedHeight(32)  # トップバーの高さを32pxに設定
        top_bar.setStyleSheet("""
            QWidget {
                background-color: #2C3E50;
                color: white;
            }
        """)
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(5, 2, 5, 2)  # 上下のマージンを2pxに設定
        top_bar_layout.setSpacing(4)  # ボタン間のスペースを4pxに設定
        
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
        
        for btn in [self.clear_btn, self.cti_copy_btn, 
                   self.screenshot_btn, self.spreadsheet_btn, self.settings_btn]:
            btn.setStyleSheet(button_style)
        
        # ボタンの接続
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        self.cti_copy_btn.clicked.connect(self.copy_cti_to_clipboard)
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        self.spreadsheet_btn.clicked.connect(self.write_to_spreadsheet)
        self.settings_btn.clicked.connect(self.show_settings)
        
        # ボタンをレイアウトに追加（指定された順序で）
        top_bar_layout.addWidget(self.clear_btn)
        top_bar_layout.addWidget(self.cti_copy_btn)
        top_bar_layout.addWidget(self.screenshot_btn)
        top_bar_layout.addWidget(self.spreadsheet_btn)
        top_bar_layout.addWidget(self.settings_btn)
        
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
        
        self.era_combo = CustomComboBox()
        self.era_combo.addItems(["昭和", "平成", "西暦"])
        self.era_combo.setFixedWidth(60)  # 幅を60pxに設定
        birth_input_layout.addWidget(self.era_combo)
        
        self.year_combo = CustomComboBox()
        # 初期値として昭和の年を設定
        self.year_combo.addItems([str(i) for i in range(1, 65)])
        self.year_combo.setEditable(True)
        self.year_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.year_combo.lineEdit().setMaxLength(4)  # 最大4桁
        self.year_combo.lineEdit().setValidator(QIntValidator(1, 9999))  # 1-9999の範囲で制限
        self.year_combo.setFixedWidth(60)  # 幅を60pxに設定
        birth_input_layout.addWidget(self.year_combo)
        birth_input_layout.addWidget(QLabel("年"))
        
        self.month_combo = CustomComboBox()
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setEditable(True)
        self.month_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.month_combo.lineEdit().setMaxLength(2)  # 最大2桁
        self.month_combo.lineEdit().setValidator(QIntValidator(1, 12))  # 1-12の範囲で制限
        self.month_combo.setFixedWidth(40)  # 幅を40pxに設定
        birth_input_layout.addWidget(self.month_combo)
        birth_input_layout.addWidget(QLabel("月"))
        
        self.day_combo = CustomComboBox()
        self.day_combo.addItems([str(i) for i in range(1, 32)])
        self.day_combo.setEditable(True)
        self.day_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.day_combo.lineEdit().setMaxLength(2)  # 最大2桁
        self.day_combo.lineEdit().setValidator(QIntValidator(1, 31))  # 1-31の範囲で制限
        self.day_combo.setFixedWidth(40)  # 幅を40pxに設定
        birth_input_layout.addWidget(self.day_combo)
        birth_input_layout.addWidget(QLabel("日"))
        
        birth_layout.addLayout(birth_input_layout)
        input_layout.addLayout(birth_layout)
        
        # 受注者名
        input_layout.addWidget(QLabel("受注者名"))
        self.order_person_input = QLineEdit()
        input_layout.addWidget(self.order_person_input)
        
        # 社番を追加
        input_layout.addWidget(QLabel("社番"))
        self.employee_number_input = QLineEdit()
        input_layout.addWidget(self.employee_number_input)
        
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
        address_container = QHBoxLayout()
        self.address_input = QLineEdit()
        address_container.addWidget(self.address_input)
        
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
            self.employee_number_input.textChanged.connect(self.reset_background_color)
            self.nd_input.textChanged.connect(self.reset_background_color)
            
            # ボタンのシグナル接続
            self.area_search_btn.clicked.connect(self.search_service_area)
            self.map_btn.clicked.connect(self.open_street_view)
        else:
            # 使いやすいモード用のシグナル設定
            # 開始ボタンのシグナルは既に接続済み
            pass
    
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
        
        self.relationship_input.clear()
        # 社番はクリアしない（保持する）
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
        """メニューの初期化"""
        menubar = self.menuBar()
        
        # ファイルメニュー
        file_menu = menubar.addMenu("ファイル")
        
        exit_action = file_menu.addAction("終了")
        exit_action.triggered.connect(self.close)
        
        # ヘルプメニュー
        help_menu = menubar.addMenu("ヘルプ")
        
        update_action = help_menu.addAction("アップデート設定")
        update_action.triggered.connect(self.show_update_dialog)
        
        about_action = help_menu.addAction("バージョン情報")
        about_action.triggered.connect(self.show_about_dialog)
        
    def show_update_dialog(self):
        """アップデート設定ダイアログを表示する"""
        dialog = UpdateDialog(self)
        dialog.exec()
        
    def show_about_dialog(self):
        """バージョン情報ダイアログを表示する"""
        from version import VERSION, APP_NAME
        
        QMessageBox.about(
            self,
            "バージョン情報",
            f"{APP_NAME} v{VERSION}\n\n"
            "© 2024 Your Company Name\n"
            "All rights reserved."
        )

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

    def generate_preview_text(self):
        """
        プレビューテキストを生成
        
        Returns:
            str: 生成されたプレビューテキスト
        """
        try:
            logging.info("プレビューテキストの生成を開始")
            
            # 設定からフォーマットテンプレートを取得
            template = self.settings.get('format_template', '')
            logging.info(f"フォーマットテンプレート: {template}")
            
            # テンプレートが空の場合はエラー
            if not template:
                logging.error("フォーマットテンプレートが設定されていません")
                QMessageBox.warning(self, "警告", "フォーマットテンプレートが設定されていません。\n設定画面でテンプレートを設定してください。")
                return None
            
            # シンプルモードの場合
            if self.current_mode == 'simple':
                logging.info("シンプルモードでのプレビュー生成")
                # 入力データの取得
                data = {
                    'operator': self.operator_input.text(),
                    'available_time': self.available_time_input.text(),
                    'contractor': self.contractor_input.text(),
                    'furigana': self.furigana_input.text(),
                    'era': self.era_combo.currentText(),
                    'year': self.year_combo.currentText(),
                    'month': self.month_combo.currentText(),
                    'day': self.day_combo.currentText(),
                    'order_person': self.order_person_input.text(),
                    'employee_number': self.employee_number_input.text(),
                    'fee': self.fee_input.text(),
                    'net_usage': self.net_usage_combo.currentText(),
                    'family_approval': self.family_approval_combo.currentText(),
                    'other_number': self.other_number_input.text(),
                    'phone_device': self.phone_device_input.text(),
                    'forbidden_line': self.forbidden_line_input.text(),
                    'nd': self.nd_input.text(),
                    'relationship': self.relationship_input.text(),
                    'postal_code': self.postal_code_input.text(),
                    'address': self.address_input.text(),
                    'list_name': self.list_name_input.text(),
                    'list_furigana': self.list_furigana_input.text(),
                    'list_phone': self.list_phone_input.text(),
                    'list_postal_code': self.list_postal_code_input.text(),
                    'list_address': self.list_address_input.text(),
                    'current_line': self.current_line_combo.currentText(),
                    'order_date': self.order_date_input.text(),
                    'judgment': self.judgment_combo.currentText()
                }
            # 使いやすいモードの場合
            else:
                logging.info("使いやすいモードでのプレビュー生成")
                # 各ダイアログのデータを取得
                address_data = getattr(self, 'address_data', {})
                list_data = getattr(self, 'list_data', {})
                orderer_data = getattr(self, 'orderer_data', {})
                order_data = getattr(self, 'current_dialog', None)
                
                logging.info(f"住所データ: {address_data}")
                logging.info(f"リストデータ: {list_data}")
                logging.info(f"受注者データ: {orderer_data}")
                
                if order_data:
                    order_data = order_data.get_order_data()
                    logging.info(f"受注データ: {order_data}")
                else:
                    logging.warning("受注データが取得できません")
                    order_data = {}
                
                # データを統合
                data = {
                    **address_data,
                    **list_data,
                    **orderer_data,
                    **order_data
                }
                logging.info(f"統合されたデータ: {data}")
                
                # データが空の場合はエラー
                if not data:
                    logging.error("営コメ作成時に必要なデータが取得できません")
                    return None
            
            # テンプレートの置換
            preview_text = template
            for key, value in data.items():
                placeholder = f"{{{key}}}"
                preview_text = preview_text.replace(placeholder, str(value))
                logging.debug(f"プレースホルダー {placeholder} を {value} に置換")
            
            logging.info("プレビューテキストの生成が完了")
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

