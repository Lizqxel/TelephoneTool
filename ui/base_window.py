"""
ベースウィンドウモジュール

このモジュールは、アプリケーションのウィンドウの基本機能を提供します。
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
                              QSizePolicy, QProgressBar)
from PySide6.QtCore import Qt, QTimer, QPoint, QUrl, QEvent, QObject, Signal, QThread
from PySide6.QtGui import QFont, QIntValidator, QClipboard, QPixmap, QIcon, QDesktopServices

from version import VERSION, GITHUB_OWNER, GITHUB_REPO, APP_NAME
from ui.settings_dialog import SettingsDialog
from services.area_search import search_service_area
from utils.format_utils import (format_phone_number, format_phone_number_without_hyphen,
                               format_postal_code, convert_to_half_width)
from typing import Dict, Any, List, Optional, Union, Tuple
from ui.update_dialog import UpdateDialog

class BaseWindow(QMainWindow):
    """ベースウィンドウクラス"""
    
    def __init__(self):
        """初期化"""
        super().__init__()
        
        # バージョン情報の設定
        self.version = "1.0.0"
        
        # モード変更フラグ（設定ダイアログ用）
        self.mode_changed = False
        self.new_mode = None
        
        # ログ設定
        self.setup_logging()
        
        # 設定ファイルのパスを設定
        self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # 設定を読み込む
        self.settings = {}
        self.load_settings()
        
        # アクティブな検索スレッドを保持するリスト
        self.active_search_threads = []
        
        # ウィンドウの基本設定
        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.setGeometry(100, 100, 800, 600)
        
        # メインウィジェットの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウトの作成
        self.main_layout = QVBoxLayout(main_widget)
        
        # フォントサイズの設定
        font_size = self.settings.get('font_size', 10)
        self.set_font_size(font_size)
    
    def set_font_size(self, size):
        """フォントサイズを設定する"""
        try:
            font = QFont()
            font.setPointSize(size)
            self.setFont(font)
            if hasattr(self, 'preview_text'):
                self.preview_text.setFont(font)
            logging.info(f"フォントサイズを {size} に設定しました")
        except Exception as e:
            logging.error(f"フォントサイズの設定中にエラーが発生しました: {e}")
    
    def setup_logging(self):
        """ログ設定を行う"""
        try:
            log_dir = "logs"
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = os.path.join(log_dir, f"app_{timestamp}.log")
            
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
    
    def load_settings(self):
        """設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                    logging.info("設定ファイルを読み込みました")
            else:
                self.settings = {}
                logging.warning("設定ファイルが存在しません")
            
            if 'format_template' not in self.settings or not self.settings['format_template']:
                logging.error("フォーマットテンプレートが設定されていません")
                QMessageBox.warning(self, "警告", "フォーマットテンプレートが設定されていません。\n設定画面でテンプレートを設定してください。")
                return
            
            logging.info(f"フォーマットテンプレート: {self.settings['format_template']}")
            
            font_size = self.settings.get('font_size', 10)
            logging.info(f"フォントサイズを {font_size} に設定しました")
            
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.update_settings()
            
        except Exception as e:
            logging.error(f"設定の読み込み中にエラーが発生しました: {e}", exc_info=True)
            self.settings = {}
            QMessageBox.critical(self, "エラー", f"設定の読み込み中にエラーが発生しました: {e}")
    
    def closeEvent(self, event):
        """ウィンドウを閉じる際の処理"""
        try:
            if hasattr(self, 'active_search_threads'):
                for thread in self.active_search_threads:
                    if thread and thread.isRunning():
                        logging.info("アクティブな検索スレッドを停止します")
                        thread.stop()
                self.active_search_threads.clear()
            
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.stop_monitoring()
                
            event.accept()
        except Exception as e:
            logging.error(f"アプリケーション終了処理中にエラー: {e}")
            event.accept()
    
    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            self.load_settings()
            self.apply_font_size()
            
            if hasattr(self, 'mode_changed') and self.mode_changed and hasattr(self, 'new_mode'):
                self.mode_changed = False
                new_mode = self.new_mode
                self.new_mode = None
                
                if new_mode == 'simple':
                    self.init_simple_mode()
                else:
                    self.init_easy_mode()
            else:
                self.update()
                for widget in self.findChildren(QWidget):
                    widget.update()
            
            logging.info("設定を更新しました")
    
    def init_menu(self):
        """メニューバーの初期化"""
        menubar = self.menuBar()
        menubar.clear()
        
        file_menu = menubar.addMenu("ファイル")
        exit_action = file_menu.addAction("終了")
        exit_action.triggered.connect(self.close)
        
        help_menu = menubar.addMenu("ヘルプ")
        update_action = help_menu.addAction("アップデートの確認")
        update_action.triggered.connect(self.show_update_dialog)
        
        about_action = help_menu.addAction("バージョン情報")
        about_action.triggered.connect(self.show_about_dialog)
        
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
        dialog.settings_file = self.settings_file
        dialog.exec()
    
    def show_about_dialog(self):
        """バージョン情報ダイアログを表示する"""
        msg = f"{APP_NAME} v{VERSION}\n\n"
        msg += "ライセンス: MIT License"
        QMessageBox.information(self, "バージョン情報", msg)
    
    def check_for_updates(self):
        """アップデートをチェック"""
        try:
            url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            response = requests.get(url)
            response.raise_for_status()
            latest_release = response.json()
            
            latest_version = latest_release["tag_name"].lstrip("v")
            current_version = VERSION
            
            if latest_version > current_version:
                msg = f"新しいバージョン v{latest_version} が利用可能です。\n"
                msg += f"現在のバージョン: v{current_version}\n\n"
                msg += "更新しますか？"
                
                reply = QMessageBox.question(self, "アップデート", msg,
                                          QMessageBox.StandardButton.Yes |
                                          QMessageBox.StandardButton.No)
                
                if reply == QMessageBox.StandardButton.Yes:
                    dialog = UpdateDialog(self)
                    dialog.settings_file = self.settings_file
                    dialog.download_and_apply_update(latest_release)
        except Exception as e:
            logging.error(f"アップデートチェック中にエラー: {e}") 