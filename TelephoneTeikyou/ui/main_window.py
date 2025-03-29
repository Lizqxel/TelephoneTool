"""
メインウィンドウモジュール

このモジュールは、アプリケーションのメインウィンドウを定義します。
"""

import os
import sys
import json
import logging
import webbrowser
from urllib.parse import quote
from datetime import datetime
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QProgressBar,
    QMessageBox, QComboBox, QCheckBox, QSpacerItem,
    QSizePolicy, QGroupBox, QScrollArea,
    QApplication
)
from PySide6.QtCore import Qt, QThread, QObject, Signal, Slot, QEvent, QTimer
from PySide6.QtGui import QFont, QIcon, QPixmap, QCursor

from services.area_search import search_service_area
from ui.settings_dialog import SettingsDialog
from services.oneclick import OneClickService
from services.phone_button_monitor import PhoneButtonMonitor

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
    
    def run(self):
        """
        提供エリア検索を実行し、結果をシグナルで通知する
        """
        try:
            # 進捗状況を通知するコールバック関数を定義
            def progress_callback(message):
                self.progress.emit(message)

            # 検索を実行
            result = search_service_area(
                self.postal_code,
                self.address,
                progress_callback=progress_callback
            )
            self.finished.emit(result)
        except Exception as e:
            logging.error(f"検索処理中にエラーが発生: {str(e)}")
            self.progress.emit("エラーが発生しました")
            self.finished.emit({
                "status": "error",
                "message": f"検索処理中にエラーが発生: {str(e)}"
            })

class MainWindow(QMainWindow):
    """メインウィンドウクラス"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("コールセンター業務効率化ツール")
        self.setMinimumSize(600, 400)
        
        # スレッドとワーカーの初期化
        self.thread = None
        self.worker = None
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # CTIサービスの初期化
        self.cti_service = OneClickService()
        
        # トップバーを作成し、メインレイアウトに追加
        self.create_top_bar(main_layout)
        
        # 入力フォームエリアをスクロール可能に
        form_widget = QWidget()
        form_layout = QVBoxLayout(form_widget)
        self.create_input_form(form_layout)
        
        # スクロールエリアの作成
        scroll_area = QScrollArea()
        scroll_area.setWidget(form_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # スクロールエリアのスタイルシート
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
        
        # スクロールエリアをメインレイアウトに追加
        main_layout.addWidget(scroll_area)
        
        # 設定の読み込み
        self.load_settings()
        
        # フォントサイズを適用
        self.apply_font_size()
        
        # 電話ボタン監視の初期化と開始
        self.phone_monitor = PhoneButtonMonitor(self.fetch_cti_data)
        self.phone_monitor.start_monitoring()
    
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
        
        # 顧客情報取得ボタン
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
        
        # 入力クリアボタン
        self.clear_btn = QPushButton("入力クリア")
        self.clear_btn.clicked.connect(self.clear_all_inputs)
        
        # 設定ボタン
        self.settings_btn = QPushButton("設定")
        self.settings_btn.clicked.connect(self.show_settings)
        
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
        
        # ボタンにスタイルを適用
        for btn in [self.clear_btn, self.settings_btn]:
            btn.setStyleSheet(button_style)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            top_bar_layout.addWidget(btn)
        
        parent_layout.addWidget(top_bar)
    
    def create_input_form(self, parent_layout):
        """入力フォームを作成"""
        # 住所情報セクション
        address_group = QGroupBox("住所情報")
        address_layout = QVBoxLayout()
        
        # 郵便番号
        address_layout.addWidget(QLabel("郵便番号"))
        self.postal_code_input = QLineEdit()
        self.postal_code_input.setPlaceholderText("例: 123-4567")
        address_layout.addWidget(self.postal_code_input)
        
        # 住所
        address_layout.addWidget(QLabel("住所"))
        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText("例: 大阪府大阪市中央区城見2-1-61")
        address_layout.addWidget(self.address_input)
        
        # 地図表示ボタン
        self.map_btn = QPushButton("地図を表示")
        self.map_btn.setStyleSheet("""
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
        self.map_btn.clicked.connect(self.show_map)
        address_layout.addWidget(self.map_btn)
        
        # 提供エリア検索ボタン
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
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.area_search_btn.clicked.connect(self.search_service_area)
        address_layout.addWidget(self.area_search_btn)
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFixedHeight(2)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #E3F2FD;
            }
            QProgressBar::chunk {
                background-color: #3498DB;
            }
        """)
        address_layout.addWidget(self.progress_bar)
        
        # 結果表示
        self.result_label = QLabel("提供エリア: 未検索")
        self.result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #95a5a6;
                border-radius: 4px;
                background-color: #f8f9fa;
                color: #95a5a6;
            }
        """)
        address_layout.addWidget(self.result_label)
        
        # スクリーンショットボタン
        self.screenshot_btn = QPushButton("スクリーンショット表示")
        self.screenshot_btn.setEnabled(False)
        self.screenshot_btn.setStyleSheet("""
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
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        address_layout.addWidget(self.screenshot_btn)
        
        address_group.setLayout(address_layout)
        parent_layout.addWidget(address_group)
    
    def load_settings(self):
        """設定ファイルを読み込む"""
        try:
            # 初期設定を設定
            self.settings = {
                'font_size': 11,
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
                    
            logging.info("設定を読み込みました")
                
        except Exception as e:
            logging.error(f"設定の読み込みに失敗しました: {str(e)}")
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
    
    def save_settings(self):
        """設定をファイルに保存"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            logging.info("設定を保存しました")
        except Exception as e:
            logging.error(f"設定の保存に失敗しました: {str(e)}")
            QMessageBox.warning(self, "エラー", f"設定の保存に失敗しました: {str(e)}")
    
    def show_settings(self):
        """設定ダイアログを表示"""
        dialog = SettingsDialog(self)
        # 現在の設定を設定
        dialog.set_settings(self.settings)
        
        if dialog.exec():
            try:
                # ダイアログの設定を取得
                new_settings = dialog.get_settings()
                # 設定を更新
                self.settings.update(new_settings)
                # 設定を保存
                self.save_settings()
                # 設定を再読み込み
                self.load_settings()
                # フォントサイズを適用
                self.apply_font_size()
                logging.info("設定を更新しました")
            except Exception as e:
                logging.error(f"設定の更新中にエラー: {str(e)}")
                QMessageBox.warning(self, "エラー", f"設定の更新中にエラーが発生しました: {str(e)}")
    
    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        self.postal_code_input.clear()
        self.address_input.clear()
        self.result_label.setText("提供エリア: 未検索")
        self.result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #95a5a6;
                border-radius: 4px;
                background-color: #f8f9fa;
                color: #95a5a6;
            }
        """)
    
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
            
            # 検索ボタンを無効化
            self.area_search_btn.setEnabled(False)
            self.area_search_btn.setText("検索中...")
            
            # プログレスバーを表示
            self.progress_bar.setVisible(True)
            
            # 検索ステータスを更新
            self.result_label.setText("提供エリア: 検索を開始します...")
            self.result_label.setStyleSheet("""
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
            self.thread.finished.connect(self.thread.deleteLater)  # スレッド終了時にクリーンアップ
            self.thread.start()
            
        except Exception as e:
            logging.error(f"検索処理の開始時にエラー: {str(e)}")
            self.result_label.setText("提供エリア: エラーが発生しました")
            self.result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #E74C3C;
                    border-radius: 4px;
                    background-color: #FFEBEE;
                    color: #E74C3C;
                }
            """)
            self.area_search_btn.setEnabled(True)
            self.area_search_btn.setText("検索")
            self.progress_bar.setVisible(False)
            self.cleanup_thread()
    
    def update_search_progress(self, message):
        """検索の進捗状況を更新"""
        self.result_label.setText(f"提供エリア: {message}")
        self.result_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                padding: 5px;
                border: 1px solid #3498DB;
                border-radius: 4px;
                background-color: #E3F2FD;
                color: #2980B9;
            }
        """)
        QApplication.processEvents()
    
    def on_search_completed(self, result):
        """検索完了時の処理"""
        # プログレスバーを非表示
        self.progress_bar.setVisible(False)
        self.area_search_btn.setEnabled(True)
        self.area_search_btn.setText("検索")
        
        try:
            status = result.get("status")
            message = result.get("message", "")
            details = result.get("details", {})
            
            if status == "available":
                self.result_label.setText("提供エリア: 提供可能")
                self.result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #27AE60;
                        border-radius: 4px;
                        background-color: #E8F5E9;
                        color: #27AE60;
                    }
                """)
            elif status == "unavailable":
                self.result_label.setText("提供エリア: 未提供")
                self.result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #E74C3C;
                        border-radius: 4px;
                        background-color: #FFEBEE;
                        color: #E74C3C;
                    }
                """)
            else:
                self.result_label.setText("提供エリア: 判定失敗")
                self.result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #F39C12;
                        border-radius: 4px;
                        background-color: #FFF3E0;
                        color: #F39C12;
                    }
                """)
            
            # 詳細情報がある場合は表示
            if details and self.settings.get('browser_settings', {}).get('show_popup', True):
                details_text = "\n".join([f"{k}: {v}" for k, v in details.items()])
                QMessageBox.information(self, "検索結果", details_text)
            
            # スクリーンショットパスを更新
            if "screenshot" in result:
                self.update_screenshot_button(result["screenshot"])
            
        except Exception as e:
            logging.error(f"結果の表示中にエラー: {str(e)}")
            self.result_label.setText("提供エリア: エラー")
            self.result_label.setStyleSheet("""
                QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #E74C3C;
                    border-radius: 4px;
                    background-color: #FFEBEE;
                    color: #E74C3C;
                }
            """)
            QMessageBox.warning(self, "エラー", str(e))
    
    def update_screenshot_button(self, screenshot_path=None):
        """スクリーンショットボタンを更新"""
        if hasattr(self, 'screenshot_btn'):
            if screenshot_path:
                self.screenshot_path = screenshot_path
                self.screenshot_btn.setEnabled(True)
            else:
                self.screenshot_btn.setEnabled(False)
    
    def fetch_cti_data(self):
        """CTIデータを取得"""
        try:
            data = self.cti_service.get_all_fields_data()
            if data:
                self.postal_code_input.setText(data.postal_code)
                self.address_input.setText(data.address)
                logging.info("CTIデータの取得に成功しました")
            else:
                logging.warning("CTIデータの取得に失敗しました")
                QMessageBox.warning(self, "エラー", "CTIデータの取得に失敗しました。\nCTIメインウィンドウが開いているか確認してください。")
        except Exception as e:
            logging.error(f"CTIデータの取得中にエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"CTIデータの取得中にエラーが発生しました: {str(e)}")
    
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
            label.setPixmap(pixmap)
            
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
    
    def cleanup_thread(self):
        """スレッドとワーカーをクリーンアップ"""
        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        if self.worker:
            self.worker.deleteLater()
        self.thread = None
        self.worker = None
    
    def closeEvent(self, event):
        """ウィンドウを閉じる際の処理"""
        # 電話ボタン監視を停止
        if hasattr(self, 'phone_monitor'):
            self.phone_monitor.stop_monitoring()
        
        # スレッドをクリーンアップ
        self.cleanup_thread()
        
        event.accept()
    
    def apply_font_size(self):
        """フォントサイズを適用する"""
        try:
            # 設定からフォントサイズを取得
            font_size = self.settings.get('font_size', 11)
            
            # アプリケーション全体のフォントを設定
            app = QApplication.instance()
            font = QFont()
            font.setPointSize(font_size)
            app.setFont(font)
            
            # スタイルシートを使用してフォントサイズを設定
            self.setStyleSheet(f"* {{ font-size: {font_size}pt; }}")
            
            # メインウィンドウの全てのウィジェットに対してフォントを再設定
            for widget in self.findChildren(QWidget):
                widget_font = widget.font()
                widget_font.setPointSize(font_size)
                widget.setFont(widget_font)
                widget.update()
            
            logging.info(f"フォントサイズを {font_size} に設定しました")
            
        except Exception as e:
            logging.error(f"フォントサイズ適用エラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"フォントサイズの適用に失敗しました: {str(e)}")
    
    def show_map(self):
        """Googleマップで住所を表示"""
        try:
            address = self.address_input.text().strip()
            if not address:
                QMessageBox.warning(self, "入力エラー", "住所を入力してください。")
                return
            
            # 住所をURLエンコード
            encoded_address = quote(address)
            
            # GoogleマップのURL
            map_url = f"https://www.google.com/maps/search/?api=1&query={encoded_address}"
            
            # デフォルトのブラウザで開く
            webbrowser.open(map_url)
            logging.info(f"地図を表示: {address}")
            
        except Exception as e:
            logging.error(f"地図表示エラー: {str(e)}")
            QMessageBox.critical(
                self,
                "エラー",
                f"地図の表示中にエラーが発生しました: {str(e)}"
            ) 