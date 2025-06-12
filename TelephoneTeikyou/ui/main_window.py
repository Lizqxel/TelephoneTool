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
from PySide6.QtCore import Qt, QThread, QObject, Signal, Slot, QEvent, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QIcon, QPixmap, QCursor

from services.area_search import search_service_area
from ui.settings_dialog import SettingsDialog
from services.oneclick import OneClickService
from services.phone_button_monitor import PhoneButtonMonitor
from services.cti_status_monitor import CTIStatusMonitor


class CancellationError(Exception):
    """検索キャンセル時に発生する例外"""
    pass


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


class MainWindow(QMainWindow):
    """メインウィンドウクラス"""
    
    # カスタムシグナルを追加
    trigger_service_area_search = Signal()
    
    def __init__(self):
        """メインウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("フレッツ光クロス用 - コールセンター業務効率化ツール")
        self.setMinimumSize(600, 400)
        
        # スレッドとワーカーの初期化
        self.thread = None
        self.worker = None
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout()
        main_widget.setLayout(self.main_layout)
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # CTIサービスの初期化
        self.cti_service = OneClickService()
        
        # CTI状態監視の初期化と開始
        self.cti_status_monitor = CTIStatusMonitor(
            on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
            on_call_ended_callback=self.on_cti_call_ended,
            on_talking_started_callback=self.on_cti_talking_started
        )
        self.cti_status_monitor.start_monitoring()
        
        # トップバーを作成し、メインレイアウトに追加
        self.create_top_bar(self.main_layout)
        
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
        self.main_layout.addWidget(scroll_area)
        
        # 設定の読み込み
        self.load_settings()
        
        # フォントサイズを適用
        self.apply_font_size()
        
        # 電話ボタン監視の初期化と開始
        self.phone_monitor = PhoneButtonMonitor(self.fetch_cti_data)
        self.phone_monitor.start_monitoring()
        
        # シグナルをスロットに接続
        self.trigger_service_area_search.connect(self._execute_service_area_search)
    
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
                background-color: #2ECC71;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #27AE60;
            }
            QPushButton:pressed {
                background-color: #219A52;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.area_search_btn.clicked.connect(self.search_service_area)
        address_layout.addWidget(self.area_search_btn)
        
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
        address_layout.addWidget(self.progress_bar)
        
        # 結果表示ラベル
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
        self.screenshot_btn = QPushButton("スクリーンショットを表示")
        self.screenshot_btn.setEnabled(False)
        self.screenshot_btn.setStyleSheet("""
            QPushButton {
                background-color: #9B59B6;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #8E44AD;
            }
            QPushButton:pressed {
                background-color: #7D3C98;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.screenshot_btn.clicked.connect(self.show_screenshot)
        address_layout.addWidget(self.screenshot_btn)
        
        # 住所情報グループをレイアウトに追加
        address_group.setLayout(address_layout)
        parent_layout.addWidget(address_group)
        
        # 垂直方向のスペーサーを追加
        parent_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
    
    def load_settings(self):
        """設定ファイルを読み込む"""
        try:
            # 初期設定を設定
            self.settings = {
                'font_size': 11,
                'cti_settings': {
                    'enable_cti': True,
                    'enable_auto_cti_processing': True,
                    'cti_monitor_interval': 0.2
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
                    
            logging.info("設定を読み込みました")
            
            # CTI監視設定を適用
            self.apply_cti_settings()
                
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
                # CTI監視設定を適用
                self.apply_cti_settings()
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
        self.result_label.setText("提供エリア: キャンセル中...")
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

        # バックエンド処理のキャンセル
        if hasattr(self, 'worker'):
            self.worker.cancel()
            # キャンセル完了を待つため、ボタンとプログレスバーはそのまま維持

    def reset_search_button(self):
        """検索ボタンを初期状態に戻す"""
        self.area_search_btn.setEnabled(True)
        self.area_search_btn.setText("提供エリア検索")
        self.area_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ECC71;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #27AE60;
            }
            QPushButton:pressed {
                background-color: #219A52;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.area_search_btn.clicked.disconnect()
        self.area_search_btn.clicked.connect(self.search_service_area)

    def on_search_completed(self, result):
        """検索完了時の処理"""
        # プログレスバーを非表示
        self.progress_bar.setVisible(False)
        
        status = result.get("status", "failure")
        
        if status == "cancelled":
            self.result_label.setText("提供エリア: 検索がキャンセルされました")
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
            # キャンセル完了後に検索ボタンを初期状態に戻す
            self.reset_search_button()
            return
        
        # キャンセル以外の完了時の処理
        self.reset_search_button()
        
        try:
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
            elif status == "apartment":
                self.result_label.setText("提供エリア: 集合住宅")
                self.result_label.setStyleSheet("""
                    QLabel {
                    font-size: 14px;
                    padding: 5px;
                    border: 1px solid #FF9800;
                    border-radius: 4px;
                    background-color: #FFF3E0;
                    color: #E65100;
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
        # CTI状態監視を停止
        if hasattr(self, 'cti_status_monitor'):
            self.cti_status_monitor.stop_monitoring()
        
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
            self.result_label.setText(message)
            self.result_label.setStyleSheet("color: #666666;")
            
        except Exception as e:
            logging.error(f"進捗更新中にエラー: {str(e)}")
            self.result_label.setText(message)
    
    def on_cti_dialing_to_talking(self):
        """発信中→通話中の状態変化時のコールバック処理"""
        try:
            logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
            logging.info("自動処理を開始します: 顧客情報取得 → 提供判定検索")
            
            # 顧客情報取得を実行
            data = self.cti_service.get_all_fields_data()
            if data:
                # 入力フィールドに設定
                self.postal_code_input.setText(data.postal_code)
                self.address_input.setText(data.address)
                logging.info("CTIデータの取得に成功しました")
                
                # メインスレッドで提供判定検索を実行するようにシグナルを発行
                self.trigger_service_area_search.emit()
                logging.info("提供判定検索を要求しました")
            else:
                logging.warning("CTIデータの取得に失敗しました")
                QMessageBox.warning(self, "エラー", "CTIデータの取得に失敗しました。\nCTIメインウィンドウが開いているか確認してください。")
                
        except Exception as e:
            logging.error(f"発信中→通話中の自動処理でエラーが発生: {str(e)}")
            QMessageBox.critical(self, "エラー", f"自動処理中にエラーが発生しました: {str(e)}")

    def on_cti_talking_started(self):
        """通話中状態開始時のコールバック処理"""
        try:
            logging.info("★★★ 通話中状態を検出 ★★★")
        except Exception as e:
            logging.error(f"通話中状態開始時の処理でエラーが発生: {str(e)}")

    def on_cti_call_ended(self):
        """通話終了時（通話中→待ち受け中）のコールバック処理"""
        try:
            logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
        except Exception as e:
            logging.error(f"通話終了時の処理でエラーが発生: {str(e)}")

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
            
            logging.info("CTI監視設定を適用しました")
            
        except Exception as e:
            logging.error(f"CTI監視設定の適用中にエラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"CTI監視設定の適用中にエラーが発生しました: {str(e)}")

    def _execute_service_area_search(self):
        """提供判定検索を実行（メインスレッドで実行）"""
        try:
            postal_code = self.postal_code_input.text().strip()
            address = self.address_input.text().strip()
            logging.info(f"自動検索準備 - 郵便番号: {postal_code}, 住所: {address}")
            
            if postal_code and address:
                logging.info("提供判定検索を自動実行します")
                logging.info(f"検索ボタンの状態 - 有効: {self.area_search_btn.isEnabled()}, 表示: {self.area_search_btn.isVisible()}")
                
                # 検索ボタンの状態を確認して実行
                if self.area_search_btn.isEnabled() and self.area_search_btn.isVisible():
                    self.search_service_area()
                    logging.info("提供判定検索を実行しました")
                else:
                    logging.warning("検索ボタンが無効または非表示のため、検索をスキップします")
            else:
                logging.warning(f"郵便番号または住所が空のため、提供判定検索をスキップします")
        except Exception as e:
            logging.error(f"提供判定検索の実行時にエラー: {str(e)}") 