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
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineProfile, QWebEngineSettings

from services.area_search import search_service_area, set_cancel_flag, clear_cancel_flag
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
    
    # カスタムシグナルを追加（PySide6では型を明示的に指定）
    trigger_service_area_search = Signal()
    trigger_auto_search = Signal()
    
    def __init__(self):
        """メインウィンドウの初期化"""
        super().__init__()
        self.setWindowTitle("フレッツ光クロス用 - コールセンター業務効率化ツール")
        self.setMinimumSize(600, 400)
        
        # スレッドとワーカーの初期化
        self.thread = None
        self.worker = None
        
        # Google検索カウンターの初期化（5件ごとにWebView再初期化）
        self.google_search_count = 0
        self.webview_refresh_interval = 5
        
        # メインウィジェットとレイアウトの設定
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout()
        main_widget.setLayout(self.main_layout)
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # CTIサービスの初期化
        self.cti_service = OneClickService()
        
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
        
        # CTI状態監視の初期化と開始（設定に基づいて制御）
        cti_monitoring_enabled = self.settings.get('enable_cti_monitoring', True)
        cti_auto_processing_enabled = self.settings.get('enable_auto_cti_processing', True)
        logging.info(f"★★★ メイン設定読み込み ★★★")
        logging.info(f"- CTI監視設定: {cti_monitoring_enabled}")
        logging.info(f"- CTI自動処理設定: {cti_auto_processing_enabled}")
        
        if cti_monitoring_enabled:
            if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                self.cti_status_monitor = CTIStatusMonitor(
                    on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
                    on_call_ended_callback=self.on_cti_call_ended,
                    on_talking_started_callback=self.on_cti_talking_started,
                    on_cancel_processing_callback=self.on_cancel_processing_request
                )
                # 通話時間閾値を設定
                call_duration_threshold = self.settings.get('call_duration_threshold', 0)
                self.cti_status_monitor.set_call_duration_threshold(call_duration_threshold)
                
                # 自動処理設定を反映
                self.cti_status_monitor.enable_auto_processing = cti_auto_processing_enabled
                logging.info(f"- 自動処理設定を反映: {cti_auto_processing_enabled}")
                
                # CTI自動処理用のシグナル・スロット接続（CTI監視開始前に設定）
                self._setup_auto_search_signal_connection()
                logging.info("- trigger_auto_search シグナル・スロット接続を設定しました（QueuedConnection使用）")
                
                # 接続状態を確認
                try:
                    # テスト用のシグナル発行
                    logging.info("★★★ シグナル・スロット接続テストを実行します ★★★")
                    QTimer.singleShot(100, lambda: self._test_signal_connection())
                except Exception as e:
                    logging.error(f"シグナル・スロット接続テスト中にエラー: {str(e)}")
                
                # シグナル・スロット接続完了後にCTI監視を開始
                self.cti_status_monitor.start_monitoring()
                logging.info("CTI状態監視を開始しました")
        else:
            logging.info("CTI監視が設定で無効になっています")
            self.cti_status_monitor = None
        
        # 自動処理の重複実行防止用フラグ
        if not hasattr(self, 'is_auto_processing'):
            self.is_auto_processing = False
            self.last_auto_processing_time = 0
        
        # 復旧システム用のタイマー
        self.recovery_timer = QTimer()
        self.recovery_timer.setSingleShot(True)
        self.recovery_timer.timeout.connect(self.force_reset_processing)
        
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
        address_group.setObjectName("address_group")  # オブジェクト名を設定
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
        
        # 電話番号（任意）
        address_layout.addWidget(QLabel("電話番号（任意）"))
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("例: 0312345678 または 03-1234-5678")
        address_layout.addWidget(self.phone_input)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
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
        button_layout.addWidget(self.map_btn)
        
        # WebViewリフレッシュボタン
        self.refresh_webview_btn = QPushButton("Google更新")
        self.refresh_webview_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:pressed {
                background-color: #E65100;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.refresh_webview_btn.clicked.connect(self.manual_refresh_webview)
        self.refresh_webview_btn.setToolTip("Google検索ウィンドウを手動で更新（reCAPTCHA対策）")
        button_layout.addWidget(self.refresh_webview_btn)
        
        address_layout.addLayout(button_layout)
        
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
        
        # 結果表示エリア（横並び配置用）
        self.result_layout = QHBoxLayout()
        
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
        self.result_layout.addWidget(self.result_label)
        self.result_layout.addStretch()  # 右側にスペースを追加
        
        # 結果表示エリアをメインレイアウトに追加
        address_layout.addLayout(self.result_layout)
        
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
        
        # 検索回数表示ラベル
        self.search_count_label = QLabel("Google検索回数: 0回")
        self.search_count_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #666666;
                padding: 2px;
                background-color: #f0f0f0;
                border-radius: 3px;
                margin: 2px;
            }
        """)
        address_layout.addWidget(self.search_count_label)
        
        # QWebEngineView（Google検索結果表示用）をここで生成・追加
        self.web_view = QWebEngineView()
        self.web_view.setVisible(False)  # 初期表示は非表示
        self.web_view.setMinimumHeight(300)
        self.web_view.setMaximumHeight(500)
        self.web_view.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        address_layout.addWidget(self.web_view)
        
        # 住所情報グループをレイアウトに追加
        address_group.setLayout(address_layout)
        parent_layout.addWidget(address_group)
    
    def load_settings(self):
        """設定ファイルを読み込む"""
        try:
            # 初期設定を設定
            self.settings = {
                'font_size': 11,
                'call_duration_threshold': 0,
                'enable_cti_monitoring': True,
                'enable_auto_cti_processing': True,
                'cti_monitor_interval': 0.5,
                'cti_settings': {
                    'enable_cti': True,
                    'enable_auto_cti_processing': True,
                    'cti_monitor_interval': 0.5
                },
                'browser_settings': {
                    'headless': True,
                    'disable_images': True,
                    'show_popup': True,
                    'auto_close': True,
                    'page_load_timeout': 60,
                    'script_timeout': 60,
                    'use_external_browser': False
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
            
            # WebView再初期化間隔を設定から取得
            self.webview_refresh_interval = self.settings.get('webview_refresh_interval', 5)
            logging.info(f"設定を読み込みました - WebView再初期化間隔: {self.webview_refresh_interval}件")
                
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
                logging.info("設定を更新しました")
            except Exception as e:
                logging.error(f"設定の更新中にエラー: {str(e)}")
                QMessageBox.warning(self, "エラー", f"設定の更新中にエラーが発生しました: {str(e)}")
    
    def clear_all_inputs(self):
        """全ての入力フィールドをクリア"""
        self.postal_code_input.clear()
        self.address_input.clear()
        self.phone_input.clear()
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
        
        # Google検索カウンターもリセット
        self.google_search_count = 0
        self.update_search_count_display()
        logging.info("Google検索カウンターをリセットしました")
    
    def search_service_area(self):
        """提供エリア検索を開始"""
        postal_code = self.postal_code_input.text().strip()
        address = self.address_input.text().strip()
        phone = self.phone_input.text().strip()
        
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
            
            # --- ここからGoogle検索埋め込み処理 ---
            self.start_google_search_embed(phone, address)
            # --- ここまで ---
            
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
            
            # 復旧ボタンを表示
            self.show_recovery_buttons()
            
            # 5秒後に強制リセット
            self.recovery_timer.start(5000)
            
            # キャンセル完了後に検索ボタンを初期状態に戻す
            self.reset_search_button()
            return
        
        # キャンセル以外の完了時の処理
        self.reset_search_button()
        
        # 復旧ボタンを非表示
        self.hide_recovery_buttons()
        
        # 復旧タイマーを停止
        if hasattr(self, 'recovery_timer'):
            self.recovery_timer.stop()
        
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
            elif status == "investigation":
                self.result_label.setText("提供エリア: 要調査")
                self.result_label.setStyleSheet("""
                    QLabel {
                        font-size: 14px;
                        padding: 5px;
                        border: 1px solid #F1C40F;
                        border-radius: 4px;
                        background-color: #FFF9E3;
                        color: #B7950B;
                        font-weight: bold;
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
                self.phone_input.setText(data.phone)
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
        try:
            # CTI状態監視を停止
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                self.cti_status_monitor.stop_monitoring()
                logging.info("CTI状態監視を停止しました")
            
            # 電話ボタン監視を停止
            if hasattr(self, 'phone_monitor'):
                self.phone_monitor.stop_monitoring()
                logging.info("電話ボタン監視を停止しました")
            
            event.accept()
        except Exception as e:
            logging.error(f"アプリケーション終了処理中にエラー: {e}")
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
    
    def manual_refresh_webview(self):
        """
        手動でWebViewをリフレッシュする（reCAPTCHA対策）
        """
        try:
            logging.info("手動でWebViewをリフレッシュします")
            self.refresh_webview()
            
            # リフレッシュ後に現在の検索を再実行
            phone = self.phone_input.text().strip()
            address = self.address_input.text().strip()
            
            if phone or address:
                # カウンターを減らして再実行（重複カウントを避ける）
                self.google_search_count -= 1
                self.start_google_search_embed(phone, address)
                logging.info("検索を再実行しました")
            else:
                QMessageBox.information(
                    self, 
                    "情報", 
                    "検索ウィンドウをリフレッシュしました。\n電話番号または住所を入力して検索を行ってください。"
                )
                
        except Exception as e:
            logging.error(f"手動WebViewリフレッシュエラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"検索ウィンドウのリフレッシュに失敗しました: {str(e)}")
    
    def update_search_count_display(self):
        """
        Google検索回数の表示を更新する
        """
        try:
            next_refresh = self.webview_refresh_interval - (self.google_search_count % self.webview_refresh_interval)
            if next_refresh == self.webview_refresh_interval:
                next_refresh = 0
            
            if hasattr(self, 'search_count_label'):
                if next_refresh == 0:
                    self.search_count_label.setText(f"Google検索回数: {self.google_search_count}回 (次回更新)")
                    self.search_count_label.setStyleSheet("""
                        QLabel {
                            font-size: 12px;
                            color: #E65100;
                            padding: 2px;
                            background-color: #FFF3E0;
                            border: 1px solid #FF9800;
                            border-radius: 3px;
                            margin: 2px;
                            font-weight: bold;
                        }
                    """)
                else:
                    self.search_count_label.setText(f"Google検索回数: {self.google_search_count}回 (あと{next_refresh}回で更新)")
                    self.search_count_label.setStyleSheet("""
                        QLabel {
                            font-size: 12px;
                            color: #666666;
                            padding: 2px;
                            background-color: #f0f0f0;
                            border-radius: 3px;
                            margin: 2px;
                        }
                    """)
        except Exception as e:
            logging.error(f"検索回数表示更新エラー: {str(e)}")
    
    def refresh_webview(self):
        """WebViewを再初期化する"""
        try:
            self.setup_webview()
            logging.debug("WebViewを再初期化しました")
        except Exception as e:
            logging.error(f"WebViewの再初期化中にエラー: {str(e)}")

    def setup_webview(self):
        """WebViewの初期化"""
        try:
            from PySide6.QtCore import Qt
            
            # 外部ブラウザ設定を確認
            use_external_browser = self.settings.get('browser_settings', {}).get('use_external_browser', False)
            
            # WebViewのレイアウトを取得または作成
            if not hasattr(self, 'webview_layout'):
                self.webview_layout = QVBoxLayout()
                self.right_layout.addLayout(self.webview_layout)
            else:
                # 既存のウィジェットをクリア
                while self.webview_layout.count():
                    item = self.webview_layout.takeAt(0)
                    if item.widget():
                        item.widget().deleteLater()
            
            # 既存のWebViewをクリア
            if hasattr(self, 'web_view') and self.web_view is not None:
                self.web_view.setParent(None)
                self.web_view.deleteLater()
                self.web_view = None
            
            # 外部ブラウザ設定が無効の場合のみWebViewを作成
            if not use_external_browser:
                self.web_view = QWebEngineView()
                self.web_view.setUrl("about:blank")
                self.webview_layout.addWidget(self.web_view)
                logging.debug("WebViewを初期化しました")
            else:
                # 外部ブラウザ設定が有効な場合は代わりにメッセージを表示
                message_label = QLabel("電話番号検索は外部ブラウザで開く設定になっています")
                message_label.setAlignment(Qt.AlignCenter)
                self.webview_layout.addWidget(message_label)
                logging.debug("外部ブラウザ設定が有効のため、WebViewは無効化されています")
            
        except Exception as e:
            logging.error(f"WebViewの初期化中にエラー: {str(e)}")

    def start_google_search_embed(self, phone, address):
        """
        電話番号または住所でGoogle検索し、QWebEngineViewに結果を表示する
        .osrp-blkがあればその部分まで自動スクロール（右端）する
        5件ごとにWebViewを再初期化してreCAPTCHA対策を行う
        """
        try:
            import time
            import random
            import webbrowser
            
            # 検索回数をカウント
            self.google_search_count += 1
            logging.debug(f"Google検索実行: {self.google_search_count}回目")
            
            # 検索回数表示を更新
            self.update_search_count_display()
            
            # 検索クエリを作成
            search_query = ""
            if phone:
                # ハイフンなしで検索
                phone_no_hyphen = phone.replace("-", "")
                search_query = phone_no_hyphen
            else:
                # 住所で検索
                search_query = address
            
            url = f"https://www.google.com/search?q={search_query}"
            logging.debug(f"Google検索URL: {url}")
            
            # 外部ブラウザ設定を確認
            use_external_browser = self.settings.get('browser_settings', {}).get('use_external_browser', False)
            
            if use_external_browser:
                # 外部ブラウザで開く
                logging.debug("外部ブラウザで検索を実行します")
                webbrowser.open(url)
                return
            
            # WebViewが無効な場合は処理を終了
            if not hasattr(self, 'web_view') or self.web_view is None:
                logging.debug("WebViewが無効なため、検索を中止します")
                return
            
            # 連続検索の間隔調整（reCAPTCHA対策）
            if self.google_search_count > 1:
                # 2回目以降は少し間隔を空ける
                delay = random.uniform(0.5, 1.5)
                time.sleep(delay)
                logging.debug(f"検索間隔調整: {delay:.1f}秒待機")
            
            # 指定件数ごとにWebViewを再初期化
            if self.google_search_count % self.webview_refresh_interval == 0:
                logging.debug(f"{self.webview_refresh_interval}回目の検索のため、WebViewを再初期化します")
                self.refresh_webview()
            
            self.web_view.setVisible(False)
            self.web_view.setUrl(url)
            self.web_view.setVisible(True)
            
            def scroll_to_osrp_blk():
                js = """
                    (function(){
                        var blk = document.querySelector('.osrp-blk');
                        if (blk) blk.scrollIntoView({behavior: 'auto', block: 'nearest', inline: 'end'});
                    })();
                """
                self.web_view.page().runJavaScript(js)
            self.web_view.loadFinished.connect(scroll_to_osrp_blk)
            
        except Exception as e:
            logging.error(f"Google検索埋め込み処理エラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"検索処理でエラーが発生しました: {str(e)}")

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

    def on_cti_dialing_to_talking(self):
        """発信中→通話中の状態変化時のコールバック処理"""
        try:
            import time
            current_time = time.time()
            
            logging.info("★★★ on_cti_dialing_to_talking コールバックが呼び出されました ★★★")
            logging.info(f"- 自動処理中フラグ: {self.is_auto_processing}")
            
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
            
            # 重要な状態変化なのでINFOレベルで出力
            logging.info("CTI状態変化検出: 発信中 → 通話中")
            logging.info("自動処理を開始します: 顧客情報取得 → 提供判定検索")
            
                        # 顧客情報を取得
            data = self.cti_service.get_all_fields_data()
            if data:
                # 入力フィールドに設定
                self.postal_code_input.setText(data.postal_code)
                self.address_input.setText(data.address)
                self.phone_input.setText(data.phone)
                logging.debug("CTIデータの取得に成功しました")
                
                # 自動検索シグナルを発行（TelephoneToolと同じ方式）
                logging.info("★★★ trigger_auto_search.emit() を実行します ★★★")
                # 1秒後にシグナルを発行（TelephoneToolと同じタイミング）
                import threading
                def delayed_trigger():
                    try:
                        self.trigger_auto_search.emit()
                        logging.info("★★★ trigger_auto_search.emit() を実行しました ★★★")
                    except Exception as e:
                        logging.error(f"シグナル送信中にエラー: {str(e)}")
                        # エラー時はフラグをリセット
                        self.is_auto_processing = False
                        logging.debug("自動処理フラグをリセットしました")
                        
                timer = threading.Timer(1.0, delayed_trigger)
                timer.daemon = True
                timer.start()
            else:
                logging.warning("CTIデータの取得に失敗しました")
                QMessageBox.warning(self, "エラー", "CTIデータの取得に失敗しました。\nCTIメインウィンドウが開いているか確認してください。")
                
        except Exception as e:
            logging.error(f"発信中→通話中の処理でエラーが発生: {str(e)}")
            # エラー時もフラグをリセット
            if hasattr(self, 'is_auto_processing'):
                self.is_auto_processing = False

    def on_cti_call_ended(self):
        """通話終了時のコールバック処理"""
        try:
            # 通話終了は重要な状態変化なのでINFOレベルで出力
            logging.info("通話が終了しました")
        except Exception as e:
            logging.error(f"通話終了時の処理でエラーが発生: {str(e)}")

    def on_cti_talking_started(self):
        """通話中状態開始時のコールバック処理"""
        try:
            # 通話開始は頻繁に発生するのでDEBUGレベルで出力
            logging.debug("通話中状態を検出")
        except Exception as e:
            logging.error(f"通話中状態開始時の処理でエラーが発生: {str(e)}")

    def apply_cti_settings(self):
        """CTI監視設定を適用する"""
        try:
            cti_settings = self.settings.get('cti_settings', {})
            
            # CTI監視の有効/無効を設定
            if cti_settings.get('enable_cti', True):
                if not hasattr(self, 'cti_status_monitor') or self.cti_status_monitor is None:
                    self.cti_status_monitor = CTIStatusMonitor(
                        on_dialing_to_talking_callback=self.on_cti_dialing_to_talking,
                        on_call_ended_callback=self.on_cti_call_ended,
                        on_talking_started_callback=self.on_cti_talking_started,
                        on_cancel_processing_callback=self.on_cancel_processing_request
                    )
                    # 通話時間閾値を設定
                    call_duration_threshold = self.settings.get('call_duration_threshold', 0)
                    self.cti_status_monitor.set_call_duration_threshold(call_duration_threshold)
                    
                    # シグナル・スロット接続を設定（apply_cti_settingsでも必要）
                    self._setup_auto_search_signal_connection()
                    logging.info("- trigger_auto_search シグナル・スロット接続を設定しました（apply_cti_settings内）")
                    
                if not self.cti_status_monitor.is_monitoring:
                    self.cti_status_monitor.start_monitoring()
                    logging.info("CTI監視を開始しました")
            else:
                if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                    if self.cti_status_monitor.is_monitoring:
                        self.cti_status_monitor.stop_monitoring()
                        logging.info("CTI監視を停止しました")
            
            # 自動処理の有効/無効を設定
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor is not None:
                auto_processing_enabled = cti_settings.get('enable_auto_cti_processing', True)
                self.cti_status_monitor.enable_auto_processing = auto_processing_enabled
                
                # 監視間隔を設定
                monitor_interval = cti_settings.get('cti_monitor_interval', 0.5)
                self.cti_status_monitor.monitor_interval = monitor_interval
                
                logging.info(f"★★★ CTI監視設定を更新 ★★★")
                logging.info(f"- 自動処理有効: {auto_processing_enabled}")
                logging.info(f"- 監視間隔: {monitor_interval}秒")
            
            logging.info("CTI監視設定を適用しました")
            
        except Exception as e:
            logging.error(f"CTI監視設定の適用中にエラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"CTI監視設定の適用中にエラーが発生しました: {str(e)}")
    
    def on_cancel_processing_request(self):
        """
        処理キャンセル要求を受信した時の処理
        """
        try:
            logging.info("★★★ 処理キャンセル要求を受信しました ★★★")
            
            # エリア検索のキャンセルフラグを設定
            set_cancel_flag()
            logging.info("- エリア検索のキャンセルフラグを設定しました")
            
            # 現在のUIボタン状態をログ出力
            if hasattr(self, 'area_search_btn'):
                button_text = self.area_search_btn.text()
                button_enabled = self.area_search_btn.isEnabled()
                logging.info(f"- 現在のUIボタン状態: '{button_text}' (有効: {button_enabled})")
                
                # キャンセルボタンが表示されている場合はクリック
                if button_text == "キャンセル" and button_enabled:
                    logging.info("- UIキャンセルボタンをプログラム的にクリックします")
                    self.area_search_btn.click()
                else:
                    logging.info("- キャンセルボタンが無効または非表示のため、直接キャンセル処理を実行")
                    self.cancel_search()
            
            # ワーカーとスレッドの状態をログ出力
            worker_running = hasattr(self, 'worker') and self.worker is not None
            thread_running = hasattr(self, 'thread') and self.thread is not None and self.thread.isRunning()
            logging.info(f"- ワーカー実行中: {worker_running}")
            logging.info(f"- スレッド実行中: {thread_running}")
            
            # 自動処理フラグをリセット
            self.is_auto_processing = False
            logging.info("- 自動処理フラグをリセットしました")
            
        except Exception as e:
            logging.error(f"キャンセル処理要求の処理中にエラー: {str(e)}")
    
    @Slot()
    def auto_search_service_area(self):
        """
        自動提供判定検索（UIボタン経由で統一）
        """
        try:
            logging.info("★★★★★ auto_search_service_area が呼び出されました ★★★★★")
            logging.info("★★★★★ シグナル・スロット接続が正常に動作しています ★★★★★")
            
            # 自動処理フラグをリセット（処理開始時）
            self.is_auto_processing = False
            
            # 入力データの確認
            postal_code = self.postal_code_input.text().strip()
            address = self.address_input.text().strip()
            
            if not postal_code or not address:
                logging.warning("郵便番号または住所が空のため、自動検索をスキップします")
                return
            
            logging.info(f"検索データ - 郵便番号: {postal_code}, 住所: {address}")
            
            # UIボタン経由で検索を実行（手動実行と完全に統一）
            if hasattr(self, 'area_search_btn') and self.area_search_btn.isEnabled():
                logging.info("★★★ 提供エリア検索ボタンをプログラム的にクリックします ★★★")
                self.area_search_btn.click()
                logging.info("★★★ 自動提供判定検索が正常に開始されました ★★★")
            else:
                button_text = self.area_search_btn.text() if hasattr(self, 'area_search_btn') else "不明"
                button_enabled = self.area_search_btn.isEnabled() if hasattr(self, 'area_search_btn') else False
                logging.warning(f"提供エリア検索ボタンが無効のため、自動検索をスキップします - ボタン状態: '{button_text}' (有効: {button_enabled})")
                
        except Exception as e:
            logging.error(f"自動提供判定検索中にエラー: {str(e)}")
            # エラー時も自動処理フラグをリセット
            self.is_auto_processing = False
    
    def force_reset_processing(self):
        """
        強制的に処理状態をリセット（タイムアウト対策）
        """
        try:
            logging.warning("★★★ 強制リセット処理を実行します ★★★")
            
            # キャンセルフラグをクリア
            clear_cancel_flag()
            
            # ワーカーとスレッドを強制終了
            if hasattr(self, 'worker') and self.worker:
                self.worker.cancel()
                self.worker.deleteLater()
                self.worker = None
                logging.info("- ワーカーを強制終了しました")
            
            if hasattr(self, 'thread') and self.thread and self.thread.isRunning():
                self.thread.quit()
                self.thread.wait(1000)  # 1秒待機
                if self.thread.isRunning():
                    self.thread.terminate()
                    self.thread.wait()
                self.thread = None
                logging.info("- スレッドを強制終了しました")
            
            # UI状態をリセット
            self.reset_search_button()
            self.progress_bar.setVisible(False)
            self.result_label.setText("提供エリア: 未検索")
            self.result_label.setStyleSheet("")
            
            # 自動処理フラグをリセット
            self.is_auto_processing = False
            
            # CTI監視を再開
            if hasattr(self, 'cti_status_monitor') and self.cti_status_monitor:
                if not self.cti_status_monitor.is_monitoring:
                    self.cti_status_monitor.start_monitoring()
                    logging.info("- CTI監視を再開しました")
            
            logging.info("★★★ 強制リセット処理が完了しました ★★★")
            
        except Exception as e:
            logging.error(f"強制リセット処理中にエラー: {str(e)}")
    
    def reload_application(self):
        """
        アプリケーションをリロード（初期状態に戻す）
        """
        try:
            reply = QMessageBox.question(
                self, 
                "確認", 
                "アプリケーションを初期状態に戻しますか？\n（入力データはクリアされます）",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                logging.info("★★★ アプリケーションリロードを実行します ★★★")
                
                # 強制リセット処理を実行
                self.force_reset_processing()
                
                # 入力データをクリア
                self.clear_all_inputs()
                
                # 復旧ボタンを非表示
                self.hide_recovery_buttons()
                
                logging.info("★★★ アプリケーションリロードが完了しました ★★★")
                QMessageBox.information(self, "完了", "アプリケーションが初期状態に戻りました")
                
        except Exception as e:
            logging.error(f"アプリケーションリロード中にエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"リロード中にエラーが発生しました: {str(e)}")
    
    def restart_application(self):
        """
        アプリケーションを完全に再起動
        """
        try:
            reply = QMessageBox.question(
                self, 
                "確認", 
                "アプリケーションを完全に再起動しますか？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                logging.info("★★★ アプリケーション再起動を実行します ★★★")
                
                # 新しいプロセスを起動
                import subprocess
                import sys
                subprocess.Popen([sys.executable] + sys.argv)
                
                # 現在のプロセスを終了
                QApplication.quit()
                
        except Exception as e:
            logging.error(f"アプリケーション再起動中にエラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"再起動中にエラーが発生しました: {str(e)}")
    
    def show_recovery_buttons(self):
        """
        復旧ボタンを表示
        """
        try:
            if not hasattr(self, 'recovery_buttons_widget'):
                # 復旧ボタン用のウィジェットを作成
                self.recovery_buttons_widget = QWidget()
                recovery_layout = QHBoxLayout(self.recovery_buttons_widget)
                recovery_layout.setContentsMargins(0, 5, 0, 0)
                
                # リロードボタン
                self.reload_btn = QPushButton("リロード")
                self.reload_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #FF9800;
                        color: white;
                        border: none;
                        padding: 6px 12px;
                        text-align: center;
                        font-size: 12px;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: #F57C00;
                    }
                    QPushButton:pressed {
                        background-color: #E65100;
                    }
                """)
                self.reload_btn.clicked.connect(self.reload_application)
                self.reload_btn.setToolTip("アプリケーションを初期状態に戻す")
                
                # 再起動ボタン
                self.restart_btn = QPushButton("再起動")
                self.restart_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #F44336;
                        color: white;
                        border: none;
                        padding: 6px 12px;
                        text-align: center;
                        font-size: 12px;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: #D32F2F;
                    }
                    QPushButton:pressed {
                        background-color: #B71C1C;
                    }
                """)
                self.restart_btn.clicked.connect(self.restart_application)
                self.restart_btn.setToolTip("アプリケーションを完全に再起動")
                
                recovery_layout.addWidget(self.reload_btn)
                recovery_layout.addWidget(self.restart_btn)
                recovery_layout.addStretch()
                
                # 結果表示エリアの下に追加
                if hasattr(self, 'result_layout'):
                    self.result_layout.addWidget(self.recovery_buttons_widget)
            
            self.recovery_buttons_widget.setVisible(True)
            
        except Exception as e:
            logging.error(f"復旧ボタン表示中にエラー: {str(e)}")
    
    def hide_recovery_buttons(self):
        """
        復旧ボタンを非表示
        """
        try:
            if hasattr(self, 'recovery_buttons_widget'):
                self.recovery_buttons_widget.setVisible(False)
        except Exception as e:
            logging.error(f"復旧ボタン非表示中にエラー: {str(e)}")

    def _setup_auto_search_signal_connection(self):
        """
        auto_search用のシグナル・スロット接続を安全に設定する
        """
        try:
            # 既存の接続を安全に切断
            try:
                # 接続されているかチェック
                if hasattr(self.trigger_auto_search, 'disconnect'):
                    self.trigger_auto_search.disconnect()
            except RuntimeError:
                # 既に切断されている場合は無視
                pass
            except Exception as e:
                logging.debug(f"シグナル切断時の警告（無視可能）: {str(e)}")
            
            # 新しい接続を設定（@Slotデコレータ付きメソッドに接続）
            self.trigger_auto_search.connect(
                self.auto_search_service_area, 
                Qt.ConnectionType.QueuedConnection
            )
            
            logging.info("★★★ シグナル・スロット接続を安全に設定しました（@Slot付きメソッド） ★★★")
            
            # 接続テストを実行
            QTimer.singleShot(50, self._test_signal_connection)
            
        except Exception as e:
            logging.error(f"シグナル・スロット接続設定中にエラー: {str(e)}")
    
    def _test_signal_connection(self):
        """
        シグナル・スロット接続をテストする
        """
        try:
            logging.info("★★★ シグナル・スロット接続テスト開始 ★★★")
            self.trigger_auto_search.emit()
            logging.info("★★★ テスト用シグナルを発行しました ★★★")
        except Exception as e:
            logging.error(f"シグナル・スロット接続テスト中にエラー: {str(e)}")