"""
設定ダイアログ

このモジュールは、アプリケーションの設定を管理するための
ダイアログUIを提供します。
"""

import json
import os
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                              QPushButton, QMessageBox, QSlider,
                              QGroupBox, QCheckBox, QScrollArea, QWidget)
from PySide6.QtCore import Qt


class SettingsDialog(QDialog):
    """設定ダイアログクラス"""
    
    def __init__(self, parent=None):
        """
        設定ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
        """
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setFixedSize(700, 400)  # ダイアログサイズを固定
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # デフォルトのフォントサイズ
        self.default_font_size = 9
        
        # デフォルトのブラウザ設定
        self.default_browser_settings = {
            "headless": False,
            "disable_images": True,
            "show_popup": True,
            "page_load_timeout": 30,
            "script_timeout": 30,
            "auto_close": False
        }
        
        # 親ウィンドウから現在の設定を取得
        if parent and hasattr(parent, 'settings'):
            self.current_font_size = parent.settings.get('font_size', self.default_font_size)
            self.current_browser_settings = parent.settings.get('browser_settings', self.default_browser_settings)
        else:
            self.current_font_size = self.default_font_size
            self.current_browser_settings = self.default_browser_settings
        
        # メインレイアウト
        main_layout = QVBoxLayout(self)
        
        # スクロールエリアの設定
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # スクロールエリア内のコンテンツウィジェット
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(10)
        
        # フォントサイズ設定グループ
        font_size_group = QGroupBox("フォントサイズ設定")
        font_size_layout = QVBoxLayout()
        
        # フォントサイズスライダー
        font_size_slider_layout = QHBoxLayout()
        self.font_size_label = QLabel(f"フォントサイズ: {self.current_font_size}pt")
        font_size_slider_layout.addWidget(self.font_size_label)
        
        self.font_size_slider = QSlider(Qt.Orientation.Horizontal)
        self.font_size_slider.setMinimum(8)
        self.font_size_slider.setMaximum(24)
        self.font_size_slider.setValue(self.current_font_size)
        self.font_size_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.font_size_slider.setTickInterval(1)
        self.font_size_slider.valueChanged.connect(self.update_font_size_label)
        font_size_slider_layout.addWidget(self.font_size_slider)
        
        # フォントサイズリセットボタン
        self.font_size_reset_btn = QPushButton("デフォルトに戻す")
        self.font_size_reset_btn.clicked.connect(self.reset_font_size)
        font_size_slider_layout.addWidget(self.font_size_reset_btn)
        
        font_size_layout.addLayout(font_size_slider_layout)
        font_size_group.setLayout(font_size_layout)
        content_layout.addWidget(font_size_group)
        
        # ブラウザ設定グループ
        browser_group = QGroupBox("ブラウザ設定")
        browser_layout = QVBoxLayout()
        
        # ブラウザ設定の説明
        browser_description = QLabel("提供エリア検索時のブラウザ動作を設定します。")
        browser_description.setWordWrap(True)
        browser_layout.addWidget(browser_description)
        
        # ヘッドレスモード設定
        self.headless_checkbox = QCheckBox("ヘッドレスモード（ブラウザを表示しない）")
        self.headless_checkbox.setChecked(self.current_browser_settings.get("headless", False))
        self.headless_checkbox.setToolTip("有効にするとブラウザが画面に表示されなくなり、処理が軽くなります")
        browser_layout.addWidget(self.headless_checkbox)
        
        # 画像読み込み無効化設定
        self.disable_images_checkbox = QCheckBox("画像読み込みを無効化（高速化）")
        self.disable_images_checkbox.setChecked(self.current_browser_settings.get("disable_images", True))
        self.disable_images_checkbox.setToolTip("有効にすると画像の読み込みをスキップし、処理が大幅に軽くなります")
        browser_layout.addWidget(self.disable_images_checkbox)
        
        # ポップアップ表示設定
        self.popup_checkbox = QCheckBox("結果をポップアップで表示する")
        self.popup_checkbox.setChecked(self.current_browser_settings.get("show_popup", True))
        self.popup_checkbox.setToolTip("無効にすると提供判定結果のポップアップが表示されなくなります")
        browser_layout.addWidget(self.popup_checkbox)
        
        # ブラウザ自動終了設定
        self.auto_close_checkbox = QCheckBox("ブラウザを自動的に閉じる")
        self.auto_close_checkbox.setChecked(self.current_browser_settings.get("auto_close", False))
        self.auto_close_checkbox.setToolTip("有効にするとブラウザウィンドウが自動的に閉じられます")
        browser_layout.addWidget(self.auto_close_checkbox)
        
        # タイムアウト設定
        timeout_group = QGroupBox("タイムアウト設定")
        timeout_layout = QVBoxLayout()
        
        # ページ読み込みタイムアウト
        page_load_timeout_layout = QHBoxLayout()
        self.page_load_timeout_label = QLabel(f"ページ読み込みタイムアウト: {self.current_browser_settings.get('page_load_timeout', 30)}秒")
        page_load_timeout_layout.addWidget(self.page_load_timeout_label)
        
        self.page_load_timeout_slider = QSlider(Qt.Orientation.Horizontal)
        self.page_load_timeout_slider.setMinimum(10)
        self.page_load_timeout_slider.setMaximum(120)
        self.page_load_timeout_slider.setValue(self.current_browser_settings.get('page_load_timeout', 30))
        self.page_load_timeout_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.page_load_timeout_slider.setTickInterval(10)
        self.page_load_timeout_slider.valueChanged.connect(self.update_page_load_timeout_label)
        page_load_timeout_layout.addWidget(self.page_load_timeout_slider)
        
        timeout_layout.addLayout(page_load_timeout_layout)
        
        # スクリプトタイムアウト
        script_timeout_layout = QHBoxLayout()
        self.script_timeout_label = QLabel(f"スクリプトタイムアウト: {self.current_browser_settings.get('script_timeout', 30)}秒")
        script_timeout_layout.addWidget(self.script_timeout_label)
        
        self.script_timeout_slider = QSlider(Qt.Orientation.Horizontal)
        self.script_timeout_slider.setMinimum(10)
        self.script_timeout_slider.setMaximum(120)
        self.script_timeout_slider.setValue(self.current_browser_settings.get('script_timeout', 30))
        self.script_timeout_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.script_timeout_slider.setTickInterval(10)
        self.script_timeout_slider.valueChanged.connect(self.update_script_timeout_label)
        script_timeout_layout.addWidget(self.script_timeout_slider)
        
        timeout_layout.addLayout(script_timeout_layout)
        timeout_group.setLayout(timeout_layout)
        browser_layout.addWidget(timeout_group)
        
        browser_group.setLayout(browser_layout)
        content_layout.addWidget(browser_group)
        
        # スクロールエリアにコンテンツを設定
        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        # OKボタン
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)
        
        # キャンセルボタン
        cancel_button = QPushButton("キャンセル")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)
        
        main_layout.addLayout(button_layout)
    
    def update_font_size_label(self, value):
        """フォントサイズラベルを更新"""
        self.font_size_label.setText(f"フォントサイズ: {value}pt")
    
    def reset_font_size(self):
        """フォントサイズをデフォルトに戻す"""
        self.font_size_slider.setValue(self.default_font_size)
    
    def update_page_load_timeout_label(self, value):
        """ページ読み込みタイムアウトラベルを更新"""
        self.page_load_timeout_label.setText(f"ページ読み込みタイムアウト: {value}秒")
    
    def update_script_timeout_label(self, value):
        """スクリプトタイムアウトラベルを更新"""
        self.script_timeout_label.setText(f"スクリプトタイムアウト: {value}秒")
    
    def set_settings(self, settings):
        """
        設定を適用する
        
        Args:
            settings (dict): 設定データ
        """
        try:
            # フォントサイズ設定
            font_size = settings.get('font_size', self.default_font_size)
            self.font_size_slider.setValue(font_size)
            
            # ブラウザ設定
            browser_settings = settings.get('browser_settings', self.default_browser_settings)
            
            # チェックボックスの設定
            self.headless_checkbox.setChecked(browser_settings.get('headless', False))
            self.disable_images_checkbox.setChecked(browser_settings.get('disable_images', True))
            self.popup_checkbox.setChecked(browser_settings.get('show_popup', True))
            self.auto_close_checkbox.setChecked(browser_settings.get('auto_close', False))
            
            # タイムアウト設定
            self.page_load_timeout_slider.setValue(browser_settings.get('page_load_timeout', 30))
            self.script_timeout_slider.setValue(browser_settings.get('script_timeout', 30))
            
            # 現在の設定を更新
            self.current_font_size = font_size
            self.current_browser_settings = browser_settings
            
        except Exception as e:
            logging.error(f"設定の適用中にエラー: {str(e)}")
            QMessageBox.warning(self, "エラー", f"設定の適用中にエラーが発生しました: {str(e)}")
    
    def get_settings(self):
        """
        現在の設定を取得
        
        Returns:
            dict: 設定データ
        """
        return {
            'font_size': self.font_size_slider.value(),
            'browser_settings': {
                'headless': self.headless_checkbox.isChecked(),
                'disable_images': self.disable_images_checkbox.isChecked(),
                'show_popup': self.popup_checkbox.isChecked(),
                'auto_close': self.auto_close_checkbox.isChecked(),
                'page_load_timeout': self.page_load_timeout_slider.value(),
                'script_timeout': self.script_timeout_slider.value()
            }
        } 