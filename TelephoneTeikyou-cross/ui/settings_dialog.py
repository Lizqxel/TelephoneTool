"""
設定ダイアログ

このモジュールは、アプリケーションの設定を管理するための
ダイアログUIを提供します。
"""

import json
import os
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                              QPushButton, QMessageBox, QSlider,
                              QGroupBox, QCheckBox, QScrollArea, QWidget,
                              QTabWidget, QSpinBox, QDoubleSpinBox)
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
        self.setMinimumWidth(400)
        
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
            "auto_close": False,
            "use_external_browser": False
        }
        
        # 親ウィンドウから現在の設定を取得
        if parent and hasattr(parent, 'settings'):
            self.current_font_size = parent.settings.get('font_size', self.default_font_size)
            self.current_browser_settings = parent.settings.get('browser_settings', self.default_browser_settings)
        else:
            self.current_font_size = self.default_font_size
            self.current_browser_settings = self.default_browser_settings
        
        # メインレイアウト
        layout = QVBoxLayout()
        
        # タブウィジェット
        tab_widget = QTabWidget()
        
        # 一般設定タブ
        general_tab = QWidget()
        general_layout = QVBoxLayout()
        
        # フォントサイズ設定
        font_group = QGroupBox("フォントサイズ")
        font_layout = QHBoxLayout()
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 24)
        font_layout.addWidget(QLabel("フォントサイズ:"))
        font_layout.addWidget(self.font_size_spin)
        font_group.setLayout(font_layout)
        general_layout.addWidget(font_group)
        
        general_tab.setLayout(general_layout)
        tab_widget.addTab(general_tab, "一般")
        
        # CTI監視設定タブ
        cti_tab = QWidget()
        cti_layout = QVBoxLayout()
        
        # CTI監視設定グループ
        cti_group = QGroupBox("CTI監視設定")
        cti_group_layout = QVBoxLayout()
        
        # CTI監視の有効/無効
        self.enable_cti_check = QCheckBox("CTI監視を有効にする")
        cti_group_layout.addWidget(self.enable_cti_check)
        
        # CTI自動処理の有効/無効
        self.enable_auto_cti_check = QCheckBox("CTI自動処理を有効にする")
        cti_group_layout.addWidget(self.enable_auto_cti_check)
        
        # 監視間隔設定
        interval_layout = QHBoxLayout()
        self.cti_interval_spin = QDoubleSpinBox()
        self.cti_interval_spin.setRange(0.1, 1.0)
        self.cti_interval_spin.setSingleStep(0.1)
        self.cti_interval_spin.setDecimals(1)
        interval_layout.addWidget(QLabel("監視間隔（秒）:"))
        interval_layout.addWidget(self.cti_interval_spin)
        cti_group_layout.addLayout(interval_layout)
        
        # 通話時間閾値設定
        call_duration_layout = QHBoxLayout()
        self.call_duration_spin = QSpinBox()
        self.call_duration_spin.setRange(0, 60)
        self.call_duration_spin.setSuffix(" 秒")
        self.call_duration_spin.setSpecialValueText("即座に実行")
        call_duration_layout.addWidget(QLabel("自動処理開始までの通話時間:"))
        call_duration_layout.addWidget(self.call_duration_spin)
        cti_group_layout.addLayout(call_duration_layout)
        
        # 説明ラベル
        help_label = QLabel("※ 0秒の場合は通話開始と同時に自動処理を実行します")
        help_label.setStyleSheet("color: #666666; font-size: 10px;")
        cti_group_layout.addWidget(help_label)
        
        cti_group.setLayout(cti_group_layout)
        cti_layout.addWidget(cti_group)
        
        cti_tab.setLayout(cti_layout)
        tab_widget.addTab(cti_tab, "CTI監視")
        
        # ブラウザ設定タブ
        browser_tab = QWidget()
        browser_layout = QVBoxLayout()
        
        # ブラウザ設定グループ
        browser_group = QGroupBox("ブラウザ設定")
        browser_group_layout = QVBoxLayout()
        
        # ヘッドレスモード
        self.headless_check = QCheckBox("ヘッドレスモードを使用")
        browser_group_layout.addWidget(self.headless_check)
        
        # 画像読み込みの無効化
        self.disable_images_check = QCheckBox("画像読み込みを無効化")
        browser_group_layout.addWidget(self.disable_images_check)
        
        # ポップアップ表示
        self.show_popup_check = QCheckBox("検索結果のポップアップを表示")
        browser_group_layout.addWidget(self.show_popup_check)
        
        # 自動クローズ
        self.auto_close_check = QCheckBox("検索完了後にブラウザを自動クローズ")
        browser_group_layout.addWidget(self.auto_close_check)
        
        # 外部ブラウザ使用
        self.use_external_browser_check = QCheckBox("電話番号検索を外部ブラウザで開く（ノートパソコン用推奨設定）")
        browser_group_layout.addWidget(self.use_external_browser_check)
        
        # タイムアウト設定
        timeout_layout = QHBoxLayout()
        self.page_timeout_spin = QSpinBox()
        self.page_timeout_spin.setRange(30, 180)
        timeout_layout.addWidget(QLabel("ページ読み込みタイムアウト（秒）:"))
        timeout_layout.addWidget(self.page_timeout_spin)
        browser_group_layout.addLayout(timeout_layout)
        
        script_timeout_layout = QHBoxLayout()
        self.script_timeout_spin = QSpinBox()
        self.script_timeout_spin.setRange(30, 180)
        script_timeout_layout.addWidget(QLabel("スクリプト実行タイムアウト（秒）:"))
        script_timeout_layout.addWidget(self.script_timeout_spin)
        browser_group_layout.addLayout(script_timeout_layout)
        
        browser_group.setLayout(browser_group_layout)
        browser_layout.addWidget(browser_group)
        
        browser_tab.setLayout(browser_layout)
        tab_widget.addTab(browser_tab, "ブラウザ")
        
        layout.addWidget(tab_widget)
        
        # OKとキャンセルボタン
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("キャンセル")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def set_settings(self, settings):
        """設定を読み込んでUIに反映"""
        # フォントサイズ
        self.font_size_spin.setValue(settings.get('font_size', self.default_font_size))
        
        # CTI監視設定
        cti_settings = settings.get('cti_settings', {})
        self.enable_cti_check.setChecked(cti_settings.get('enable_cti', True))
        self.enable_auto_cti_check.setChecked(cti_settings.get('enable_auto_cti_processing', True))
        self.cti_interval_spin.setValue(cti_settings.get('cti_monitor_interval', 0.2))
        
        # 通話時間閾値設定（メイン設定から取得）
        self.call_duration_spin.setValue(settings.get('call_duration_threshold', 0))
        
        # ブラウザ設定
        browser_settings = settings.get('browser_settings', self.default_browser_settings)
        self.headless_check.setChecked(browser_settings.get('headless', False))
        self.disable_images_check.setChecked(browser_settings.get('disable_images', True))
        self.show_popup_check.setChecked(browser_settings.get('show_popup', True))
        self.auto_close_check.setChecked(browser_settings.get('auto_close', False))
        self.use_external_browser_check.setChecked(browser_settings.get('use_external_browser', False))
        self.page_timeout_spin.setValue(browser_settings.get('page_load_timeout', 30))
        self.script_timeout_spin.setValue(browser_settings.get('script_timeout', 30))
    
    def get_settings(self):
        """UIの設定を取得"""
        return {
            'font_size': self.font_size_spin.value(),
            'call_duration_threshold': self.call_duration_spin.value(),
            'cti_settings': {
                'enable_cti': self.enable_cti_check.isChecked(),
                'enable_auto_cti_processing': self.enable_auto_cti_check.isChecked(),
                'cti_monitor_interval': self.cti_interval_spin.value()
            },
            'browser_settings': {
                'headless': self.headless_check.isChecked(),
                'disable_images': self.disable_images_check.isChecked(),
                'show_popup': self.show_popup_check.isChecked(),
                'auto_close': self.auto_close_check.isChecked(),
                'use_external_browser': self.use_external_browser_check.isChecked(),
                'page_load_timeout': self.page_timeout_spin.value(),
                'script_timeout': self.script_timeout_spin.value()
            }
        } 