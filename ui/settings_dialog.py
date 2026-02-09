"""
設定ダイアログ

このモジュールは、アプリケーションの設定を管理するための
ダイアログUIを提供します。
"""

import json
import os
import sys
import logging
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                              QTextEdit, QPushButton, QMessageBox, QSlider,
                              QGroupBox, QSpinBox, QCheckBox, QScrollArea, QWidget,
                              QRadioButton, QLineEdit, QTableWidget, QTableWidgetItem)
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHeaderView


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
        self.setFixedSize(700, 600)  # ダイアログサイズを固定
        
        # 設定ファイルのパスを絶対パスで設定
        if getattr(sys, 'frozen', False):
            # exeファイルとして実行されている場合
            self.settings_file = os.path.join(os.path.dirname(sys.executable), 'settings.json')
        else:
            # 通常のPythonスクリプトとして実行されている場合
            self.settings_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
        
        logging.info(f"設定ファイルのパス: {self.settings_file}")
        
        # デフォルトのフォーマットテンプレート
        self.default_format = """対応者（お客様の名前）：{operator}
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
案内料金：2500円"""

        # デフォルトのフォントサイズ
        self.default_font_size = 9
        
        # デフォルトの遅延時間（秒）
        self.default_delay = 0
        
        # デフォルトのブラウザ設定
        self.default_browser_settings = {
            "headless": False,
            "disable_images": True,
            "enable_screenshots": True,
            "show_popup": True,
            "page_load_timeout": 30,
            "script_timeout": 30,
            "auto_close": False
        }

        # 法人モード設定のデフォルト
        self.default_corporate_settings = {
            "allow_manual_contractor": False,
            "auto_copy_operator_to_contractor": True
        }
        
        # 親ウィンドウから現在の設定を取得
        if parent and hasattr(parent, 'settings'):
            self.current_font_size = parent.settings.get('font_size', self.default_font_size)
            self.current_delay = parent.settings.get('delay_seconds', self.default_delay)
            self.current_browser_settings = parent.settings.get('browser_settings', self.default_browser_settings)
            self.current_corporate_settings = parent.settings.get('corporate_settings', self.default_corporate_settings)
            current_call_duration = parent.settings.get('call_duration_threshold', 0)  # デフォルトは0秒
        else:
            self.current_font_size = self.default_font_size
            self.current_delay = self.default_delay
            self.current_browser_settings = self.default_browser_settings
            self.current_corporate_settings = self.default_corporate_settings
            current_call_duration = 0
        
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
        
        # 遅延時間設定グループ
        delay_group = QGroupBox("遅延時間設定")
        delay_layout = QVBoxLayout()
        
        # 遅延時間の説明
        delay_description = QLabel("CTIボタンクリック後、情報取得を開始するまでの遅延時間を設定します。")
        delay_description.setWordWrap(True)
        delay_layout.addWidget(delay_description)
        
        # 遅延時間スピンボックス
        delay_spin_layout = QHBoxLayout()
        delay_spin_layout.addWidget(QLabel("遅延時間:"))
        
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(0, 300)  # 0-300秒
        self.delay_spin.setValue(self.current_delay)
        self.delay_spin.setSuffix(" 秒")
        delay_spin_layout.addWidget(self.delay_spin)
        
        # 遅延時間リセットボタン
        self.delay_reset_btn = QPushButton("デフォルトに戻す")
        self.delay_reset_btn.clicked.connect(self.reset_delay)
        delay_spin_layout.addWidget(self.delay_reset_btn)
        
        delay_layout.addLayout(delay_spin_layout)
        delay_group.setLayout(delay_layout)
        content_layout.addWidget(delay_group)
        
        # モード選択グループ
        mode_group = QGroupBox("モード選択")
        mode_layout = QVBoxLayout()
        
        # モード選択の説明
        mode_description = QLabel("アプリケーションの動作モードを切り替えます。")
        mode_description.setWordWrap(True)
        mode_layout.addWidget(mode_description)
        
        # モード選択レイアウト
        mode_select_layout = QHBoxLayout()
        
        # 通常モードラジオボタン
        self.simple_mode_radio = QRadioButton("通常モード")
        self.simple_mode_radio.setToolTip("機能をすべて表示するモード")
        mode_select_layout.addWidget(self.simple_mode_radio)
        
        # 法人モードラジオボタン
        self.corporate_mode_radio = QRadioButton("法人モード")
        self.corporate_mode_radio.setToolTip("法人向けの入力補助を追加するモード")
        mode_select_layout.addWidget(self.corporate_mode_radio)
        
        # 現在のモードを設定
        if hasattr(parent, 'current_mode'):
            if parent.current_mode == 'simple':
                self.simple_mode_radio.setChecked(True)
            elif parent.current_mode == 'corporate':
                self.corporate_mode_radio.setChecked(True)
            else:
                self.simple_mode_radio.setChecked(True)
        else:
            self.simple_mode_radio.setChecked(True)  # デフォルトは通常モード
        
        mode_layout.addLayout(mode_select_layout)
        mode_group.setLayout(mode_layout)
        content_layout.addWidget(mode_group)

        # 法人モード設定グループ
        corporate_group = QGroupBox("法人モード設定")
        corporate_layout = QVBoxLayout()

        corporate_desc = QLabel("法人モードでの入力補助の設定です。")
        corporate_desc.setWordWrap(True)
        corporate_layout.addWidget(corporate_desc)

        self.corporate_auto_copy_checkbox = QCheckBox("同期して自動入力する（対応者名・契約者名・フリガナ）")
        self.corporate_auto_copy_checkbox.setChecked(
            self.current_corporate_settings.get("auto_copy_operator_to_contractor", True)
        )
        corporate_layout.addWidget(self.corporate_auto_copy_checkbox)

        corporate_group.setLayout(corporate_layout)
        content_layout.addWidget(corporate_group)
        
        # CTI監視設定グループ
        cti_monitor_group = QGroupBox("CTI監視設定")
        cti_monitor_layout = QVBoxLayout()
        
        # CTI監視設定の説明
        cti_monitor_description = QLabel("CTI状態変化の監視と自動処理の設定を行います。")
        cti_monitor_description.setWordWrap(True)
        cti_monitor_layout.addWidget(cti_monitor_description)
        
        # CTI監視有効/無効設定
        self.cti_monitoring_checkbox = QCheckBox("CTI状態監視を有効にする")
        self.cti_monitoring_checkbox.setToolTip("有効にするとCTI状態の変化を監視し、発信中から通話中への変化時に自動で顧客情報取得と提供判定を実行します")
        
        # 現在のCTI監視設定を読み込み
        if hasattr(parent, 'settings'):
            current_cti_enabled = parent.settings.get('enable_cti_monitoring', True)
            current_cti_interval = parent.settings.get('cti_monitor_interval', 0.2)
            current_cti_cooldown = parent.settings.get('cti_auto_processing_cooldown', 3.0)
            current_call_duration = parent.settings.get('call_duration_threshold', 0)  # デフォルトは0秒
        else:
            current_cti_enabled = True
            current_cti_interval = 0.2
            current_cti_cooldown = 3.0
            current_call_duration = 0
            
        self.cti_monitoring_checkbox.setChecked(current_cti_enabled)
        cti_monitor_layout.addWidget(self.cti_monitoring_checkbox)
        
        # CTI自動処理有効/無効設定
        self.cti_auto_processing_checkbox = QCheckBox("CTI自動処理を有効にする")
        self.cti_auto_processing_checkbox.setToolTip("有効にするとCTI状態変化時に自動で顧客情報取得と提供判定を実行します")
        
        # 現在のCTI自動処理設定を読み込み
        if hasattr(parent, 'settings'):
            current_auto_processing = parent.settings.get('enable_auto_cti_processing', True)
        else:
            current_auto_processing = True
            
        self.cti_auto_processing_checkbox.setChecked(current_auto_processing)
        cti_monitor_layout.addWidget(self.cti_auto_processing_checkbox)
        
        # CTI監視間隔設定
        cti_interval_layout = QHBoxLayout()
        cti_interval_layout.addWidget(QLabel("監視間隔:"))
        
        self.cti_interval_spin = QSpinBox()
        self.cti_interval_spin.setRange(100, 2000)  # 100ms-2000ms
        self.cti_interval_spin.setValue(int(current_cti_interval * 1000))  # 秒をミリ秒に変換
        self.cti_interval_spin.setSuffix(" ms")
        self.cti_interval_spin.setToolTip("CTI状態をチェックする間隔です。短いほど反応が良くなりますが、CPU負荷が高くなります")
        cti_interval_layout.addWidget(self.cti_interval_spin)
        
        cti_monitor_layout.addLayout(cti_interval_layout)
        
        # CTI自動処理クールダウン時間設定
        cti_cooldown_layout = QHBoxLayout()
        cti_cooldown_layout.addWidget(QLabel("自動処理クールダウン:"))
        
        self.cti_cooldown_spin = QSpinBox()
        self.cti_cooldown_spin.setRange(1, 30)  # 1-30秒
        self.cti_cooldown_spin.setValue(int(current_cti_cooldown))
        self.cti_cooldown_spin.setSuffix(" 秒")
        self.cti_cooldown_spin.setToolTip("同じ状態変化の重複実行を防ぐための最小間隔です")
        cti_cooldown_layout.addWidget(self.cti_cooldown_spin)
        
        cti_monitor_layout.addLayout(cti_cooldown_layout)
        
        # 通話時間設定
        call_duration_layout = QHBoxLayout()
        call_duration_layout.addWidget(QLabel("通話時間設定:"))
        
        self.call_duration_spin = QSpinBox()
        self.call_duration_spin.setRange(0, 300)  # 0-300秒
        self.call_duration_spin.setValue(int(current_call_duration))
        self.call_duration_spin.setSuffix(" 秒")
        self.call_duration_spin.setToolTip("「発信中」→「通話中」に変化した時に、この時間以上経過した場合に自動実行します。0秒の場合は即時実行します。")
        call_duration_layout.addWidget(self.call_duration_spin)
        
        cti_monitor_layout.addLayout(call_duration_layout)
        
        # CTI設定リセットボタン
        cti_reset_layout = QHBoxLayout()
        cti_reset_layout.addStretch()
        
        self.cti_reset_btn = QPushButton("CTI設定をデフォルトに戻す")
        self.cti_reset_btn.clicked.connect(self.reset_cti_settings)
        cti_reset_layout.addWidget(self.cti_reset_btn)
        
        cti_monitor_layout.addLayout(cti_reset_layout)
        cti_monitor_group.setLayout(cti_monitor_layout)
        content_layout.addWidget(cti_monitor_group)
        
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
        
        # スクリーンショット有効化設定
        self.screenshots_checkbox = QCheckBox("提供判定時のスクリーンショットを有効にする")
        self.screenshots_checkbox.setChecked(self.current_browser_settings.get("enable_screenshots", True))
        self.screenshots_checkbox.setToolTip("無効にすると提供判定時のスクリーンショット保存を行いません")
        browser_layout.addWidget(self.screenshots_checkbox)
        
        # ブラウザ自動終了設定
        self.auto_close_checkbox = QCheckBox("ブラウザを自動的に閉じる")
        self.auto_close_checkbox.setChecked(self.current_browser_settings.get("auto_close", False))
        self.auto_close_checkbox.setToolTip("有効にするとブラウザウィンドウが自動的に閉じられます")
        browser_layout.addWidget(self.auto_close_checkbox)
        
        # タイムアウト設定
        timeout_layout = QHBoxLayout()
        timeout_layout.addWidget(QLabel("ページ読み込みタイムアウト:"))
        
        self.page_timeout_spin = QSpinBox()
        self.page_timeout_spin.setRange(10, 120)  # 10-120秒
        self.page_timeout_spin.setValue(self.current_browser_settings.get("page_load_timeout", 30))
        self.page_timeout_spin.setSuffix(" 秒")
        timeout_layout.addWidget(self.page_timeout_spin)
        
        timeout_layout.addWidget(QLabel("スクリプトタイムアウト:"))
        
        self.script_timeout_spin = QSpinBox()
        self.script_timeout_spin.setRange(10, 120)  # 10-120秒
        self.script_timeout_spin.setValue(self.current_browser_settings.get("script_timeout", 30))
        self.script_timeout_spin.setSuffix(" 秒")
        timeout_layout.addWidget(self.script_timeout_spin)
        
        browser_layout.addLayout(timeout_layout)
        
        # ブラウザ設定リセットボタン
        browser_reset_layout = QHBoxLayout()
        browser_reset_layout.addStretch()
        
        self.browser_reset_btn = QPushButton("ブラウザ設定をデフォルトに戻す")
        self.browser_reset_btn.clicked.connect(self.reset_browser_settings)
        browser_reset_layout.addWidget(self.browser_reset_btn)
        
        browser_layout.addLayout(browser_reset_layout)
        browser_group.setLayout(browser_layout)
        content_layout.addWidget(browser_group)

        # 共有トークン設定グループ
        token_group = QGroupBox("共有トークン")
        token_layout = QVBoxLayout()
        token_desc = QLabel("Googleフォーム送信の認証用トークンです。入力はマスクされます。\n空欄のまま保存はできません。")
        token_desc.setWordWrap(True)
        token_layout.addWidget(token_desc)

        row = QHBoxLayout()
        row.addWidget(QLabel("共有トークン："))
        self.token_edit = QLineEdit()
        self.token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.token_edit.setPlaceholderText("共有トークンを入力")
        row.addWidget(self.token_edit)
        token_layout.addLayout(row)

        token_group.setLayout(token_layout)
        content_layout.addWidget(token_group)

        # 転記先（Googleフォーム送信）設定グループ
        dest_group = QGroupBox("転記先（Googleフォーム送信）")
        dest_layout = QVBoxLayout()

        dest_desc = QLabel("settings.json の googleFormPosting.destinations を編集できます。\n表示名・スプレッドシートURL・シート名を管理します。")
        dest_desc.setWordWrap(True)
        dest_layout.addWidget(dest_desc)

        # 転記先テーブル
        self.dest_table = QTableWidget(0, 3, self)
        self.dest_table.setHorizontalHeaderLabels(["表示名", "スプシURL", "シート名"])
        header = self.dest_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.dest_table.setSelectionBehavior(self.dest_table.SelectionBehavior.SelectRows)
        self.dest_table.setSelectionMode(self.dest_table.SelectionMode.SingleSelection)
        dest_layout.addWidget(self.dest_table)

        # ボタン群
        dest_btns = QHBoxLayout()
        self.dest_add_btn = QPushButton("追加")
        self.dest_edit_btn = QPushButton("編集")
        self.dest_del_btn = QPushButton("削除")
        dest_btns.addStretch()
        dest_btns.addWidget(self.dest_add_btn)
        dest_btns.addWidget(self.dest_edit_btn)
        dest_btns.addWidget(self.dest_del_btn)
        dest_layout.addLayout(dest_btns)

        dest_group.setLayout(dest_layout)
        content_layout.addWidget(dest_group)
        
        # CTIフォーマットグループ
        cti_format_group = QGroupBox("CTIフォーマットテンプレート")
        cti_format_layout = QVBoxLayout()
        
        # 説明ラベル
        description = QLabel("CTIフォーマットのテンプレートを編集できます。\n"
                            "以下のプレースホルダーが使用可能です：\n"
                            "{operator}, {available_time}, {mobile}, {stakeholder}, {contractor}, {furigana}, {birth_date}, {postal_code}, {address}, "
                            "{list_name}, {list_furigana}, {list_phone}, {list_postal_code}, {list_address}, "
                            "{current_line}, {order_date}, {order_person}, {judgment}, {fee}, {net_usage}, {family_approval}, {remarks}, "
                            "{other_number}, {phone_device}, {forbidden_line}, {nd}")
        description.setWordWrap(True)
        cti_format_layout.addWidget(description)
        
        # テキスト編集エリア
        self.format_edit = QTextEdit()
        self.format_edit.setPlaceholderText("フォーマットテンプレートを入力してください")
        self.format_edit.setMinimumHeight(300)  # テキスト編集エリアの高さを増やす
        cti_format_layout.addWidget(self.format_edit)
        
        cti_format_group.setLayout(cti_format_layout)
        content_layout.addWidget(cti_format_group)
        
        # スクロールエリアにウィジェットを設定
        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area, 1)  # 1は伸縮比率
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        # リセットボタン
        self.reset_btn = QPushButton("デフォルトに戻す")
        self.reset_btn.clicked.connect(self.reset_to_default)
        button_layout.addWidget(self.reset_btn)
        
        # スペーサー
        button_layout.addStretch()
        
        # キャンセルボタン
        self.cancel_btn = QPushButton("キャンセル")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        # 保存ボタン
        self.save_btn = QPushButton("保存")
        self.save_btn.clicked.connect(self.accept)
        self.save_btn.setDefault(True)
        button_layout.addWidget(self.save_btn)
        
        main_layout.addLayout(button_layout)
        
        # 設定の読み込み
        self.load_settings()

        # 転記先ボタンのハンドラ
        self.dest_add_btn.clicked.connect(self._on_add_destination)
        self.dest_edit_btn.clicked.connect(self._on_edit_destination)
        self.dest_del_btn.clicked.connect(self._on_delete_destination)
    
    def update_font_size_label(self, value):
        """フォントサイズラベルを更新する"""
        self.font_size_label.setText(f"フォントサイズ: {value}pt")
    
    def reset_font_size(self):
        """フォントサイズをデフォルトに戻す"""
        self.font_size_slider.setValue(self.default_font_size)
    
    def reset_delay(self):
        """遅延時間をデフォルトに戻す"""
        self.delay_spin.setValue(self.default_delay)
    
    def reset_browser_settings(self):
        """ブラウザ設定をデフォルトに戻す"""
        self.headless_checkbox.setChecked(self.default_browser_settings["headless"])
        self.disable_images_checkbox.setChecked(self.default_browser_settings["disable_images"])
        # enable_screenshots をデフォルトに戻す
        self.screenshots_checkbox.setChecked(self.default_browser_settings.get("enable_screenshots", True))
        self.popup_checkbox.setChecked(self.default_browser_settings["show_popup"])
        self.auto_close_checkbox.setChecked(self.default_browser_settings["auto_close"])
        self.page_timeout_spin.setValue(self.default_browser_settings["page_load_timeout"])
        self.script_timeout_spin.setValue(self.default_browser_settings["script_timeout"])
    
    def reset_cti_settings(self):
        """CTI設定をデフォルトに戻す"""
        self.cti_monitoring_checkbox.setChecked(True)
        self.cti_auto_processing_checkbox.setChecked(True)
        self.cti_interval_spin.setValue(200)  # 0.2秒
        self.cti_cooldown_spin.setValue(3)  # 3秒
        self.call_duration_spin.setValue(0)  # 0秒
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    # フル設定を保持（保存時に他キーを保持するため）
                    self._full_settings = dict(settings)
                    format_template = settings.get('format_template', self.default_format)
                    font_size = settings.get('font_size', self.default_font_size)
                    delay_seconds = settings.get('delay_seconds', self.default_delay)
                    browser_settings = settings.get('browser_settings', self.default_browser_settings)
                    mode = settings.get('mode', 'simple')
                    
                    # CTI監視設定の読み込み
                    cti_monitoring_enabled = settings.get('enable_cti_monitoring', True)
                    cti_auto_processing_enabled = settings.get('enable_auto_cti_processing', True)
                    cti_monitor_interval = settings.get('cti_monitor_interval', 0.2)
                    cti_auto_processing_cooldown = settings.get('cti_auto_processing_cooldown', 3.0)
                    current_call_duration = settings.get('call_duration_threshold', 0)  # デフォルトは0秒
                    
                    self.format_edit.setText(format_template)
                    self.font_size_slider.setValue(font_size)
                    self.delay_spin.setValue(delay_seconds)
                    
                    # CTI監視設定の設定
                    self.cti_monitoring_checkbox.setChecked(cti_monitoring_enabled)
                    self.cti_auto_processing_checkbox.setChecked(cti_auto_processing_enabled)
                    self.cti_interval_spin.setValue(int(cti_monitor_interval * 1000))  # 秒をミリ秒に変換
                    self.cti_cooldown_spin.setValue(int(cti_auto_processing_cooldown))
                    
                    # モード設定の読み込み
                    if mode == 'corporate':
                        self.corporate_mode_radio.setChecked(True)
                    else:
                        self.simple_mode_radio.setChecked(True)

                    # 法人モード設定の読み込み
                    corporate_settings = settings.get('corporate_settings', self.default_corporate_settings)
                    if isinstance(corporate_settings, dict):
                        if "auto_copy_operator_to_contractor" in corporate_settings:
                            auto_copy = corporate_settings.get("auto_copy_operator_to_contractor", True)
                        else:
                            allow_manual = corporate_settings.get("allow_manual_contractor", False)
                            auto_copy = not allow_manual
                        self.corporate_auto_copy_checkbox.setChecked(bool(auto_copy))
                    
                    # ブラウザ設定の読み込み
                    self.headless_checkbox.setChecked(browser_settings.get("headless", False))
                    self.disable_images_checkbox.setChecked(browser_settings.get("disable_images", True))
                    self.screenshots_checkbox.setChecked(browser_settings.get("enable_screenshots", True))
                    self.popup_checkbox.setChecked(browser_settings.get("show_popup", True))
                    self.auto_close_checkbox.setChecked(browser_settings.get("auto_close", False))
                    self.page_timeout_spin.setValue(browser_settings.get("page_load_timeout", 30))
                    self.script_timeout_spin.setValue(browser_settings.get("script_timeout", 30))

                    # 転記先の読み込み
                    gfp = settings.get('googleFormPosting', {}) if isinstance(settings, dict) else {}
                    dests = gfp.get('destinations', []) if isinstance(gfp, dict) else []
                    self._populate_dest_table(dests)

                    # 共有トークンの読み込み
                    try:
                        token_val = str(gfp.get('tokenValue', '')) if isinstance(gfp, dict) else ''
                        self.token_edit.setText(token_val)
                    except Exception:
                        self.token_edit.setText('')
            else:
                self.format_edit.setText(self.default_format)
                self.font_size_slider.setValue(self.default_font_size)
                self.delay_spin.setValue(self.default_delay)
                self.simple_mode_radio.setChecked(True)  # デフォルトは通常モード
                self.corporate_auto_copy_checkbox.setChecked(
                    self.default_corporate_settings.get("auto_copy_operator_to_contractor", True)
                )
                self.reset_browser_settings()
                self.reset_cti_settings()  # CTI設定もデフォルトに
                self._full_settings = {}
                self._populate_dest_table([])
                self.token_edit.setText('')
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
            self.format_edit.setText(self.default_format)
            self.font_size_slider.setValue(self.default_font_size)
            self.delay_spin.setValue(self.default_delay)
            self.simple_mode_radio.setChecked(True)  # デフォルトは通常モード
            self.corporate_auto_copy_checkbox.setChecked(
                self.default_corporate_settings.get("auto_copy_operator_to_contractor", True)
            )
            self.reset_browser_settings()
            self.reset_cti_settings()  # CTI設定もデフォルトに
            self._full_settings = {}
            self._populate_dest_table([])
            self.token_edit.setText('')
    
    def save_settings(self):
        """設定をファイルに保存する"""
        try:
            # 現在のモードを取得
            if hasattr(self.parent(), 'current_mode'):
                previous_mode = self.parent().current_mode
            else:
                previous_mode = 'simple'
            
            # 選択されたモードを取得
            if self.corporate_mode_radio.isChecked():
                new_mode = 'corporate'
            else:
                new_mode = 'simple'
            
            # ブラウザ設定を取得
            browser_settings = {
                "headless": self.headless_checkbox.isChecked(),
                "disable_images": self.disable_images_checkbox.isChecked(),
                "enable_screenshots": self.screenshots_checkbox.isChecked(),
                "show_popup": self.popup_checkbox.isChecked(),
                "auto_close": self.auto_close_checkbox.isChecked(),
                "page_load_timeout": self.page_timeout_spin.value(),
                "script_timeout": self.script_timeout_spin.value()
            }
            
            # 既存設定をベースに上書き（未知のキーを保持）
            base = {}
            try:
                if os.path.exists(self.settings_file):
                    with open(self.settings_file, 'r', encoding='utf-8') as rf:
                        base = json.load(rf) or {}
            except Exception:
                base = {}

            auto_copy = self.corporate_auto_copy_checkbox.isChecked()
            settings = {
                'format_template': self.format_edit.toPlainText(),
                'font_size': self.font_size_slider.value(),
                'delay_seconds': self.delay_spin.value(),
                'browser_settings': browser_settings,
                'mode': new_mode,
                'show_mode_selection': False,  # モード選択ダイアログを次回から表示しない
                'corporate_settings': {
                    'allow_manual_contractor': not auto_copy,
                    'auto_copy_operator_to_contractor': auto_copy
                },
                # CTI監視設定を追加
                'enable_cti_monitoring': self.cti_monitoring_checkbox.isChecked(),
                'enable_auto_cti_processing': self.cti_auto_processing_checkbox.isChecked(),
                'cti_monitor_interval': self.cti_interval_spin.value() / 1000.0,  # ミリ秒を秒に変換
                'cti_auto_processing_cooldown': float(self.cti_cooldown_spin.value()),
                'call_duration_threshold': self.call_duration_spin.value()
            }

            # googleFormPosting を反映（destinations と tokenValue）
            gfp = base.get('googleFormPosting', {}) if isinstance(base, dict) else {}
            if not isinstance(gfp, dict):
                gfp = {}
            gfp['destinations'] = self._collect_dest_table()

            # 共有トークンの反映（空文字も許容）。
            # ユーザーが空欄にした場合は空文字を保存し、既存の値を保持する挙動ではなく
            # 明示的に空文字を格納することで、settings.json の構造が常に存在するようにする。
            token_val = (self.token_edit.text() or '').strip()
            gfp['tokenValue'] = token_val

            # Fill missing googleFormPosting keys with sensible defaults so that
            # downstream code can assume their presence. These defaults mirror
            # the ones used in first_run_setup.ensure_settings_file().
            default_form_url = (
                "https://docs.google.com/forms/d/e/1FAIpQLSfoDjvD0gyYxqcmADbwBtJ9CSsVB7QAHX8gSRG9sLXzsSNXYQ/formResponse"
            )
            if not gfp.get('formUrl'):
                gfp['formUrl'] = default_form_url

            if not gfp.get('timezone'):
                gfp['timezone'] = 'Asia/Tokyo'

            if not isinstance(gfp.get('retryPolicy'), dict):
                gfp['retryPolicy'] = {"maxAttempts": 3, "backoffSeconds": [1, 3, 10]}

            if not isinstance(gfp.get('defaults'), dict):
                gfp['defaults'] = {
                    "kanKatsu": "岩田管轄",
                    "shozai": "NA光",
                    "kubun": "新規",
                    "zenkakuResult": "前確入力待ち"
                }

            if not isinstance(gfp.get('choices'), dict):
                gfp['choices'] = {}

            # entryMap default (same mapping as initial setup)
            if not isinstance(gfp.get('entryMap'), dict):
                gfp['entryMap'] = {
                    "kanKatsu": "entry.1653774096",
                    "kakutokuSha": "entry.990044023",
                    "kakutokuId": "entry.20053472",
                    "listName": "entry.311805903",
                    "shozai": "entry.444975123",
                    "kubun": "entry.858726565",
                    "kadenTime": "entry.1423382151",
                    "freeBox": "entry.1331784643",
                    "tosDate": "entry.1643496135",
                    "zenkakuCallDate": "entry.2129109581",
                    "zenkakuResult": "entry.1899186247",
                    "sharedToken": "entry.447700198",
                    "spreadsheetUrl": "entry.574347119",
                    "sheetName": "entry.556971462"
                }

            base['googleFormPosting'] = gfp

            # 上書きキーを設定
            base.update(settings)

            # 安全に保存（tmp→置換）
            tmp_path = self.settings_file + ".tmp"
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(base, f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, self.settings_file)
            
            # モードが変更された場合、UIを再構築
            if previous_mode != new_mode and hasattr(self.parent(), 'current_mode'):
                self.parent().current_mode = new_mode
                # モード変更フラグを設定
                self.parent().mode_changed = True
                self.parent().new_mode = new_mode
                
                # 親ウィンドウのUIを再構築
                if hasattr(self.parent(), 'reconstruct_ui'):
                    self.parent().reconstruct_ui()
            
            return True
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"設定の保存に失敗しました: {str(e)}")
            return False
    
    def reset_to_default(self):
        """設定をデフォルトに戻す"""
        self.format_edit.setText(self.default_format)
        self.font_size_slider.setValue(self.default_font_size)
        self.delay_spin.setValue(self.default_delay)
        self.simple_mode_radio.setChecked(True)  # デフォルトは通常モード
        self.corporate_auto_copy_checkbox.setChecked(
            self.default_corporate_settings.get("auto_copy_operator_to_contractor", True)
        )
        self.reset_browser_settings()
        self.reset_cti_settings()  # CTI設定もデフォルトに戻す
    
    def accept(self):
        """ダイアログを受け入れる（OKボタン）"""
        if self.save_settings():
            super().accept()
    
    def get_settings(self):
        """現在の設定を取得する"""
        browser_settings = {
            "headless": self.headless_checkbox.isChecked(),
            "disable_images": self.disable_images_checkbox.isChecked(),
            "enable_screenshots": self.screenshots_checkbox.isChecked(),
            "show_popup": self.popup_checkbox.isChecked(),
            "auto_close": self.auto_close_checkbox.isChecked(),
            "page_load_timeout": self.page_timeout_spin.value(),
            "script_timeout": self.script_timeout_spin.value()
        }
        
        # 選択されたモードを取得
        if self.corporate_mode_radio.isChecked():
            mode = 'corporate'
        else:
            mode = 'simple'
        
        auto_copy = self.corporate_auto_copy_checkbox.isChecked()
        return {
            'format_template': self.format_edit.toPlainText(),
            'font_size': self.font_size_slider.value(),
            'delay_seconds': self.delay_spin.value(),
            'browser_settings': browser_settings,
            'mode': mode,
            'show_mode_selection': False,  # モード選択ダイアログを次回から表示しない
            'corporate_settings': {
                'allow_manual_contractor': not auto_copy,
                'auto_copy_operator_to_contractor': auto_copy
            },
            # CTI監視設定を追加
            'enable_cti_monitoring': self.cti_monitoring_checkbox.isChecked(),
            'enable_auto_cti_processing': self.cti_auto_processing_checkbox.isChecked(),
            'cti_monitor_interval': self.cti_interval_spin.value() / 1000.0,  # ミリ秒を秒に変換
            'cti_auto_processing_cooldown': float(self.cti_cooldown_spin.value()),
            'call_duration_threshold': self.call_duration_spin.value()
        } 

    # ===== 転記先（destinations）編集機能 =====
    def _populate_dest_table(self, dests):
        try:
            self.dest_table.setRowCount(0)
            for d in dests or []:
                row = self.dest_table.rowCount()
                self.dest_table.insertRow(row)
                self.dest_table.setItem(row, 0, QTableWidgetItem(str(d.get('label', ''))))
                self.dest_table.setItem(row, 1, QTableWidgetItem(str(d.get('spreadsheetUrl', ''))))
                self.dest_table.setItem(row, 2, QTableWidgetItem(str(d.get('sheetName', ''))))
        except Exception:
            pass

    def _collect_dest_table(self):
        dests = []
        rows = self.dest_table.rowCount()
        for r in range(rows):
            label = (self.dest_table.item(r, 0).text() if self.dest_table.item(r, 0) else '').strip()
            url = (self.dest_table.item(r, 1).text() if self.dest_table.item(r, 1) else '').strip()
            sname = (self.dest_table.item(r, 2).text() if self.dest_table.item(r, 2) else '').strip()
            if label or url or sname:
                dests.append({
                    'label': label,
                    'spreadsheetUrl': url,
                    'sheetName': sname,
                })
        return dests

    def _on_add_destination(self):
        data = self._prompt_destination()
        if data is None:
            return
        row = self.dest_table.rowCount()
        self.dest_table.insertRow(row)
        self.dest_table.setItem(row, 0, QTableWidgetItem(data.get('label','')))
        self.dest_table.setItem(row, 1, QTableWidgetItem(data.get('spreadsheetUrl','')))
        self.dest_table.setItem(row, 2, QTableWidgetItem(data.get('sheetName','')))

    def _on_edit_destination(self):
        row = self.dest_table.currentRow()
        if row < 0:
            return
        cur = {
            'label': self.dest_table.item(row, 0).text() if self.dest_table.item(row, 0) else '',
            'spreadsheetUrl': self.dest_table.item(row, 1).text() if self.dest_table.item(row, 1) else '',
            'sheetName': self.dest_table.item(row, 2).text() if self.dest_table.item(row, 2) else '',
        }
        data = self._prompt_destination(cur)
        if data is None:
            return
        self.dest_table.setItem(row, 0, QTableWidgetItem(data.get('label','')))
        self.dest_table.setItem(row, 1, QTableWidgetItem(data.get('spreadsheetUrl','')))
        self.dest_table.setItem(row, 2, QTableWidgetItem(data.get('sheetName','')))

    def _on_delete_destination(self):
        row = self.dest_table.currentRow()
        if row < 0:
            return
        label = self.dest_table.item(row, 0).text() if self.dest_table.item(row, 0) else '(未設定)'
        if QMessageBox.question(self, "削除確認", f"『{label}』を削除しますか？") != QMessageBox.Yes:
            return
        self.dest_table.removeRow(row)

    def _prompt_destination(self, init=None):
        dlg = QDialog(self)
        dlg.setWindowTitle("転記先の編集")
        lay = QVBoxLayout(dlg)
        le_label = QLineEdit(dlg); le_label.setPlaceholderText("表示名（例: 岩田部署管理ドライブ）")
        le_url = QLineEdit(dlg);   le_url.setPlaceholderText("スプレッドシートURL（必須）")
        le_snm = QLineEdit(dlg);   le_snm.setPlaceholderText("シート名（必須）")

        def add_row(caption, w):
            row = QHBoxLayout(); row.addWidget(QLabel(caption + '：')); row.addWidget(w); lay.addLayout(row)
        add_row("表示名", le_label)
        add_row("スプシURL", le_url)
        add_row("シート名", le_snm)
        btns = QHBoxLayout(); btn_ok = QPushButton("OK", dlg); btn_ng = QPushButton("キャンセル", dlg)
        btns.addStretch(); btns.addWidget(btn_ok); btns.addWidget(btn_ng); lay.addLayout(btns)

        if init:
            le_label.setText(init.get('label',''))
            le_url.setText(init.get('spreadsheetUrl',''))
            le_snm.setText(init.get('sheetName',''))

        data_holder = {'ok': False}
        def _ok():
            if not le_label.text().strip():
                QMessageBox.warning(dlg, "入力エラー", "表示名は必須です。")
                return
            if not le_url.text().strip():
                QMessageBox.warning(dlg, "入力エラー", "スプレッドシートURLは必須です。")
                return
            if not le_snm.text().strip():
                QMessageBox.warning(dlg, "入力エラー", "シート名は必須です。")
                return
            data_holder['ok'] = True
            dlg.accept()
        btn_ok.clicked.connect(_ok)
        btn_ng.clicked.connect(dlg.reject)
        if dlg.exec() == QDialog.Accepted and data_holder['ok']:
            return {
                'label': le_label.text().strip(),
                'spreadsheetUrl': le_url.text().strip(),
                'sheetName': le_snm.text().strip(),
            }
        return None