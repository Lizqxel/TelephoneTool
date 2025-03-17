"""
設定ダイアログ

このモジュールは、アプリケーションの設定を管理するための
ダイアログUIを提供します。
"""

import json
import os
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                              QTextEdit, QPushButton, QMessageBox, QSlider,
                              QGroupBox, QSpinBox, QCheckBox)
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
        self.setMinimumSize(600, 500)  # 高さを少し増やす
        
        # 設定ファイルのパス
        self.settings_file = "settings.json"
        
        # デフォルトのフォーマットテンプレート
        self.default_format = """対応者（お客様の名前）：{operator}
工事希望日
★出やすい時間帯：携帯：{mobile}
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

他番号：なし
電話機：プッシュ
禁止回線：なし
ND：

備考：{remarks}
お客様が今使っている回線：アナログ
案内料金：2500円
※リスト名との関係性："""

        # デフォルトのフォントサイズ
        self.default_font_size = 9
        
        # デフォルトの遅延時間（秒）
        self.default_delay = 0
        
        # デフォルトのブラウザ設定
        self.default_browser_settings = {
            "headless": False,
            "show_popup": True,
            "page_load_timeout": 30,
            "script_timeout": 30
        }
        
        # 親ウィンドウから現在の設定を取得
        if parent and hasattr(parent, 'settings'):
            self.current_font_size = parent.settings.get('font_size', self.default_font_size)
            self.current_delay = parent.settings.get('delay_seconds', self.default_delay)
            self.current_browser_settings = parent.settings.get('browser_settings', self.default_browser_settings)
        else:
            self.current_font_size = self.default_font_size
            self.current_delay = self.default_delay
            self.current_browser_settings = self.default_browser_settings
        
        # レイアウトの設定
        layout = QVBoxLayout(self)
        
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
        layout.addWidget(font_size_group)
        
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
        layout.addWidget(delay_group)
        
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
        layout.addWidget(browser_group)
        
        # 説明ラベル
        description = QLabel("CTIフォーマットのテンプレートを編集できます。\n"
                            "以下のプレースホルダーが使用可能です：\n"
                            "{operator}, {mobile}, {contractor}, {furigana}, {birth_date}, {postal_code}, {address}, "
                            "{list_name}, {list_furigana}, {list_phone}, {list_postal_code}, {list_address}, "
                            "{current_line}, {order_date}, {order_person}, {judgment}, {fee}, {net_usage}, {family_approval}, {remarks}")
        description.setWordWrap(True)
        layout.addWidget(description)
        
        # テキスト編集エリア
        self.format_edit = QTextEdit()
        self.format_edit.setPlaceholderText("フォーマットテンプレートを入力してください")
        layout.addWidget(self.format_edit)
        
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
        
        layout.addLayout(button_layout)
        
        # 設定の読み込み
        self.load_settings()
    
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
        self.popup_checkbox.setChecked(self.default_browser_settings["show_popup"])
        self.page_timeout_spin.setValue(self.default_browser_settings["page_load_timeout"])
        self.script_timeout_spin.setValue(self.default_browser_settings["script_timeout"])
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    format_template = settings.get('format_template', self.default_format)
                    font_size = settings.get('font_size', self.default_font_size)
                    delay_seconds = settings.get('delay_seconds', self.default_delay)
                    browser_settings = settings.get('browser_settings', self.default_browser_settings)
                    
                    self.format_edit.setText(format_template)
                    self.font_size_slider.setValue(font_size)
                    self.delay_spin.setValue(delay_seconds)
                    
                    # ブラウザ設定の読み込み
                    self.headless_checkbox.setChecked(browser_settings.get("headless", False))
                    self.disable_images_checkbox.setChecked(browser_settings.get("disable_images", True))
                    self.popup_checkbox.setChecked(browser_settings.get("show_popup", True))
                    self.page_timeout_spin.setValue(browser_settings.get("page_load_timeout", 30))
                    self.script_timeout_spin.setValue(browser_settings.get("script_timeout", 30))
            else:
                self.format_edit.setText(self.default_format)
                self.font_size_slider.setValue(self.default_font_size)
                self.delay_spin.setValue(self.default_delay)
                self.reset_browser_settings()
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
            self.format_edit.setText(self.default_format)
            self.font_size_slider.setValue(self.default_font_size)
            self.delay_spin.setValue(self.default_delay)
            self.reset_browser_settings()
    
    def save_settings(self):
        """設定をファイルに保存する"""
        try:
            # ブラウザ設定を取得
            browser_settings = {
                "headless": self.headless_checkbox.isChecked(),
                "disable_images": self.disable_images_checkbox.isChecked(),
                "show_popup": self.popup_checkbox.isChecked(),
                "page_load_timeout": self.page_timeout_spin.value(),
                "script_timeout": self.script_timeout_spin.value()
            }
            
            settings = {
                'format_template': self.format_edit.toPlainText(),
                'font_size': self.font_size_slider.value(),
                'delay_seconds': self.delay_spin.value(),
                'browser_settings': browser_settings
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"設定の保存に失敗しました: {str(e)}")
            return False
    
    def reset_to_default(self):
        """設定をデフォルトに戻す"""
        self.format_edit.setText(self.default_format)
        self.font_size_slider.setValue(self.default_font_size)
        self.delay_spin.setValue(self.default_delay)
        self.reset_browser_settings()
    
    def accept(self):
        """ダイアログを受け入れる（OKボタン）"""
        if self.save_settings():
            super().accept()
    
    def get_settings(self):
        """現在の設定を取得する"""
        browser_settings = {
            "headless": self.headless_checkbox.isChecked(),
            "disable_images": self.disable_images_checkbox.isChecked(),
            "show_popup": self.popup_checkbox.isChecked(),
            "page_load_timeout": self.page_timeout_spin.value(),
            "script_timeout": self.script_timeout_spin.value()
        }
        
        return {
            'format_template': self.format_edit.toPlainText(),
            'font_size': self.font_size_slider.value(),
            'delay_seconds': self.delay_spin.value(),
            'browser_settings': browser_settings
        } 