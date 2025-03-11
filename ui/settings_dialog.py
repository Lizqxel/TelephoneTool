"""
設定ダイアログ

このモジュールは、アプリケーションの設定を管理するための
ダイアログUIを提供します。
"""

import json
import os
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                              QTextEdit, QPushButton, QMessageBox)
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
        self.setMinimumSize(600, 400)
        
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
        
        # レイアウトの設定
        layout = QVBoxLayout(self)
        
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
    
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    format_template = settings.get('format_template', self.default_format)
                    self.format_edit.setText(format_template)
            else:
                self.format_edit.setText(self.default_format)
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"設定の読み込みに失敗しました: {str(e)}")
            self.format_edit.setText(self.default_format)
    
    def save_settings(self):
        """設定をファイルに保存する"""
        try:
            settings = {
                'format_template': self.format_edit.toPlainText()
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
    
    def accept(self):
        """ダイアログを受け入れる（OKボタン）"""
        if self.save_settings():
            super().accept()
    
    def get_settings(self):
        """現在の設定を取得する"""
        return {
            'format_template': self.format_edit.toPlainText()
        } 