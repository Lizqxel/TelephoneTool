"""
アップデート進捗を表示するダイアログ

このモジュールは、アプリケーションのアップデート進捗を
表示するためのポップアップダイアログを提供します。
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QProgressBar,
    QPushButton, QMessageBox
)
from PySide6.QtCore import Qt, QTimer

class UpdateProgressDialog(QDialog):
    """アップデート進捗ダイアログ"""
    
    def __init__(self, parent=None):
        """初期化"""
        super().__init__(parent)
        self.setWindowTitle("アップデート")
        self.setFixedWidth(400)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        
        self.init_ui()
        
    def init_ui(self):
        """UIの初期化"""
        layout = QVBoxLayout(self)
        
        # ステータスラベル
        self.status_label = QLabel("アップデートを準備中...")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # キャンセルボタン
        self.cancel_button = QPushButton("キャンセル")
        self.cancel_button.clicked.connect(self.cancel_update)
        layout.addWidget(self.cancel_button)
        
    def update_progress(self, value, status_text=None):
        """進捗を更新"""
        self.progress_bar.setValue(value)
        if status_text:
            self.status_label.setText(status_text)
            
    def cancel_update(self):
        """アップデートをキャンセル"""
        reply = QMessageBox.question(
            self,
            "アップデートのキャンセル",
            "アップデートをキャンセルしますか？\nキャンセルすると、アプリケーションは現在のバージョンのままです。",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.reject()
            
    def show_error(self, error_message):
        """エラーメッセージを表示"""
        self.status_label.setText("エラーが発生しました")
        self.progress_bar.setValue(0)
        self.cancel_button.setText("閉じる")
        QMessageBox.critical(self, "エラー", error_message) 