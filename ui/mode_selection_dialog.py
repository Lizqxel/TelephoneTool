"""
モード選択ダイアログ

このモジュールは、アプリケーションのモード選択ダイアログを提供します。
ユーザーは「シンプルモード」または「使いやすいモード」を選択できます。
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QPushButton, QLabel,
    QMessageBox, QCheckBox, QWidget, QHBoxLayout
)
from PySide6.QtCore import Qt
import logging

class ModeSelectionDialog(QDialog):
    """
    モード選択ダイアログクラス
    
    アプリケーションの動作モードを選択するためのダイアログを提供します。
    """
    
    def __init__(self, parent=None):
        """
        モード選択ダイアログの初期化
        
        Args:
            parent: 親ウィンドウ
        """
        super().__init__(parent)
        self.setWindowTitle("モード選択")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.setMinimumHeight(300)
        
        # レイアウトの作成
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # タイトルラベル
        title_label = QLabel("使用するモードを選択してください")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #333;
                margin: 10px;
                padding: 10px;
                background-color: #f5f5f5;
                border-radius: 5px;
            }
        """)
        layout.addWidget(title_label)
        
        # 説明ラベル
        description_label = QLabel(
            "通常モード：すべての入力項目を1画面で入力\n"
            "誘導モード：入力項目を複数の画面に分けて入力\n"
            "法人モード：通常モードをベースに法人向け入力補助を追加"
        )
        description_label.setAlignment(Qt.AlignCenter)
        description_label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #666;
                margin: 5px;
                padding: 10px;
                background-color: #f9f9f9;
                border-radius: 5px;
            }
        """)
        layout.addWidget(description_label)
        
        # 通常モードボタン
        simple_button = QPushButton("通常モード")
        simple_button.setMinimumHeight(50)
        simple_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 16px;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        simple_button.clicked.connect(lambda: self.select_mode('simple'))
        layout.addWidget(simple_button)
        
        # 誘導モードボタン
        easy_button = QPushButton("誘導モード")
        easy_button.setMinimumHeight(50)
        easy_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 16px;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
        """)
        easy_button.clicked.connect(lambda: self.select_mode('easy'))
        layout.addWidget(easy_button)

        # 法人モードボタン
        corporate_button = QPushButton("法人モード")
        corporate_button.setMinimumHeight(50)
        corporate_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
                font-size: 16px;
                margin: 5px;
            }
            QPushButton:hover {
                background-color: #FB8C00;
            }
            QPushButton:pressed {
                background-color: #F57C00;
            }
        """)
        corporate_button.clicked.connect(lambda: self.select_mode('corporate'))
        layout.addWidget(corporate_button)
        
        # 次回から表示しないチェックボックス
        checkbox_widget = QWidget()
        checkbox_layout = QHBoxLayout()
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        
        self.show_again_checkbox = QCheckBox("次回から表示しない")
        self.show_again_checkbox.setChecked(True)
        self.show_again_checkbox.setStyleSheet("""
            QCheckBox {
                font-size: 14px;
                color: #666;
                padding: 5px;
            }
            QCheckBox:hover {
                color: #333;
            }
        """)
        checkbox_layout.addWidget(self.show_again_checkbox)
        checkbox_layout.addStretch()
        checkbox_widget.setLayout(checkbox_layout)
        layout.addWidget(checkbox_widget)
        
        self.setLayout(layout)
        
        # 選択されたモードを保存する変数
        self.selected_mode = None
        
    def select_mode(self, mode):
        """
        モードを選択し、ダイアログを閉じる
        
        Args:
            mode: 選択されたモード（'simple'、'easy'、'corporate'）
        """
        try:
            self.selected_mode = mode
            logging.info(f"モード {mode} が選択されました")
            self.accept()
        except Exception as e:
            logging.error(f"モード選択中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"モードの選択中にエラーが発生しました: {e}")
            
    def get_selected_mode(self):
        """
        選択されたモードを取得
        
        Returns:
            str: 選択されたモード（'simple'または'easy'）
        """
        return self.selected_mode
        
    def should_show_again(self):
        """
        次回から表示しないかどうかを取得
        
        Returns:
            bool: 次回から表示しない場合はTrue
        """
        return not self.show_again_checkbox.isChecked() 