"""
カスタムウィジェットモジュール

このモジュールは、アプリケーションで使用するカスタムウィジェットを提供します。
"""

from PySide6.QtWidgets import QComboBox
from PySide6.QtCore import Qt

class CustomComboBox(QComboBox):
    """スクロールでの値変更を防止するカスタムコンボボックス"""
    
    def wheelEvent(self, event):
        """ホイールイベントを無視"""
        event.ignore() 