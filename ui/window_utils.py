"""
ウィンドウ関連のユーティリティ関数を提供するモジュール

このモジュールは、ウィンドウ操作に関する共通のユーティリティ関数を提供します。
主な機能：
- ウィンドウの位置とサイズの管理
- ウィンドウのスタイル設定
- ダイアログの表示
- エラーメッセージの表示

制限事項：
- ウィンドウの最小サイズは800x600
- ウィンドウの最大サイズは1920x1080
"""

import os
import sys
import json
import logging
from typing import Dict, Optional, Tuple

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QMessageBox, QDialog,
    QVBoxLayout, QLabel, QPushButton, QFileDialog
)
from PySide6.QtCore import Qt, QSize, QPoint, QRect
from PySide6.QtGui import QFont, QIcon

def center_window(window: QMainWindow) -> None:
    """
    ウィンドウを画面中央に配置します。
    
    Args:
        window (QMainWindow): 中央配置するウィンドウ
    """
    try:
        screen = window.screen().availableGeometry()
        size = window.size()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        window.move(x, y)
    except Exception as e:
        logging.error(f"ウィンドウの中央配置中にエラーが発生しました: {str(e)}")

def set_window_style(window: QMainWindow) -> None:
    """
    ウィンドウのスタイルを設定します。
    
    Args:
        window (QMainWindow): スタイルを設定するウィンドウ
    """
    try:
        # 最小サイズの設定
        window.setMinimumSize(800, 600)
        
        # 最大サイズの設定
        window.setMaximumSize(1920, 1080)
        
        # フォントの設定
        font = QFont("Meiryo UI", 9)
        window.setFont(font)
        
        # スタイルシートの設定
        window.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
            QTableWidget {
                border: 1px solid #cccccc;
                gridline-color: #dddddd;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                border: 1px solid #cccccc;
            }
        """)
    except Exception as e:
        logging.error(f"ウィンドウのスタイル設定中にエラーが発生しました: {str(e)}")

def show_error_dialog(parent: QWidget, title: str, message: str) -> None:
    """
    エラーダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        message (str): エラーメッセージ
    """
    try:
        QMessageBox.critical(parent, title, message)
    except Exception as e:
        logging.error(f"エラーダイアログの表示中にエラーが発生しました: {str(e)}")

def show_warning_dialog(parent: QWidget, title: str, message: str) -> None:
    """
    警告ダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        message (str): 警告メッセージ
    """
    try:
        QMessageBox.warning(parent, title, message)
    except Exception as e:
        logging.error(f"警告ダイアログの表示中にエラーが発生しました: {str(e)}")

def show_info_dialog(parent: QWidget, title: str, message: str) -> None:
    """
    情報ダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        message (str): 情報メッセージ
    """
    try:
        QMessageBox.information(parent, title, message)
    except Exception as e:
        logging.error(f"情報ダイアログの表示中にエラーが発生しました: {str(e)}")

def show_question_dialog(parent: QWidget, title: str, message: str) -> bool:
    """
    質問ダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        message (str): 質問メッセージ
        
    Returns:
        bool: ユーザーが「はい」を選択した場合はTrue、それ以外はFalse
    """
    try:
        reply = QMessageBox.question(
            parent, title, message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        return reply == QMessageBox.Yes
    except Exception as e:
        logging.error(f"質問ダイアログの表示中にエラーが発生しました: {str(e)}")
        return False

def show_file_dialog(
    parent: QWidget,
    title: str,
    directory: str,
    filter: str
) -> Optional[str]:
    """
    ファイル選択ダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        directory (str): 初期ディレクトリ
        filter (str): ファイルフィルタ
        
    Returns:
        Optional[str]: 選択されたファイルのパス、キャンセルされた場合はNone
    """
    try:
        file_path, _ = QFileDialog.getOpenFileName(
            parent, title, directory, filter
        )
        return file_path if file_path else None
    except Exception as e:
        logging.error(f"ファイル選択ダイアログの表示中にエラーが発生しました: {str(e)}")
        return None

def show_save_dialog(
    parent: QWidget,
    title: str,
    directory: str,
    filter: str
) -> Optional[str]:
    """
    ファイル保存ダイアログを表示します。
    
    Args:
        parent (QWidget): 親ウィジェット
        title (str): ダイアログのタイトル
        directory (str): 初期ディレクトリ
        filter (str): ファイルフィルタ
        
    Returns:
        Optional[str]: 選択されたファイルのパス、キャンセルされた場合はNone
    """
    try:
        file_path, _ = QFileDialog.getSaveFileName(
            parent, title, directory, filter
        )
        return file_path if file_path else None
    except Exception as e:
        logging.error(f"ファイル保存ダイアログの表示中にエラーが発生しました: {str(e)}")
        return None 