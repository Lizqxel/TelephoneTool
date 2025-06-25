"""
PySide6互換性修正モジュール

PySide6のQEventとイベント処理に関連する互換性の問題を修正するためのモジュールです。
アプリケーション起動時に自動的に適用されます。
"""

import sys
import logging
from PySide6.QtCore import QEvent

# オリジナルの初期化メソッドを保存
original_init = QEvent.__init__

def patched_init(self, *args, **kwargs):
    """
    QEvent初期化の修正版
    
    intで直接初期化しようとした場合、適切なQEvent.Typeに変換します
    """
    try:
        # オリジナルの初期化を試みる
        original_init(self, *args, **kwargs)
    except TypeError as e:
        # 整数が渡された場合、QEvent.Typeに変換する
        if len(args) == 1 and isinstance(args[0], int):
            try:
                # QEvent.Typeとして再初期化
                event_type = QEvent.Type(args[0])
                original_init(self, event_type)
            except Exception as conv_error:
                logging.error(f"PySide6 QEvent初期化パッチ適用中にエラー: {conv_error}")
                raise
        else:
            # その他のエラーはそのまま伝播
            raise

def apply_patches():
    """パッチを適用します"""
    # QEvent.__init__をパッチ適用
    QEvent.__init__ = patched_init
    logging.info("PySide6 QEvent互換性パッチが適用されました")
    
    # PyInstallerの一時ディレクトリエラー抑制
    if getattr(sys, 'frozen', False):
        try:
            import os
            import warnings
            import ctypes
            
            # 警告の抑制
            warnings.filterwarnings("ignore", message=".*temporary directory.*")
            warnings.filterwarnings("ignore", message=".*Temp_.*")
            
            # Windows環境でのエラーダイアログ抑制
            if os.name == 'nt':
                try:
                    # SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX
                    ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x8000)
                    logging.info("Windows エラーダイアログ抑制が適用されました")
                except Exception as e:
                    logging.warning(f"Windows エラーダイアログ抑制の適用に失敗: {e}")
            
            logging.info("PyInstaller一時ディレクトリエラー抑制パッチが適用されました")
            
        except Exception as e:
            logging.warning(f"一時ディレクトリエラー抑制パッチの適用に失敗: {e}") 