"""
提供エリア判定ツール

このスクリプトは、NTT西日本の提供エリア判定と
顧客情報取得機能を提供するGUIアプリケーションです。

主な機能：
- 提供エリア判定
- 顧客情報取得
- フォントサイズ設定
- ブラウザ設定
"""

import sys
import os
import warnings
import logging
from PySide6.QtWidgets import QApplication, QMessageBox

from ui.main_window import MainWindow

# PyInstallerの一時ディレクトリ警告を抑制
if getattr(sys, 'frozen', False):
    # 一時ディレクトリ関連の警告を無視
    warnings.filterwarnings("ignore", category=UserWarning)
    
    # PyInstallerの一時ディレクトリ削除エラーを無視するための設定
    try:
        import atexit
        def cleanup_temp_dirs():
            """一時ディレクトリのクリーンアップ処理（エラーを無視）"""
            try:
                # 何もしない（PyInstallerの自動クリーンアップに任せる）
                pass
            except Exception:
                # エラーを無視
                pass
        
        # 終了時処理を登録
        atexit.register(cleanup_temp_dirs)
    except Exception:
        pass


# ログの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('app.log', encoding='utf-8')
    ]
)


def main():
    """アプリケーションのメイン関数"""
    try:
        # アプリケーションの作成
        app = QApplication(sys.argv)
        
        # PyInstallerの一時ディレクトリ警告ダイアログを無効化
        if getattr(sys, 'frozen', False):
            # 一時ディレクトリエラーのダイアログ表示を抑制
            try:
                import ctypes
                # Windows APIでメッセージボックスを無効化（Windows環境のみ）
                if os.name == 'nt':
                    # SEM_FAILCRITICALERRORS フラグを設定
                    ctypes.windll.kernel32.SetErrorMode(0x0001)
            except Exception:
                pass
        
        # メインウィンドウの作成と表示
        window = MainWindow()
        window.show()
        
        # アプリケーションの実行
        sys.exit(app.exec())
        
    except Exception as e:
        # 起動時エラーを処理（一時ディレクトリエラー以外）
        if "temporary directory" not in str(e).lower() and "temp_" not in str(e).lower():
            logging.error(f"アプリケーション起動エラー: {e}")
            try:
                app = QApplication(sys.argv) if 'app' not in locals() else app
                QMessageBox.critical(None, "起動エラー", f"アプリケーションの起動中にエラーが発生しました:\n{str(e)}")
            except:
                print(f"起動エラー: {e}")
        else:
            # 一時ディレクトリエラーは無視して再試行
            logging.warning(f"一時ディレクトリエラーを無視して続行: {e}")
            try:
                app = QApplication(sys.argv) if 'app' not in locals() else app
                window = MainWindow()
                window.show()
                sys.exit(app.exec())
            except Exception as retry_error:
                logging.error(f"再試行でもエラー: {retry_error}")
                sys.exit(1)


if __name__ == "__main__":
    main() 