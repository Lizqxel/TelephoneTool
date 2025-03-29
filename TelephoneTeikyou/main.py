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
import logging
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow


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
    # アプリケーションの作成
    app = QApplication(sys.argv)
    
    # メインウィンドウの作成と表示
    window = MainWindow()
    window.show()
    
    # アプリケーションの実行
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 