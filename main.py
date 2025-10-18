"""
コールセンター業務効率化ツール

このスクリプトは、コールセンター業務の効率化を目的としたGUIアプリケーションです。
PySide6を使用してUIを構築し、Google Spreadsheetsとの連携機能を提供します。

主な機能：
- 顧客情報の入力
- CTIフォーマットの生成
- スプレッドシートへのデータ転記
- クリップボード監視機能
- 提供エリア検索機能
"""

import sys
import logging
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow
from first_run_setup import ensure_settings_file


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
    # 初回セットアップ（settings.json 生成と共有トークンの取得）
    try:
        ensure_settings_file()
    except Exception as e:
        logging.warning(f"初回セットアップ中に問題が発生しました: {e}")
    
    # メインウィンドウの作成と表示
    window = MainWindow()
    window.show()
    
    # アプリケーションの実行
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 