"""
ブラウザオプション統合モジュール

このモジュールは、最適化されたブラウザオプションをSeleniumドライバー初期化に統合します。
"""

import os
import logging
from selenium import webdriver
from browser_options import get_optimized_options

def get_optimized_driver(headless=False):
    """
    最適化されたChromeドライバーを取得します
    
    Args:
        headless (bool): ヘッドレスモードを有効にするかどうか
        
    Returns:
        WebDriver: 最適化された設定のChromeドライバー
    """
    try:
        # 最適化されたオプションを取得
        options = get_optimized_options(headless)
        
        # ユーザーデータディレクトリの設定（オプション）
        chrome_data_dir = os.path.join(os.getcwd(), 'chrome_data')
        if os.path.exists(chrome_data_dir):
            options.add_argument(f'--user-data-dir={chrome_data_dir}')
        
        # Chromeドライバーの初期化
        driver = webdriver.Chrome(options=options)
        
        # ウィンドウ位置を調整（画面の中央に配置）
        driver.set_window_position(0, 0)
        
        return driver
    except Exception as e:
        logging.error(f"最適化されたドライバーの初期化中にエラー: {e}")
        # 通常のChromeドライバーにフォールバック
        return webdriver.Chrome() 