"""
Webドライバーサービス

このモジュールは、Seleniumを使用したWebドライバーの
作成と管理に関する機能を提供します。
"""

import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def create_driver():
    """
    ChromeのWebDriverインスタンスを作成する関数
    
    Returns:
        webdriver.Chrome: 設定済みのChromeドライバーインスタンス
        
    Raises:
        Exception: ドライバーの作成に失敗した場合
    """
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    # options.add_argument("--headless")  # デバッグ用にヘッドレスモードを無効化
    options.add_argument("--window-size=1920,1080")  # ウィンドウサイズを設定
    options.add_argument("--start-maximized")  # ウィンドウを最大化
    options.add_argument("--disable-extensions")  # 拡張機能を無効化
    options.add_argument("--disable-popup-blocking")  # ポップアップブロックを無効化
    options.add_argument("--disable-blink-features=AutomationControlled")  # 自動化検出を回避
    options.add_argument("--lang=ja")  # 言語を日本語に設定
    
    # User-Agentを設定
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")
    
    try:
        # ChromeDriverManagerを使用して自動的にドライバーをダウンロード
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        # ウィンドウを最大化（デバッグ用）
        driver.maximize_window()
        
        # ページの読み込みタイムアウトを設定
        driver.set_page_load_timeout(30)
        
        # JavaScriptの実行を待機
        driver.set_script_timeout(30)
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗しました: {str(e)}")
        raise 