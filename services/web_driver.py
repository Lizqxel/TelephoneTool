"""
Webドライバーサービス

このモジュールは、Seleniumを使用したWebドライバーの
作成と管理に関する機能を提供します。
"""

import logging
import json
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def create_driver(headless=False):
    """
    ChromeのWebDriverインスタンスを作成する関数
    
    Args:
        headless (bool): ヘッドレスモードを有効にするかどうか
        
    Returns:
        webdriver.Chrome: 設定済みのChromeドライバーインスタンス
        
    Raises:
        Exception: ドライバーの作成に失敗した場合
    """
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    # ブラウザ設定を読み込む
    browser_settings = load_browser_settings()
    
    # ヘッドレスモードの設定
    if headless or browser_settings.get("headless", False):
        options.add_argument("--headless=new")  # 新しいヘッドレスモード
    
    # パフォーマンス最適化のための設定
    options.add_argument("--window-size=1280,720")  # さらに小さいウィンドウサイズ
    options.add_argument("--disable-extensions")  # 拡張機能を無効化
    options.add_argument("--disable-popup-blocking")  # ポップアップブロックを無効化
    options.add_argument("--disable-blink-features=AutomationControlled")  # 自動化検出を回避
    options.add_argument("--lang=ja")  # 言語を日本語に設定
    
    # メモリ使用量を制限
    options.add_argument("--js-flags=--max-old-space-size=256")  # JavaScriptのメモリ制限をさらに小さく
    options.add_argument("--disable-infobars")  # 情報バーを無効化
    options.add_argument("--disable-notifications")  # 通知を無効化
    options.add_argument("--disable-default-apps")  # デフォルトアプリを無効化
    
    # 画像読み込みを無効化（大幅な高速化）
    if browser_settings.get("disable_images", True):
        options.add_argument("--blink-settings=imagesEnabled=false")
    
    # キャッシュを無効化
    options.add_argument("--disable-application-cache")
    options.add_argument("--disable-cache")
    
    # プロセス数を制限
    options.add_argument("--single-process")
    
    # 追加の高速化オプション
    options.add_argument("--disable-accelerated-2d-canvas")
    options.add_argument("--disable-accelerated-jpeg-decoding")
    options.add_argument("--disable-accelerated-video-decode")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-site-isolation-trials")
    
    # User-Agentを設定
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")
    
    try:
        # ChromeDriverManagerを使用して自動的にドライバーをダウンロード
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        # ページの読み込みタイムアウトを設定
        driver.set_page_load_timeout(browser_settings.get("page_load_timeout", 30))
        
        # JavaScriptの実行を待機
        driver.set_script_timeout(browser_settings.get("script_timeout", 30))
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗しました: {str(e)}")
        raise

def load_browser_settings():
    """
    ブラウザ設定をファイルから読み込む
    
    Returns:
        dict: ブラウザ設定
    """
    default_settings = {
        "headless": True,
        "page_load_timeout": 30,
        "script_timeout": 30,
        "disable_images": True,
        "show_popup": False
    }
    
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                # ブラウザ設定が含まれていない場合はデフォルト値を使用
                browser_settings = settings.get("browser_settings", default_settings)
                return browser_settings
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
    
    return default_settings 