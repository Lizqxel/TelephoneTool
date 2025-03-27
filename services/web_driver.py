"""
Seleniumウェブドライバーの設定

このモジュールは、Seleniumウェブドライバーの作成と
設定を行う機能を提供します。
"""

import logging
import json
import os
import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType


def create_driver(headless=False):
    """
    Chromeドライバーを作成する
    
    Args:
        headless (bool): ヘッドレスモードで起動するかどうか
        
    Returns:
        webdriver.Chrome: 設定済みのChromeドライバー
    """
    # Chromeオプションを設定
    chrome_options = Options()
    
    if headless:
        chrome_options.add_argument('--headless')
    
    # その他の共通オプション
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument('--ignore-certificate-errors')
    
    # ユーザーデータディレクトリを設定
    chrome_data_dir = os.path.join(os.getcwd(), 'chrome_data')
    if not os.path.exists(chrome_data_dir):
        os.makedirs(chrome_data_dir)
    chrome_options.add_argument(f'--user-data-dir={chrome_data_dir}')
    
    try:
        # ドライバーを作成
        driver = webdriver.Chrome(options=chrome_options)
        logging.info("Chromeドライバーを作成しました")
        return driver
        
    except Exception as e:
        logging.error(f"Chromeドライバーの作成に失敗: {str(e)}")
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
        "show_popup": False,
        "auto_close": False
    }
    
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                # ブラウザ設定が含まれていない場合はデフォルト値を使用
                browser_settings = settings.get("browser_settings", default_settings)
                
                # auto_closeが設定に含まれていない場合はデフォルト値を使用
                if "auto_close" not in browser_settings:
                    browser_settings["auto_close"] = default_settings["auto_close"]
                    
                return browser_settings
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
    
    return default_settings 