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
    Chromeドライバーを作成する関数
    
    Args:
        headless (bool): ヘッドレスモードで起動するかどうか
        
    Returns:
        WebDriver: 作成されたドライバーインスタンス
    """
    try:
        options = webdriver.ChromeOptions()
        
        if headless:
            options.add_argument('--headless=new')
        
        # ネットワークとページロードの最適化
        options.page_load_strategy = 'none'  # ページロードを待たずに操作を可能にする
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--no-sandbox')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-features=IsolateOrigins,site-per-process')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--dns-prefetch-disable')  # DNS prefetchを無効化
        options.add_argument('--disable-background-networking')  # バックグラウンドネットワークを無効化
        
        # メモリとキャッシュの最適化
        options.add_argument('--disable-gpu-shader-disk-cache')  # GPUシェーダーキャッシュを無効化
        options.add_argument('--disable-gpu-program-cache')  # GPUプログラムキャッシュを無効化
        options.add_argument('--disable-software-rasterizer')  # ソフトウェアラスタライザを無効化
        
        # リソース読み込みの最適化
        prefs = {
            'profile.default_content_setting_values': {
                'images': 1,  # 画像を許可
                'javascript': 1,  # JavaScriptを許可
                'cookies': 1,  # Cookieを許可
                'plugins': 2,  # プラグインをブロック
                'popups': 2,  # ポップアップをブロック
                'notifications': 2,  # 通知をブロック
                'auto_select_certificate': 2,  # 証明書自動選択をブロック
                'fullscreen': 2,  # フルスクリーンをブロック
            },
            'profile.managed_default_content_settings': {
                'images': 1,
                'javascript': 1
            },
            'profile.password_manager_enabled': False,
            'profile.default_content_settings.popups': 0,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': True,
            'disk-cache-size': 104857600,  # キャッシュサイズを100MBに設定
            'network.http.max-connections-per-server': 10,  # サーバーごとの最大接続数
            'network.http.max-persistent-connections-per-server': 5,  # 永続的な接続の最大数
        }
        options.add_experimental_option('prefs', prefs)
        
        # 不要な機能を無効化
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        
        # ウィンドウサイズを設定
        options.add_argument('--window-size=800,600')
        
        # 永続的なプロファイルディレクトリを設定
        user_data_dir = os.path.abspath(os.path.join(os.getcwd(), 'chrome_data'))
        if not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir)
        options.add_argument(f'--user-data-dir={user_data_dir}')
        
        # プロファイルを指定
        profile_directory = 'Default'
        options.add_argument(f'--profile-directory={profile_directory}')
        
        # 高速化のための追加設定
        options.add_argument('--disable-extensions')  # 拡張機能を無効化
        options.add_argument('--disable-sync')  # 同期を無効化
        options.add_argument('--disable-default-apps')  # デフォルトアプリを無効化
        options.add_argument('--no-default-browser-check')  # デフォルトブラウザチェックを無効化
        options.add_argument('--no-first-run')  # 初回実行時の処理をスキップ
        options.add_argument('--disable-prompt-on-repost')  # 再投稿時のプロンプトを無効化
        
        driver = webdriver.Chrome(options=options)
        
        # タイムアウト設定
        driver.set_page_load_timeout(30)  # ページロードのタイムアウトを30秒に設定
        driver.set_script_timeout(30)  # スクリプト実行のタイムアウトを30秒に設定
        
        logging.info("最適化されたChromeドライバーを作成しました")
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗: {str(e)}")
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