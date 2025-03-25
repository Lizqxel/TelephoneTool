"""
Webドライバーサービス

このモジュールは、Seleniumを使用したWebドライバーの
作成と管理に関する機能を提供します。
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
    ChromeのWebDriverインスタンスを作成する関数
    
    Args:
        headless (bool): ヘッドレスモードを有効にするかどうか
        
    Returns:
        webdriver.Chrome: 設定済みのChromeドライバーインスタンス
        
    Raises:
        Exception: ドライバーの作成に失敗した場合
    """
    options = Options()
    
    # 基本的な安定性のための設定
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    # ブラウザ設定を読み込む
    browser_settings = load_browser_settings()
    
    # ヘッドレスモードの設定
    # 引数でFalseが指定された場合は、必ずヘッドレスモードを無効化
    if headless == False:
        is_headless = False
        logging.info("引数指定によりヘッドレスモードを無効化します")
    else:
        # 引数がTrueまたはデフォルト値の場合は設定ファイルの値を使用
        is_headless = headless or browser_settings.get("headless", False)
    
    if is_headless:
        options.add_argument("--headless=new")  # 新しいヘッドレスモード
        logging.info("ヘッドレスモードで起動します")
    else:
        logging.info("通常モード（ウィンドウ表示）で起動します")
    
    # WebGLとグラフィックス関連のエラーを回避するための設定
    options.add_argument("--use-gl=swiftshader")  # SwiftShaderによるソフトウェアレンダリング
    options.add_argument("--use-angle=swiftshader")  # ANGLEバックエンドとしてSwiftShaderを使用
    options.add_argument("--ignore-gpu-blocklist")  # GPUブロックリストを無視
    options.add_argument("--allow-insecure-localhost")  # 安全でないlocalhostを許可
    options.add_argument("--allow-running-insecure-content")  # 安全でないコンテンツの実行を許可
    
    # ブラウザのクラッシュを防ぐための設定
    options.add_argument("--disable-features=NetworkService")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--disable-breakpad")  # クラッシュレポート機能を無効化
    options.add_argument("--disable-component-update")  # コンポーネント更新を無効化
    
    # ウィンドウサイズの設定を変更して軽量化する
    # ウィンドウサイズの設定（ヘッドレスモードでなければ適度な小さいサイズに）
    if not is_headless:
        # options.add_argument("--start-maximized")  # 最大化はしない
        options.add_argument("--window-size=800,600")  # より小さなウィンドウサイズ
        logging.info("ウィンドウサイズを 800x600 に設定しました（処理負荷軽減のため）")
    else:
        options.add_argument("--window-size=800,600")  # ヘッドレスモード用のサイズも小さく
    
    # パフォーマンス最適化のための設定
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=ja")
    
    # 通知関連の設定
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    
    # 画像読み込みを無効化（高速化） - ウィンドウ表示の場合は画像を有効化
    if browser_settings.get("disable_images", True) and is_headless:
        options.add_argument("--blink-settings=imagesEnabled=false")
    
    # プロセス関連の安定性向上設定
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    
    # User-Agentを設定
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")
    
    try:
        # PyInstallerでビルドされた場合のベースディレクトリを取得
        if getattr(sys, 'frozen', False):
            # PyInstallerでビルドされた場合
            base_dir = sys._MEIPASS
            driver_path = os.path.join(base_dir, "drivers", "chromedriver-win32", "chromedriver.exe")
        else:
            # 通常の実行の場合
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            driver_path = os.path.join(base_dir, "drivers", "chromedriver-win32", "chromedriver.exe")
        
        logging.info(f"ベースディレクトリ: {base_dir}")
        logging.info(f"ドライバーパス: {driver_path}")
        
        # ドライバーパスが存在しない場合、ChromeDriverManagerにフォールバック
        if not os.path.exists(driver_path):
            logging.warning(f"ドライバーが見つかりません: {driver_path}")
            logging.info("ChromeDriverManagerを使用してドライバーをダウンロードします")
            try:
                driver_path = ChromeDriverManager().install()
                # ダウンロードされたパスが実際に存在するか確認
                if not os.path.exists(driver_path):
                    raise FileNotFoundError(f"ダウンロードされたドライバーが見つかりません: {driver_path}")
            except Exception as e:
                logging.error(f"ChromeDriverManagerでのドライバーダウンロードに失敗: {str(e)}")
                raise
        
        logging.info(f"使用するドライバーパス: {driver_path}")
        service = ChromeService(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        
        # ページの読み込みタイムアウトを設定
        driver.set_page_load_timeout(browser_settings.get("page_load_timeout", 30))
        
        # JavaScriptの実行を待機
        driver.set_script_timeout(browser_settings.get("script_timeout", 30))
        
        # 非ヘッドレスモードの場合のウィンドウ処理修正
        # 非ヘッドレスモードの場合、ウィンドウを前面に表示
        if not is_headless:
            # JavaScriptを使ってウィンドウをフォーカス
            driver.execute_script("window.focus();")
            # ウィンドウサイズを確実に適用（最大化はしない）
            # driver.maximize_window()
            logging.info(f"ブラウザウィンドウを表示しました: {driver.get_window_size()}")
        
        logging.info("Chromeドライバーが正常に初期化されました")
        
        # 自動終了設定をドライバーに追加
        driver._auto_close = browser_settings.get("auto_close", False)
        driver._headless = is_headless
        driver._show_popup = browser_settings.get("show_popup", False)
        
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