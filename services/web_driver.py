"""
WebDriverの作成と管理を行うモジュール

このモジュールは、Selenium WebDriverの作成と管理を
担当します。
"""

import logging
import json
import os
import time
import sys
import re
import subprocess
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType


def _resolve_chromedriver_executable(installed_path: str) -> str:
    """
    webdriver_managerが返したパスから実行可能なchromedriver.exeを解決する

    Args:
        installed_path (str): webdriver_managerが返すパス

    Returns:
        str: chromedriver.exeの実パス
    """
    try:
        if installed_path and installed_path.lower().endswith("chromedriver.exe") and os.path.exists(installed_path):
            return installed_path

        path_obj = Path(installed_path)
        search_roots = []

        if path_obj.exists():
            if path_obj.is_dir():
                search_roots.append(path_obj)
            else:
                search_roots.append(path_obj.parent)

        # webdriver_managerのキャッシュ構造が変わった場合に備えて親階層も探索
        for root in list(search_roots):
            search_roots.append(root.parent)
            search_roots.append(root.parent.parent)

        checked = set()
        for root in search_roots:
            try:
                root = root.resolve()
            except Exception:
                continue

            if str(root) in checked or not root.exists() or not root.is_dir():
                continue
            checked.add(str(root))

            direct_candidate = root / "chromedriver.exe"
            if direct_candidate.exists():
                return str(direct_candidate)

            recursive_candidates = list(root.rglob("chromedriver.exe"))
            if recursive_candidates:
                # より深い階層の候補より、パスの短いものを優先
                recursive_candidates.sort(key=lambda p: len(str(p)))
                return str(recursive_candidates[0])

        raise FileNotFoundError(f"chromedriver.exeが見つかりません: {installed_path}")

    except Exception as e:
        logging.warning(f"chromedriver.exeの解決に失敗: {str(e)}")
        raise


def _find_bundled_chromedriver() -> str:
    """同梱済みのchromedriver.exeを探索して返す"""
    candidates = []
    try:
        if getattr(sys, "frozen", False):
            base_dir = Path(getattr(sys, "_MEIPASS", Path.cwd()))
        else:
            base_dir = Path(__file__).resolve().parent.parent

        candidates.extend([
            base_dir / "drivers" / "chromedriver-win32" / "chromedriver.exe",
            base_dir / "chromedriver.exe",
        ])

        for candidate in candidates:
            if candidate.exists() and candidate.is_file():
                return str(candidate)
    except Exception:
        pass

    return ""


def _extract_major_version(version_text: str):
    try:
        match = re.search(r"(\d+)\.", version_text or "")
        if match:
            return int(match.group(1))
    except Exception:
        pass
    return None


def _get_binary_version_output(executable_path: str) -> str:
    try:
        if not executable_path:
            return ""
        output = subprocess.check_output(
            [executable_path, "--version"],
            stderr=subprocess.STDOUT,
            text=True,
            timeout=3
        )
        return (output or "").strip()
    except Exception:
        return ""


def _find_chrome_binary() -> str:
    candidates = [
        os.environ.get("CHROME_PATH", ""),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return ""


def _get_chrome_major_version() -> int:
    chrome_binary = _find_chrome_binary()
    version_text = _get_binary_version_output(chrome_binary)
    return _extract_major_version(version_text)


def _get_chromedriver_major_version(chromedriver_path: str) -> int:
    version_text = _get_binary_version_output(chromedriver_path)
    return _extract_major_version(version_text)


def create_driver(headless=False, page_load_strategy: str = "normal"):
    """
    Chrome WebDriverを作成する
    
    Args:
        headless (bool): ヘッドレスモードで実行するかどうか
        
    Returns:
        WebDriver: 作成されたWebDriverインスタンス
    """
    try:
        logging.getLogger("WDM").setLevel(logging.WARNING)
        logging.getLogger("webdriver_manager").setLevel(logging.WARNING)

        # キャンセルチェック（ドライバー作成開始時）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass  # area_searchモジュールが利用できない場合はスキップ
        
        # Chromeオプションの設定
        chrome_options = Options()
        if page_load_strategy in ("normal", "eager", "none"):
            chrome_options.page_load_strategy = page_load_strategy
        if headless:
            chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-save-password-bubble')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--disable-web-security')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--window-size=1920,1080')
        
        # 画像を無効化
        prefs = {
            'profile.managed_default_content_settings.images': 2,
            'profile.default_content_setting_values.images': 2
        }
        chrome_options.add_experimental_option('prefs', prefs)
        
        # キャンセルチェック（オプション設定後）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        driver = None

        # 同梱ドライバ（バージョン一致時のみ）→ Selenium Manager の順で起動
        try:
            bundled_driver = _find_bundled_chromedriver()
            if bundled_driver:
                chrome_major = _get_chrome_major_version()
                driver_major = _get_chromedriver_major_version(bundled_driver)
                if chrome_major and driver_major and chrome_major == driver_major:
                    service = Service(executable_path=bundled_driver)
                    logging.info(f"同梱ChromeDriverを使用: {bundled_driver}")
                    driver = webdriver.Chrome(service=service, options=chrome_options)
                else:
                    logging.info(
                        f"同梱ChromeDriverをスキップ: browser={chrome_major}, driver={driver_major}"
                    )

            if driver is None:
                driver = webdriver.Chrome(options=chrome_options)
        except Exception as manager_error:
            logging.warning(f"Chrome起動で例外。Selenium Managerへフォールバックします: {str(manager_error)}")
            driver = webdriver.Chrome(options=chrome_options)
        
        # キャンセルチェック（ドライバー起動直前）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # キャンセルチェック（ドライバー作成直後）
        try:
            from services.area_search import check_cancellation
            check_cancellation()
        except (ImportError, NameError):
            pass
        
        # タイムアウト設定
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(10)
        
        return driver
        
    except Exception as e:
        logging.error(f"WebDriverの作成に失敗: {str(e)}")
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