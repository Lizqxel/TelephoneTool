"""
ブラウザオプション設定モジュール

Chromeブラウザのオプションを最適化し、GPU関連のエラーを回避するための設定を提供します。
"""

from selenium.webdriver.chrome.options import Options

def get_optimized_options(headless=False):
    """
    最適化されたChromeオプションを取得します
    
    Args:
        headless (bool): ヘッドレスモードを有効にするかどうか
        
    Returns:
        Options: 最適化されたChromeオプション
    """
    options = Options()
    
    # WebGLとハードウェアアクセラレーションに関する問題を回避
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--disable-dev-shm-usage')
    
    # クラッシュ回避のための設定
    options.add_argument('--no-sandbox')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--allow-running-insecure-content')
    
    # パフォーマンス最適化
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-infobars')
    
    # メモリ使用量の最適化
    options.add_argument('--disable-application-cache')
    options.add_argument('--disable-popup-blocking')
    
    # ヘッドレスモードの設定（必要な場合）
    if headless:
        options.add_argument('--headless=new')
    
    # ログレベルの設定
    options.add_argument('--log-level=3')  # FATALのみ
    
    # SwiftShaderのフォールバックを有効化（ソフトウェアレンダリング）
    options.add_argument('--enable-unsafe-swiftshader')
    
    # ウィンドウサイズの設定
    options.add_argument('--window-size=1280,800')
    
    return options 