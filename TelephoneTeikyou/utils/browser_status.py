"""
ブラウザステータス管理

このモジュールは、ブラウザドライバーの状態を
グローバルに管理するための機能を提供します。
"""

class BrowserStatus:
    """ブラウザの状態を管理するクラス"""
    
    def __init__(self):
        """初期化"""
        self.driver = None
        self.is_initialized = False
    
    def set_driver(self, driver):
        """
        ドライバーを設定する
        
        Args:
            driver: Seleniumのドライバーインスタンス
        """
        self.driver = driver
        self.is_initialized = True
    
    def clear_driver(self):
        """ドライバーをクリアする"""
        self.driver = None
        self.is_initialized = False

# グローバルなインスタンスを作成
browser_status = BrowserStatus() 