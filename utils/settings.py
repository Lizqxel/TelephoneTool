"""
設定機能を提供するモジュール

このモジュールは、アプリケーションの設定を管理する機能を提供します。
"""

import os
import json
import logging
from typing import Dict, Any, Optional

class Settings:
    """設定を管理するクラス"""
    
    def __init__(self, settings_file: str = "settings.json"):
        """初期化"""
        self.settings_file = settings_file
        self.settings: Dict[str, Any] = {}
        self.load_settings()
    
    def load_settings(self) -> None:
        """設定を読み込む（settings.json / setteings.json の両対応）"""
        try:
            # まずは指定されたパスを試す
            target_path = self.settings_file

            # 指定が settings.json で存在しない場合、同階層の setteings.json をフォールバック
            try:
                base = os.path.basename(self.settings_file)
                dir_ = os.path.dirname(self.settings_file) or os.getcwd()
                if not os.path.exists(target_path) and base.lower() == 'settings.json':
                    alt = os.path.join(dir_, 'setteings.json')
                    if os.path.exists(alt):
                        target_path = alt
                        # 以後の保存先もフォールバック先に合わせる
                        self.settings_file = alt
            except Exception:
                pass

            if os.path.exists(target_path):
                with open(target_path, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                # どちらも存在しない場合は空設定として初期化し、指定パスに保存
                self.settings = {}
                self.save_settings()
        except Exception as e:
            logging.error(f"設定の読み込み中にエラー: {e}")
            self.settings = {}
    
    def save_settings(self) -> None:
        """設定を保存する"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"設定の保存中にエラー: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """設定値を取得する"""
        return self.settings.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """設定値を設定する"""
        self.settings[key] = value
        self.save_settings()
    
    def update(self, settings: Dict[str, Any]) -> None:
        """複数の設定値を更新する"""
        self.settings.update(settings)
        self.save_settings()

# グローバルな設定インスタンス
settings = Settings() 