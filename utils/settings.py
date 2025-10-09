"""
設定機能を提供するモジュール

このモジュールは、アプリケーションの設定を管理する機能を提供します。
"""

import os
import json
import logging
from typing import Dict, Any, Optional

def _candidate_paths(default_name: str = "settings.json") -> list[str]:
    """設定ファイル探索候補を返す（exe直下→CWD→ソース直下の順）。

    - settings.json / setteings.json の両方を候補に含める
    - 先に見つかったものを優先
    """
    paths: list[str] = []
    try:
        # exe と同じフォルダ（ビルド後はここを最優先）
        exe_dir = os.path.dirname(os.path.abspath(os.path.expandvars(os.path.expanduser(os.sys.argv[0]))))
        paths.append(os.path.join(exe_dir, "settings.json"))
        paths.append(os.path.join(exe_dir, "setteings.json"))
    except Exception:
        pass
    try:
        # カレントディレクトリ
        cwd = os.getcwd()
        paths.append(os.path.join(cwd, "settings.json"))
        paths.append(os.path.join(cwd, "setteings.json"))
    except Exception:
        pass
    try:
        # このファイルのあるソース直下
        here = os.path.dirname(os.path.abspath(__file__))
        root = os.path.dirname(here)
        paths.append(os.path.join(root, "settings.json"))
        paths.append(os.path.join(root, "setteings.json"))
    except Exception:
        pass
    # 重複排除を保持順で
    seen = set()
    uniq: list[str] = []
    for p in paths:
        if p not in seen:
            uniq.append(p)
            seen.add(p)
    return uniq

class Settings:
    """設定を管理するクラス"""
    
    def __init__(self, settings_file: str = "settings.json"):
        """初期化"""
        self.settings_file = settings_file
        self.settings: Dict[str, Any] = {}
        self.load_settings()
    
    def load_settings(self) -> None:
        """設定を読み込む（JSONファイルのみ）。

        - 探索順: exe直下 → CWD → ソース直下
        - settings.json / setteings.json の両対応
        - 見つからなければ空設定のまま（自動保存しない）
        """
        try:
            # 呼び出し時に明示パスが指定されている場合はそれを最優先
            candidate_list = []
            if self.settings_file and os.path.isabs(self.settings_file):
                candidate_list.append(self.settings_file)
            # 既定探索候補
            candidate_list += _candidate_paths()

            loaded_path: Optional[str] = None
            for p in candidate_list:
                try:
                    if os.path.exists(p):
                        with open(p, 'r', encoding='utf-8') as f:
                            self.settings = json.load(f)
                        loaded_path = p
                        # 読み込みに成功したパスを以後の保存先にする
                        self.settings_file = p
                        break
                except Exception:
                    continue

            if loaded_path:
                logging.info(f"設定ファイルを読み込みました: {loaded_path}")
            else:
                # 見つからない場合は空設定のまま
                self.settings = {}
                logging.warning("設定ファイルが見つかりませんでした（空設定を使用します）。")
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