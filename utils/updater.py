"""
アプリケーションの自動アップデート機能を提供するモジュール

このモジュールは、GitHubのリリース機能を利用して、
アプリケーションの最新バージョンの確認とアップデートを行います。

主な機能：
- GitHubからの最新バージョン取得
- バージョン比較
- アップデートファイルのダウンロード
- アプリケーションの更新
"""

import os
import sys
import json
import logging
import tempfile
import requests
from pathlib import Path
from packaging import version
from datetime import datetime, timedelta

from version import VERSION, APP_NAME, GITHUB_OWNER, GITHUB_REPO, UPDATE_CHECK_INTERVAL

logger = logging.getLogger(__name__)

class UpdateChecker:
    """アップデートチェックを行うクラス"""
    
    def __init__(self):
        """初期化"""
        self.settings_file = Path("settings.json")
        self.default_update_settings = {
            "auto_check": True,
            "check_interval": 86400,
            "last_update_check": None,
            "skip_version": None,
            "update_channel": "stable",
            "auto_download": False,
            "backup_before_update": True,
            "update_history": []
        }
        self.load_settings()

    def load_settings(self):
        """設定ファイルを読み込む"""
        try:
            if self.settings_file.exists():
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    self.settings = json.load(f)
                
                # アップデート設定が存在しない場合は追加
                if "update_settings" not in self.settings:
                    self.settings["update_settings"] = self.default_update_settings
                    self.save_settings()
            else:
                self.settings = {
                    "update_settings": self.default_update_settings
                }
                self.save_settings()
        except Exception as e:
            logger.error(f"設定ファイルの読み込みに失敗しました: {e}")
            self.settings = {
                "update_settings": self.default_update_settings
            }

    def save_settings(self):
        """設定ファイルを保存する"""
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logger.error(f"設定ファイルの保存に失敗しました: {e}")

    def should_check_update(self):
        """アップデートチェックが必要かどうかを判定する"""
        update_settings = self.settings.get("update_settings", {})
        if not update_settings.get("auto_check", True):
            return False
        
        last_check = update_settings.get("last_update_check")
        if not last_check:
            return True
        
        check_interval = update_settings.get("check_interval", 86400)
        last_check_time = datetime.fromisoformat(last_check)
        next_check_time = last_check_time + timedelta(seconds=check_interval)
        return datetime.now() >= next_check_time

    def get_latest_release(self):
        """GitHubから最新のリリース情報を取得する"""
        try:
            update_settings = self.settings.get("update_settings", {})
            channel = update_settings.get("update_channel", "stable")
            
            # チャンネルに応じてURLを変更
            if channel == "stable":
                url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            else:
                url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
            
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            if channel == "stable":
                return response.json()
            else:
                # ベータ版の場合、最新のリリースを取得
                releases = response.json()
                if releases:
                    return releases[0]
                return None
                
        except Exception as e:
            logger.error(f"最新リリース情報の取得に失敗しました: {e}")
            return None

    def check_for_updates(self):
        """アップデートをチェックする"""
        if not self.should_check_update():
            return None, None

        # 最新リリースの取得
        latest_release = self.get_latest_release()
        if not latest_release:
            return None, None

        # バージョン比較
        latest_version = latest_release["tag_name"].lstrip("v")
        current_version = VERSION
        
        # スキップバージョンのチェック
        update_settings = self.settings.get("update_settings", {})
        skip_version = update_settings.get("skip_version")
        if skip_version and latest_version == skip_version:
            return None, None
        
        try:
            if version.parse(latest_version) > version.parse(current_version):
                # アップデート履歴に追加
                update_history = update_settings.get("update_history", [])
                update_history.append({
                    "version": latest_version,
                    "check_time": datetime.now().isoformat(),
                    "status": "available"
                })
                update_settings["update_history"] = update_history[-10:]  # 最新10件のみ保持
                update_settings["last_update_check"] = datetime.now().isoformat()
                self.save_settings()
                
                return latest_version, latest_release["assets"][0]["browser_download_url"]
        except Exception as e:
            logger.error(f"バージョン比較に失敗しました: {e}")
        
        return None, None

class Updater:
    """アップデートを実行するクラス"""
    
    def __init__(self):
        """初期化"""
        self.temp_dir = Path(tempfile.gettempdir()) / APP_NAME
        self.temp_dir.mkdir(exist_ok=True)
        self.settings_file = Path("settings.json")

    def download_update(self, url, callback=None):
        """アップデートファイルをダウンロードする"""
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()
            
            # ファイルサイズの取得
            total_size = int(response.headers.get("content-length", 0))
            
            # ダウンロードファイルのパス
            file_name = url.split("/")[-1]
            download_path = self.temp_dir / file_name
            
            # ダウンロード
            with open(download_path, "wb") as f:
                if total_size == 0:
                    f.write(response.content)
                else:
                    downloaded = 0
                    for data in response.iter_content(chunk_size=8192):
                        downloaded += len(data)
                        f.write(data)
                        if callback:
                            progress = int((downloaded / total_size) * 100)
                            callback(progress)
            
            return download_path
        except Exception as e:
            logger.error(f"アップデートファイルのダウンロードに失敗しました: {e}")
            return None

    def backup_current_version(self):
        """現在のバージョンをバックアップする"""
        try:
            # 設定ファイルを読み込む
            with open(self.settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
            
            update_settings = settings.get("update_settings", {})
            if not update_settings.get("backup_before_update", True):
                return True
            
            # バックアップディレクトリの作成
            backup_dir = Path("backups") / datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir.mkdir(parents=True, exist_ok=True)
            
            # 必要なファイルをコピー
            files_to_backup = [
                "main.py",
                "settings.json",
                "version.py",
                "requirements.txt",
                "utils",
                "ui",
                "services"
            ]
            
            for item in files_to_backup:
                src = Path(item)
                if src.exists():
                    if src.is_file():
                        shutil.copy2(src, backup_dir / src.name)
                    else:
                        shutil.copytree(src, backup_dir / src.name)
            
            # バックアップ情報を記録
            update_settings["last_backup"] = backup_dir.name
            settings["update_settings"] = update_settings
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            
            return True
        except Exception as e:
            logger.error(f"バックアップの作成に失敗しました: {e}")
            return False

    def apply_update(self, file_path):
        """アップデートを適用する"""
        try:
            # バックアップの作成
            if not self.backup_current_version():
                logger.error("バックアップの作成に失敗しました")
                return False
            
            # アップデートスクリプトの作成
            update_script = self.create_update_script(file_path)
            
            # アップデートスクリプトを実行
            if sys.platform == "win32":
                os.startfile(update_script)
            else:
                os.system(f"python {update_script} &")
            
            # 現在のプロセスを終了
            sys.exit(0)
        except Exception as e:
            logger.error(f"アップデートの適用に失敗しました: {e}")
            return False

    def create_update_script(self, file_path):
        """アップデート用のスクリプトを作成する"""
        script_path = self.temp_dir / "update.py"
        
        script_content = f'''
import os
import sys
import time
import shutil
from pathlib import Path

# 元のプロセスが終了するのを待つ
time.sleep(1)

# アプリケーションの更新
try:
    # 更新ファイルの展開
    update_file = Path("{file_path}")
    app_dir = Path("{os.getcwd()}")
    
    # ファイルの更新
    shutil.unpack_archive(update_file, app_dir)
    
    # 一時ファイルの削除
    update_file.unlink()
    
    # アプリケーションの再起動
    os.startfile(app_dir / "main.exe")
except Exception as e:
    with open("update_error.log", "w") as f:
        f.write(f"アップデートに失敗しました: {{e}}")
'''
        
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(script_content)
        
        return script_path 