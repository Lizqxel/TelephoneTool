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
import shutil
import logging
import tempfile
import requests
from pathlib import Path
from packaging import version
from datetime import datetime, timedelta
from PySide6.QtWidgets import QApplication

from version import VERSION, APP_NAME, GITHUB_OWNER, GITHUB_REPO, UPDATE_CHECK_INTERVAL
from ui.update_progress_dialog import UpdateProgressDialog

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
        self.progress_dialog = None

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
            
            # 現在のexeファイルをバックアップ
            current_exe = Path(sys.executable)
            backup_exe = backup_dir / current_exe.name
            shutil.copy2(current_exe, backup_exe)
            
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
            # プログレスダイアログの表示
            self.progress_dialog = UpdateProgressDialog()
            self.progress_dialog.show()
            
            # バックアップの作成
            self.progress_dialog.update_progress(10, "バックアップを作成中...")
            if not self.backup_current_version():
                raise Exception("バックアップの作成に失敗しました")
            
            # 現在のexeファイルのパスを取得
            current_exe = Path(sys.executable)
            new_exe = Path(file_path)
            
            # 一時ファイルとして新しいexeを配置
            self.progress_dialog.update_progress(30, "新しいバージョンを準備中...")
            temp_dir = Path(tempfile.gettempdir()) / APP_NAME
            temp_dir.mkdir(exist_ok=True)
            temp_exe = temp_dir / "TelephoneTool_new.exe"
            
            # 新しいexeを一時ディレクトリにコピー
            shutil.copy2(new_exe, temp_exe)
            
            # 更新スクリプトの作成
            self.progress_dialog.update_progress(50, "アップデートスクリプトを作成中...")
            update_script = self.create_update_script(current_exe, temp_exe)
            
            # 更新準備完了
            self.progress_dialog.update_progress(100, "アップデートの準備が完了しました")
            
            # 現在のプロセスを終了し、更新スクリプトを実行
            if sys.platform == "win32":
                os.startfile(update_script)
            else:
                os.system(f"python {update_script} &")
            
            sys.exit(0)
        except Exception as e:
            logger.error(f"アップデートの適用に失敗しました: {e}")
            if self.progress_dialog:
                self.progress_dialog.show_error(str(e))
            return False

    def create_update_script(self, current_exe, new_exe):
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

try:
    # 現在のexeファイルのパス
    current_exe = Path("{current_exe}")
    # 新しいexeファイルのパス
    new_exe = Path("{new_exe}")
    
    # 現在のexeファイルをバックアップ
    backup_exe = current_exe.with_suffix(".exe.bak")
    shutil.copy2(current_exe, backup_exe)
    
    # 新しいexeファイルで置き換え
    shutil.copy2(new_exe, current_exe)
    
    # 一時ファイルの削除
    new_exe.unlink()
    
    # アプリケーションの再起動
    os.startfile(current_exe)
except Exception as e:
    with open("update_error.log", "w") as f:
        f.write(f"アップデートに失敗しました: {{e}}")
    # エラー時はバックアップから復元
    if backup_exe.exists():
        shutil.copy2(backup_exe, current_exe)
'''
        
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(script_content)
        
        return script_path 