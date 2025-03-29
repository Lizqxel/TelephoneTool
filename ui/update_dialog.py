"""
アップデート設定ダイアログ

このモジュールは、アプリケーションのアップデート設定と
アップデート履歴を表示するダイアログを提供します。
"""

import os
import sys
import json
import logging
import requests
import subprocess
import glob
from datetime import datetime
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QCheckBox, QSpinBox, QComboBox,
    QTableWidget, QTableWidgetItem, QMessageBox,
                              QProgressDialog, QApplication, QHeaderView, QGroupBox)
from PySide6.QtCore import Qt, QTimer
from version import VERSION, GITHUB_OWNER, GITHUB_REPO

class UpdateDialog(QDialog):
    """アップデート設定ダイアログ"""
    
    def __init__(self, parent=None):
        """ダイアログの初期化"""
        super().__init__(parent)
        self.setWindowTitle("アップデート履歴")
        self.setMinimumWidth(400)
        
        # 設定ファイルのパスと設定の初期化
        self.settings_file = "settings.json"
        self.settings = {}
        self.load_settings()  # 設定を読み込む
        
        # メインレイアウト
        layout = QVBoxLayout(self)
        
        # アップデート設定グループ
        settings_group = QGroupBox("アップデート設定")
        settings_layout = QVBoxLayout()
        
        # 自動チェック設定
        self.auto_check = QCheckBox("起動時に自動チェック")
        self.auto_check.setChecked(self.settings.get("update_settings", {}).get("auto_check", True))
        settings_layout.addWidget(self.auto_check)
        
        # チェック間隔設定
        interval_layout = QHBoxLayout()
        interval_layout.addWidget(QLabel("チェック間隔:"))
        self.check_interval = QSpinBox()
        self.check_interval.setRange(1, 30)  # 1-30日
        self.check_interval.setValue(self.settings.get("update_settings", {}).get("check_interval", 86400) // 86400)  # 秒を日に変換
        self.check_interval.setSuffix("日")
        interval_layout.addWidget(self.check_interval)
        interval_layout.addStretch()
        settings_layout.addLayout(interval_layout)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # アップデート履歴テーブル
        history_label = QLabel("アップデート履歴")
        layout.addWidget(history_label)
        
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(3)
        self.history_table.setHorizontalHeaderLabels(["バージョン", "チェック日時", "状態"])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.history_table)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        # 今すぐチェックボタン
        check_button = QPushButton("今すぐチェック")
        check_button.clicked.connect(self.check_for_updates)
        button_layout.addWidget(check_button)
        
        # 設定保存ボタン
        save_button = QPushButton("設定を保存")
        save_button.clicked.connect(self.save_settings)
        button_layout.addWidget(save_button)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(self.accept)
        button_layout.addWidget(close_button)
        
        layout.addLayout(button_layout)
        
        # アップデート履歴を読み込む
        self.load_update_history()
    
    def load_update_history(self):
        """アップデート履歴を読み込む"""
        try:
            # 設定ファイルを読み込む
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                self.settings = {}
            
            # 履歴を読み込む
            history = self.settings.get("update_history", [])
            self.history_table.setRowCount(len(history))
            
            for i, entry in enumerate(history):
                self.history_table.setItem(i, 0, QTableWidgetItem(entry.get("version", "")))
                self.history_table.setItem(i, 1, QTableWidgetItem(entry.get("check_time", "")))
                self.history_table.setItem(i, 2, QTableWidgetItem(entry.get("status", "")))
        except Exception as e:
            logging.error(f"アップデート履歴の読み込み中にエラー: {e}")
    
    def check_for_updates(self):
        """アップデートをチェックする"""
        try:
            # GitHubのAPIを使用して最新リリースを取得
            url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
            response = requests.get(url)
            response.raise_for_status()
            latest_release = response.json()
            
            latest_version = latest_release["tag_name"].lstrip("v")
            current_version = VERSION
            
            if latest_version > current_version:
                # 新しいバージョンが利用可能
                msg = f"新しいバージョン v{latest_version} が利用可能です。\n"
                msg += f"現在のバージョン: v{current_version}\n\n"
                msg += "更新しますか？"
                
                reply = QMessageBox.question(self, "アップデート", msg,
                                          QMessageBox.StandardButton.Yes |
                                          QMessageBox.StandardButton.No)
                
                if reply == QMessageBox.StandardButton.Yes:
                    self.download_and_apply_update(latest_release)
                else:
                    self.add_history_entry(latest_version, "更新をスキップ")
            else:
                msg = "現在のバージョンが最新です。"
                QMessageBox.information(self, "アップデート", msg)
                self.add_history_entry(current_version, "最新版")
            
        except Exception as e:
            logging.error(f"アップデートチェック中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"アップデートのチェック中にエラーが発生しました: {e}")
        
    def load_settings(self):
        """設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
            else:
                self.settings = {}
        except Exception as e:
            logging.error(f"設定の読み込み中にエラー: {e}")
            self.settings = {}
        
    def save_settings(self):
        """設定を保存"""
        try:
            # update_settings内に設定を保存
            if "update_settings" not in self.settings:
                self.settings["update_settings"] = {}
            
            self.settings["update_settings"].update({
                "auto_check": self.auto_check.isChecked(),
                "check_interval": self.check_interval.value() * 86400  # 日を秒に変換
            })
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            
            QMessageBox.information(self, "成功", "設定を保存しました")
        except Exception as e:
            logging.error(f"設定の保存中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"設定の保存中にエラーが発生しました: {e}")
    
    def load_history(self):
        """アップデート履歴を読み込む"""
        try:
            history = self.settings.get("update_history", [])
            self.history_table.setRowCount(len(history))
            
            for i, entry in enumerate(history):
                self.history_table.setItem(i, 0, QTableWidgetItem(entry.get("version", "")))
                self.history_table.setItem(i, 1, QTableWidgetItem(entry.get("check_time", "")))
                self.history_table.setItem(i, 2, QTableWidgetItem(entry.get("status", "")))
        except Exception as e:
            logging.error(f"履歴の読み込み中にエラー: {e}")
            self.history_table.setRowCount(0)  # テーブルをクリア
    
    def add_history_entry(self, version, status):
        """履歴エントリを追加"""
        try:
            history = self.settings.get("update_history", [])
            history.insert(0, {
                "version": version,
                "check_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "status": status
            })
            self.settings["update_history"] = history
            self.save_settings()
            self.load_history()
        except Exception as e:
            logging.error(f"履歴の追加中にエラー: {e}")
    
    def download_and_apply_update(self, release):
        """アップデートをダウンロードして適用"""
        try:
            # プログレスダイアログの作成
            progress = QProgressDialog("アップデートをダウンロード中...", "キャンセル", 0, 100, self)
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setAutoReset(False)
            progress.setAutoClose(False)
            progress.show()
            
            # アセットのダウンロード
            assets = release.get("assets", [])
            if not assets:
                raise Exception("ダウンロード可能なファイルが見つかりません")
            
            # exeファイルを探す
            exe_asset = next((asset for asset in assets if asset["name"].endswith(".exe")), None)
            if not exe_asset:
                raise Exception("実行可能ファイルが見つかりません")
            
            # ダウンロードURL
            download_url = exe_asset["browser_download_url"]
            new_version = release["tag_name"].lstrip("v")
            
            # ダウンロード
            response = requests.get(download_url, stream=True)
            response.raise_for_status()
            
            # ファイルサイズの取得
            total_size = int(response.headers.get("content-length", 0))
            block_size = 8192
            downloaded = 0
            
            # 新しいバージョンを保存
            new_exe_path = f"TelephoneTool-{new_version}.exe"
            with open(new_exe_path, "wb") as f:
                for data in response.iter_content(block_size):
                    downloaded += len(data)
                    f.write(data)
                    if total_size:
                        progress_value = int((downloaded / total_size) * 100)
                        progress.setValue(progress_value)
                        if progress.wasCanceled():
                            raise Exception("ダウンロードがキャンセルされました")
            
            progress.setLabelText("アップデートを適用中...")
            progress.setValue(100)
            
            # 現在の実行ファイルのパスを取得
            current_exe = os.path.abspath(sys.executable)
            current_dir = os.path.dirname(current_exe)
            current_pid = os.getpid()
            
            # バッチファイルを作成
            batch_path = os.path.join(current_dir, "update.bat")
            with open(batch_path, "w", encoding="shift-jis") as f:
                f.write("@echo off\n")
                f.write("setlocal enabledelayedexpansion\n")
                f.write("cd /d %~dp0\n")
                
                # プロセスを終了
                f.write(f'taskkill /F /PID {current_pid} /T > nul 2>&1\n')
                f.write("timeout /t 0.5 /nobreak > nul\n")
                
                # 新しいバージョンを配置
                f.write(f'copy /Y "{new_exe_path}" "{new_exe_path}.tmp" > nul 2>&1\n')
                
                # すべての古いバージョンを削除（現在実行中のファイルを除く）
                f.write('for %%f in (TelephoneTool-*.exe) do (\n')
                f.write(f'    if /I not "%%f"=="{os.path.basename(new_exe_path)}" (\n')
                f.write('        del "%%f" > nul 2>&1\n')
                f.write('    )\n')
                f.write(')\n')
                
                # 一時ファイルを正しい名前に変更
                f.write(f'move /Y "{new_exe_path}.tmp" "{new_exe_path}" > nul 2>&1\n')
                
                # バッチファイル自身を削除
                f.write("timeout /t 0.5 /nobreak > nul\n")
                f.write("del %~f0 > nul 2>&1\n")
            
            # 履歴に追加
            self.add_history_entry(new_version, "更新完了")
            
            # バッチファイルを非表示で実行
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            # cmdを使用してバッチファイルを非表示で実行
            subprocess.Popen(['cmd', '/c', batch_path],
                           startupinfo=startupinfo,
                           creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS)
            
            # アプリケーションを終了する前に少し待機
            QTimer.singleShot(500, lambda: os._exit(0))
                
        except Exception as e:
            logging.error(f"アップデート中にエラー: {e}")
            QMessageBox.critical(self, "エラー", f"アップデート中にエラーが発生しました: {e}")
            
            # 一時ファイルの削除
            for file in [new_exe_path, f"{new_exe_path}.tmp", batch_path]:
                if os.path.exists(file):
                    try:
                        os.remove(file)
                    except Exception as e:
                        logging.error(f"一時ファイルの削除中にエラー: {e}")
                        pass 