"""
アップデート設定と履歴を表示するダイアログ

このモジュールは、アプリケーションのアップデート設定と
アップデート履歴を表示・管理するためのUIを提供します。
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QCheckBox, QSpinBox, QComboBox,
    QTableWidget, QTableWidgetItem, QMessageBox,
    QGroupBox, QFormLayout
)
from PySide6.QtCore import Qt, QTimer
from datetime import datetime
import json
from pathlib import Path

from utils.updater import UpdateChecker, Updater

class UpdateDialog(QDialog):
    """アップデート設定ダイアログ"""
    
    def __init__(self, parent=None):
        """初期化"""
        super().__init__(parent)
        self.setWindowTitle("アップデート設定")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        self.update_checker = UpdateChecker()
        self.updater = Updater()
        
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        """UIの初期化"""
        layout = QVBoxLayout(self)
        
        # アップデート設定グループ
        settings_group = QGroupBox("アップデート設定")
        settings_layout = QFormLayout()
        
        # 自動チェック
        self.auto_check = QCheckBox("自動的にアップデートをチェック")
        settings_layout.addRow("自動チェック:", self.auto_check)
        
        # チェック間隔
        interval_layout = QHBoxLayout()
        self.check_interval = QSpinBox()
        self.check_interval.setRange(3600, 604800)  # 1時間から1週間
        self.check_interval.setSingleStep(3600)  # 1時間単位
        interval_layout.addWidget(self.check_interval)
        interval_layout.addWidget(QLabel("秒"))
        settings_layout.addRow("チェック間隔:", interval_layout)
        
        # アップデートチャンネル
        self.update_channel = QComboBox()
        self.update_channel.addItems(["stable", "beta"])
        settings_layout.addRow("アップデートチャンネル:", self.update_channel)
        
        # 自動ダウンロード
        self.auto_download = QCheckBox("新バージョンを自動的にダウンロード")
        settings_layout.addRow("自動ダウンロード:", self.auto_download)
        
        # バックアップ
        self.backup_before_update = QCheckBox("アップデート前にバックアップを作成")
        settings_layout.addRow("バックアップ:", self.backup_before_update)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # アップデート履歴
        history_group = QGroupBox("アップデート履歴")
        history_layout = QVBoxLayout()
        
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(3)
        self.history_table.setHorizontalHeaderLabels(["バージョン", "チェック日時", "ステータス"])
        self.history_table.horizontalHeader().setStretchLastSection(True)
        history_layout.addWidget(self.history_table)
        
        history_group.setLayout(history_layout)
        layout.addWidget(history_group)
        
        # ボタン
        button_layout = QHBoxLayout()
        
        self.check_button = QPushButton("今すぐチェック")
        self.check_button.clicked.connect(self.check_updates)
        button_layout.addWidget(self.check_button)
        
        self.save_button = QPushButton("設定を保存")
        self.save_button.clicked.connect(self.save_settings)
        button_layout.addWidget(self.save_button)
        
        self.close_button = QPushButton("閉じる")
        self.close_button.clicked.connect(self.close)
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
        
    def load_settings(self):
        """設定を読み込む"""
        settings = self.update_checker.settings
        update_settings = settings.get("update_settings", {})
        
        self.auto_check.setChecked(update_settings.get("auto_check", True))
        self.check_interval.setValue(update_settings.get("check_interval", 86400))
        self.update_channel.setCurrentText(update_settings.get("update_channel", "stable"))
        self.auto_download.setChecked(update_settings.get("auto_download", False))
        self.backup_before_update.setChecked(update_settings.get("backup_before_update", True))
        
        # 履歴の表示
        self.update_history_table(update_settings.get("update_history", []))
        
    def save_settings(self):
        """設定を保存する"""
        settings = self.update_checker.settings
        update_settings = settings.get("update_settings", {})
        
        update_settings.update({
            "auto_check": self.auto_check.isChecked(),
            "check_interval": self.check_interval.value(),
            "update_channel": self.update_channel.currentText(),
            "auto_download": self.auto_download.isChecked(),
            "backup_before_update": self.backup_before_update.isChecked()
        })
        
        settings["update_settings"] = update_settings
        self.update_checker.settings = settings
        self.update_checker.save_settings()
        
        QMessageBox.information(self, "設定保存", "設定を保存しました。")
        
    def update_history_table(self, history):
        """履歴テーブルを更新する"""
        self.history_table.setRowCount(len(history))
        
        for i, entry in enumerate(history):
            version = QTableWidgetItem(entry.get("version", ""))
            check_time = QTableWidgetItem(
                datetime.fromisoformat(entry.get("check_time", "")).strftime("%Y-%m-%d %H:%M:%S")
            )
            status = QTableWidgetItem(entry.get("status", ""))
            
            self.history_table.setItem(i, 0, version)
            self.history_table.setItem(i, 1, check_time)
            self.history_table.setItem(i, 2, status)
            
    def check_updates(self):
        """アップデートをチェックする"""
        self.check_button.setEnabled(False)
        self.check_button.setText("チェック中...")
        
        # 非同期でチェックを実行
        QTimer.singleShot(100, self._do_check_updates)
        
    def _do_check_updates(self):
        """実際のアップデートチェックを実行"""
        try:
            latest_version, download_url = self.update_checker.check_for_updates()
            
            if latest_version:
                msg = f"新しいバージョン {latest_version} が利用可能です。\n"
                msg += "アップデートをダウンロードしますか？"
                
                reply = QMessageBox.question(
                    self, "アップデート利用可能",
                    msg,
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    self.download_and_apply_update(download_url)
            else:
                QMessageBox.information(
                    self, "アップデート確認",
                    "現在のバージョンが最新です。"
                )
        except Exception as e:
            QMessageBox.warning(
                self, "エラー",
                f"アップデートチェック中にエラーが発生しました：\n{str(e)}"
            )
        finally:
            self.check_button.setEnabled(True)
            self.check_button.setText("今すぐチェック")
            self.load_settings()  # 履歴を更新
            
    def download_and_apply_update(self, download_url):
        """アップデートをダウンロードして適用する"""
        try:
            # ダウンロード
            file_path = self.updater.download_update(
                download_url,
                callback=self._update_progress
            )
            
            if file_path:
                # アップデートの適用
                if self.updater.apply_update(file_path):
                    QMessageBox.information(
                        self, "アップデート",
                        "アップデートを適用するためにアプリケーションを再起動します。"
                    )
                else:
                    raise Exception("アップデートの適用に失敗しました。")
            else:
                raise Exception("アップデートファイルのダウンロードに失敗しました。")
                
        except Exception as e:
            QMessageBox.warning(
                self, "エラー",
                f"アップデート中にエラーが発生しました：\n{str(e)}"
            )
            
    def _update_progress(self, progress):
        """ダウンロード進捗を更新する"""
        self.check_button.setText(f"ダウンロード中... {progress}%") 