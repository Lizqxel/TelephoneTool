"""
スプレッドシート管理ダイアログ

このモジュールは、Googleフォーム転記先のスプレッドシート設定を
管理するためのダイアログUIを提供します。

概要:
- スプレッドシートの追加・編集・削除機能
- ルーティングキーとラベルの管理
- 設定ファイルへの保存・読み込み

制限事項:
- ルーティングキーは一意である必要があります
- 削除時は確認ダイアログを表示します
"""

from __future__ import annotations

import json
import os
import sys
import logging
from typing import Dict, Any, List, Optional
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QPushButton, QMessageBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QGroupBox, QComboBox, QSpinBox, QCheckBox
)
from PySide6.QtCore import Qt, Signal


class SpreadsheetManagementDialog(QDialog):
    """スプレッドシート管理ダイアログクラス"""
    
    # 設定が変更されたことを通知するシグナル
    settings_changed = Signal()
    
    def __init__(self, parent=None) -> None:
        """
        スプレッドシート管理ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
        """
        super().__init__(parent)
        self.setWindowTitle("スプレッドシート管理")
        self.setFixedSize(800, 600)
        
        # 設定ファイルのパスを設定
        self.settings_file = self._get_settings_file_path()
        
        # 現在の設定を読み込み
        self.current_settings = self._load_settings()
        self.destinations = self.current_settings.get("googleFormPosting", {}).get("destinations", [])
        
        # UIの構築
        self._setup_ui()
        
        # テーブルにデータを読み込み
        self._load_destinations_to_table()
    
    def _get_settings_file_path(self) -> str:
        """設定ファイルのパスを取得する"""
        if getattr(sys, 'frozen', False):
            # exeファイルとして実行されている場合
            return os.path.join(os.path.dirname(sys.executable), 'settings.json')
        else:
            # 通常のPythonスクリプトとして実行されている場合
            return os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings.json')
    
    def _load_settings(self) -> Dict[str, Any]:
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {}
        except Exception as e:
            logging.error(f"設定ファイルの読み込みに失敗しました: {e}")
            return {}
    
    def _save_settings(self) -> bool:
        """設定をファイルに保存する"""
        try:
            # googleFormPostingセクションを更新
            if "googleFormPosting" not in self.current_settings:
                self.current_settings["googleFormPosting"] = {}
            
            self.current_settings["googleFormPosting"]["destinations"] = self.destinations
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.current_settings, f, ensure_ascii=False, indent=2)
            
            return True
        except Exception as e:
            logging.error(f"設定ファイルの保存に失敗しました: {e}")
            QMessageBox.critical(self, "エラー", f"設定の保存に失敗しました: {str(e)}")
            return False
    
    def _setup_ui(self) -> None:
        """UIを構築する"""
        layout = QVBoxLayout(self)
        
        # 説明ラベル
        description = QLabel(
            "Googleフォーム転記先のスプレッドシート設定を管理します。\n"
            "ルーティングキーは一意である必要があります。"
        )
        description.setWordWrap(True)
        layout.addWidget(description)
        
        # テーブルウィジェット
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ラベル", "ルーティングキー", "スプレッドシートID", "シート名", "操作"])
        
        # テーブルの設定
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)  # ラベル列を伸縮
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # ルーティングキー列を内容に合わせる
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)  # スプレッドシートID列を伸縮
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)  # シート名列を内容に合わせる
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)  # 操作列を内容に合わせる
        
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.table)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        # 追加ボタン
        self.add_btn = QPushButton("追加")
        self.add_btn.clicked.connect(self._add_destination)
        button_layout.addWidget(self.add_btn)
        
        # 編集ボタン
        self.edit_btn = QPushButton("編集")
        self.edit_btn.clicked.connect(self._edit_destination)
        self.edit_btn.setEnabled(False)
        button_layout.addWidget(self.edit_btn)
        
        # 削除ボタン
        self.delete_btn = QPushButton("削除")
        self.delete_btn.clicked.connect(self._delete_destination)
        self.delete_btn.setEnabled(False)
        button_layout.addWidget(self.delete_btn)
        
        button_layout.addStretch()
        
        # 保存ボタン
        self.save_btn = QPushButton("保存")
        self.save_btn.clicked.connect(self._save_and_close)
        button_layout.addWidget(self.save_btn)
        
        # キャンセルボタン
        self.cancel_btn = QPushButton("キャンセル")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        # テーブルの選択変更イベント
        self.table.itemSelectionChanged.connect(self._on_selection_changed)
    
    def _load_destinations_to_table(self) -> None:
        """destinationsをテーブルに読み込む"""
        self.table.setRowCount(len(self.destinations))
        
        for row, dest in enumerate(self.destinations):
            # ラベル
            label_item = QTableWidgetItem(dest.get("label", ""))
            self.table.setItem(row, 0, label_item)
            
            # ルーティングキー
            route_key_item = QTableWidgetItem(dest.get("routeKey", ""))
            self.table.setItem(row, 1, route_key_item)
            
            # スプレッドシートID
            spreadsheet_id_item = QTableWidgetItem(dest.get("spreadsheetId", ""))
            self.table.setItem(row, 2, spreadsheet_id_item)
            
            # シート名
            sheet_name_item = QTableWidgetItem(dest.get("sheetName", ""))
            self.table.setItem(row, 3, sheet_name_item)
            
            # 操作ボタン
            edit_btn = QPushButton("編集")
            edit_btn.clicked.connect(lambda checked, r=row: self._edit_destination_at_row(r))
            self.table.setCellWidget(row, 4, edit_btn)
    
    def _on_selection_changed(self) -> None:
        """テーブルの選択が変更された時の処理"""
        has_selection = len(self.table.selectedItems()) > 0
        self.edit_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)
    
    def _add_destination(self) -> None:
        """新しい転記先を追加する"""
        dialog = DestinationEditDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_dest = dialog.get_destination()
            if self._validate_destination(new_dest):
                self.destinations.append(new_dest)
                self._load_destinations_to_table()
    
    def _edit_destination(self) -> None:
        """選択された転記先を編集する"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self._edit_destination_at_row(current_row)
    
    def _edit_destination_at_row(self, row: int) -> None:
        """指定された行の転記先を編集する"""
        if 0 <= row < len(self.destinations):
            dest = self.destinations[row]
            dialog = DestinationEditDialog(self, dest)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                new_dest = dialog.get_destination()
                if self._validate_destination(new_dest, exclude_index=row):
                    self.destinations[row] = new_dest
                    self._load_destinations_to_table()
    
    def _delete_destination(self) -> None:
        """選択された転記先を削除する"""
        current_row = self.table.currentRow()
        if current_row >= 0 and current_row < len(self.destinations):
            dest = self.destinations[current_row]
            reply = QMessageBox.question(
                self, 
                "削除確認", 
                f"「{dest.get('label', '')}」を削除しますか？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                del self.destinations[current_row]
                self._load_destinations_to_table()
    
    def _validate_destination(self, dest: Dict[str, str], exclude_index: Optional[int] = None) -> bool:
        """転記先の設定を検証する"""
        label = dest.get("label", "").strip()
        route_key = dest.get("routeKey", "").strip()
        spreadsheet_id = dest.get("spreadsheetId", "").strip()
        sheet_name = dest.get("sheetName", "").strip()
        
        # 必須項目チェック
        if not label:
            QMessageBox.warning(self, "入力エラー", "ラベルを入力してください。")
            return False
        
        if not route_key:
            QMessageBox.warning(self, "入力エラー", "ルーティングキーを入力してください。")
            return False
        
        if not spreadsheet_id:
            QMessageBox.warning(self, "入力エラー", "スプレッドシートIDを入力してください。")
            return False
        
        if not sheet_name:
            QMessageBox.warning(self, "入力エラー", "シート名を入力してください。")
            return False
        
        # ルーティングキーの重複チェック
        for i, existing_dest in enumerate(self.destinations):
            if exclude_index is not None and i == exclude_index:
                continue
            if existing_dest.get("routeKey", "").strip() == route_key:
                QMessageBox.warning(self, "入力エラー", f"ルーティングキー「{route_key}」は既に使用されています。")
                return False
        
        return True
    
    def _save_and_close(self) -> None:
        """設定を保存してダイアログを閉じる"""
        if self._save_settings():
            self.settings_changed.emit()
            self.accept()


class DestinationEditDialog(QDialog):
    """転記先編集ダイアログクラス"""
    
    def __init__(self, parent=None, destination: Optional[Dict[str, str]] = None) -> None:
        """
        転記先編集ダイアログの初期化
        
        Args:
            parent: 親ウィジェット
            destination: 編集する転記先の設定（Noneの場合は新規作成）
        """
        super().__init__(parent)
        self.setWindowTitle("転記先編集" if destination else "転記先追加")
        self.setFixedSize(400, 200)
        
        self.destination = destination or {}
        
        self._setup_ui()
        self._load_destination()
    
    def _setup_ui(self) -> None:
        """UIを構築する"""
        layout = QVBoxLayout(self)
        
        # ラベル入力
        label_layout = QHBoxLayout()
        label_layout.addWidget(QLabel("ラベル:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("表示名を入力してください")
        label_layout.addWidget(self.label_edit)
        layout.addLayout(label_layout)
        
        # ルーティングキー入力
        route_key_layout = QHBoxLayout()
        route_key_layout.addWidget(QLabel("ルーティングキー:"))
        self.route_key_edit = QLineEdit()
        self.route_key_edit.setPlaceholderText("一意のキーを入力してください")
        route_key_layout.addWidget(self.route_key_edit)
        layout.addLayout(route_key_layout)
        
        # スプレッドシートID入力
        spreadsheet_id_layout = QHBoxLayout()
        spreadsheet_id_layout.addWidget(QLabel("スプレッドシートID:"))
        self.spreadsheet_id_edit = QLineEdit()
        self.spreadsheet_id_edit.setPlaceholderText("スプレッドシートのIDを入力してください")
        spreadsheet_id_layout.addWidget(self.spreadsheet_id_edit)
        layout.addLayout(spreadsheet_id_layout)
        
        # シート名入力
        sheet_name_layout = QHBoxLayout()
        sheet_name_layout.addWidget(QLabel("シート名:"))
        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setPlaceholderText("シート名を入力してください")
        sheet_name_layout.addWidget(self.sheet_name_edit)
        layout.addLayout(sheet_name_layout)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # キャンセルボタン
        self.cancel_btn = QPushButton("キャンセル")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        # OKボタン
        self.ok_btn = QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)
        self.ok_btn.setDefault(True)
        button_layout.addWidget(self.ok_btn)
        
        layout.addLayout(button_layout)
    
    def _load_destination(self) -> None:
        """転記先の設定を読み込む"""
        self.label_edit.setText(self.destination.get("label", ""))
        self.route_key_edit.setText(self.destination.get("routeKey", ""))
        self.spreadsheet_id_edit.setText(self.destination.get("spreadsheetId", ""))
        self.sheet_name_edit.setText(self.destination.get("sheetName", ""))
    
    def get_destination(self) -> Dict[str, str]:
        """編集された転記先の設定を取得する"""
        return {
            "label": self.label_edit.text().strip(),
            "routeKey": self.route_key_edit.text().strip(),
            "spreadsheetId": self.spreadsheet_id_edit.text().strip(),
            "sheetName": self.sheet_name_edit.text().strip()
        }
