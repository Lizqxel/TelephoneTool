"""
スプレッドシート転記用ダイアログ

このダイアログは、Googleフォームへ送信する前に必要な最終入力
（日時・選択肢・自由記述など）をユーザーに確認・編集してもらう
ためのUIを提供します。

仕様:
- トス日(M)と前確コール日(N)は日付ピッカー（既定=当日、Nは空欄も可）
- 架電時間(J)は時刻ピッカー（HH:mm, 既定=現在時刻）
- 商材(H)、新規/見込み(I)、前確コール結果(P)はドロップダウン
- 獲得者名(B)、獲得時管理番号(C)、リスト名(D/E)等は親画面からの初期値反映可

制限事項:
- 獲得時管理番号(C)が未入力の場合はここでの入力を必須とします。
"""

from __future__ import annotations

from typing import Dict, Any
from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit, QPushButton, QCheckBox
from PySide6.QtCore import Qt, QDate, QTime


class SpreadsheetPostDialog(QDialog):
    """スプレッドシート転記前の最終確認ダイアログ"""

    def __init__(self, parent=None, initialValues: Dict[str, Any] | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("スプレッドシート転記")
        self.setModal(True)
        self.resize(480, 420)

        self.values: Dict[str, Any] = initialValues or {}

        layout = QVBoxLayout(self)

        # 転記先（Apps Script ルーティングキー）
        layout.addWidget(QLabel("転記先（シート選択）"))
        self.destCombo = QComboBox()
        for dest in (self.values.get("destinations") or []):
            self.destCombo.addItem(dest.get("label", ""), dest.get("routeKey", ""))
        layout.addWidget(self.destCombo)

        # 管轄（A）
        layout.addWidget(QLabel("管轄"))
        self.kanKatsuInput = QLineEdit(self.values.get("kanKatsu", ""))
        layout.addWidget(self.kanKatsuInput)

        # 獲得者名（B）
        layout.addWidget(QLabel("獲得者名"))
        self.kakutokuShaInput = QLineEdit(self.values.get("kakutokuSha", ""))
        layout.addWidget(self.kakutokuShaInput)

        # 獲得時管理番号（C）
        layout.addWidget(QLabel("獲得時管理番号"))
        self.kakutokuIdInput = QLineEdit(self.values.get("kakutokuId", ""))
        layout.addWidget(self.kakutokuIdInput)

        # リスト名（E）
        layout.addWidget(QLabel("リスト名"))
        self.listNameInput = QLineEdit(self.values.get("listName", ""))
        layout.addWidget(self.listNameInput)

        # 商材（H）
        layout.addWidget(QLabel("商材"))
        self.shozaiCombo = QComboBox()
        self.shozaiCombo.addItems(["NA光", "NP光"])
        self.shozaiCombo.setCurrentText(self.values.get("shozai", "NA光"))
        layout.addWidget(self.shozaiCombo)

        # 新規/見込み（I）
        layout.addWidget(QLabel("新規/見込み"))
        self.kubunCombo = QComboBox()
        self.kubunCombo.addItems(["新規", "見込み"])  # 現状は新規を想定
        self.kubunCombo.setCurrentText(self.values.get("kubun", "新規"))
        layout.addWidget(self.kubunCombo)

        # 架電時間（J）
        layout.addWidget(QLabel("架電時間 (HH:mm)"))
        self.kadenTimeEdit = QTimeEdit()
        self.kadenTimeEdit.setDisplayFormat("HH:mm")
        nowTime = QTime.currentTime()
        kadenTimeStr = self.values.get("kadenTime") or nowTime.toString("HH:mm")
        self.kadenTimeEdit.setTime(QTime.fromString(kadenTimeStr, "HH:mm"))
        layout.addWidget(self.kadenTimeEdit)

        # フリーボックス（K）
        layout.addWidget(QLabel("フリーボックス"))
        self.freeBoxEdit = QTextEdit(self.values.get("freeBox", ""))
        layout.addWidget(self.freeBoxEdit)

        # トス日（M）
        layout.addWidget(QLabel("トス日 (yyyy-MM-dd)"))
        self.tosDateEdit = QDateEdit()
        self.tosDateEdit.setDisplayFormat("yyyy-MM-dd")
        self.tosDateEdit.setCalendarPopup(True)
        today = QDate.currentDate()
        tosDateStr = self.values.get("tosDate") or today.toString("yyyy-MM-dd")
        self.tosDateEdit.setDate(QDate.fromString(tosDateStr, "yyyy-MM-dd"))
        layout.addWidget(self.tosDateEdit)

        # 前確コール日（N）
        layout.addWidget(QLabel("前確コール日 (yyyy-MM-dd)"))
        self.zenkakuDateEdit = QDateEdit()
        self.zenkakuDateEdit.setDisplayFormat("yyyy-MM-dd")
        self.zenkakuDateEdit.setCalendarPopup(True)
        zenkakuDateStr = self.values.get("zenkakuCallDate") or today.toString("yyyy-MM-dd")
        self.zenkakuDateEdit.setDate(QDate.fromString(zenkakuDateStr, "yyyy-MM-dd"))
        layout.addWidget(self.zenkakuDateEdit)

        # 前確コール日 空欄チェック
        self.emptyZenkakuCheck = QCheckBox("前確コール日を空欄にする")
        layout.addWidget(self.emptyZenkakuCheck)

        # 前確コール結果（P）
        layout.addWidget(QLabel("前確コール結果"))
        self.zenkakuResultCombo = QComboBox()
        self.zenkakuResultCombo.addItems(["前確待ち", "トス対象外", "再コール", "前確NG", "前確OK"])  # 既定=前確待ち
        self.zenkakuResultCombo.setCurrentText(self.values.get("zenkakuResult", "前確待ち"))
        layout.addWidget(self.zenkakuResultCombo)

        # ボタン群
        btnLayout = QHBoxLayout()
        self.okBtn = QPushButton("送信")
        self.cancelBtn = QPushButton("キャンセル")
        btnLayout.addWidget(self.okBtn)
        btnLayout.addWidget(self.cancelBtn)
        layout.addLayout(btnLayout)

        self.okBtn.clicked.connect(self.accept)
        self.cancelBtn.clicked.connect(self.reject)

    def getValues(self) -> Dict[str, Any]:
        """ダイアログの入力値を辞書で返す

        Returns:
            Dict[str, Any]: 送信直前の論理キー付きデータ
        """
        zenkakuDate = "" if self.emptyZenkakuCheck.isChecked() else self.zenkakuDateEdit.date().toString("yyyy-MM-dd")
        return {
            "routeKey": self.destCombo.currentData() or "",
            "kanKatsu": self.kanKatsuInput.text().strip(),
            "kakutokuSha": self.kakutokuShaInput.text().strip(),
            "kakutokuId": self.kakutokuIdInput.text().strip(),
            "listName": self.listNameInput.text().strip(),
            "shozai": self.shozaiCombo.currentText(),
            "kubun": self.kubunCombo.currentText(),
            "kadenTime": self.kadenTimeEdit.time().toString("HH:mm"),
            "freeBox": self.freeBoxEdit.toPlainText().strip(),
            "tosDate": self.tosDateEdit.date().toString("yyyy-MM-dd"),
            "zenkakuCallDate": zenkakuDate,
            "zenkakuResult": self.zenkakuResultCombo.currentText(),
        }


