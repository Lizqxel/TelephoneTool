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

from typing import Dict, Any, List, Tuple
import json
import sys
from pathlib import Path
import logging
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

        # settings.json から destinations を読み込み、initialValues の destinations とマージ
        settings_items, settings_default_key = load_destinations_from_settings()
        init_items: List[Dict[str, str]] = list(self.values.get("destinations") or [])

        # routeKey で重複を避けつつマージ（initialValues 優先の上書き）
        merged: Dict[str, Dict[str, str]] = {}
        for it in settings_items:
            rk = (it.get("routeKey") or "").strip()
            if rk:
                merged[rk] = {"label": it.get("label") or rk, "routeKey": rk}
        for it in init_items:
            rk = (it.get("routeKey") or "").strip()
            if rk:
                merged[rk] = {"label": it.get("label") or rk, "routeKey": rk}

        merged_list = list(merged.values())
        if not merged_list and init_items:
            merged_list = init_items

        for dest in merged_list:
            self.destCombo.addItem(dest.get("label", ""), dest.get("routeKey", ""))
        layout.addWidget(self.destCombo)
        # 既定の選択（initialValues に defaultRouteKey があれば優先）
        try:
            default_key = (self.values or {}).get("defaultRouteKey") or settings_default_key
            if default_key:
                for i in range(self.destCombo.count()):
                    if (self.destCombo.itemData(i) or "").strip() == default_key:
                        self.destCombo.setCurrentIndex(i)
                        break
        except Exception:
            pass

        # 管轄（A）: 編集可能なプルダウン + 自由入力可（初期は未入力）
        layout.addWidget(QLabel("管轄"))
        self.kanKatsuCombo = QComboBox()
        self.kanKatsuCombo.setEditable(True)
        # 先頭は空（未入力）を用意
        self.kanKatsuCombo.addItem("")
        # 候補リスト
        for label in ["岩田管轄", "佐藤秀管轄", "杉崎管轄", "角江管轄", "和田管轄"]:
            self.kanKatsuCombo.addItem(label)
        # 初期値（渡されていればそれを表示。なければ空のまま）
        init_k = (self.values.get("kanKatsu") or "").strip()
        if init_k:
            idx = self.kanKatsuCombo.findText(init_k)
            if idx >= 0:
                self.kanKatsuCombo.setCurrentIndex(idx)
            else:
                # 候補外ならそのままテキストとしてセット
                self.kanKatsuCombo.setCurrentText(init_k)
        layout.addWidget(self.kanKatsuCombo)

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
        self.shozaiCombo.addItems(["NA光", "NP光","サポート光","フレッツ光(1G)","フレッツ光(クロス)","USEN光(通常)","転用BIGLOBE光","くらサポ","くらサポ専売","NP光電話N","サポート光電話N","Nアナ戻し"])
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
        # チェック状態に応じて日付入力の有効/無効を切り替え
        def _toggle_zenkaku_enabled(checked: bool):
            self.zenkakuDateEdit.setEnabled(not checked)
        self.emptyZenkakuCheck.toggled.connect(_toggle_zenkaku_enabled)
        # 既定の状態反映
        _toggle_zenkaku_enabled(self.emptyZenkakuCheck.isChecked())

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

        # 入力検証付きの送信
        def _on_accept():
            # 制限事項に基づく必須チェック: 獲得時管理番号(C)
            if not self.kakutokuIdInput.text().strip():
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "入力エラー", "獲得時管理番号(C)が未入力です。入力してください。")
                return
            self.accept()

        self.okBtn.clicked.connect(_on_accept)
        self.cancelBtn.clicked.connect(self.reject)

    def getValues(self) -> Dict[str, Any]:
        """ダイアログの入力値を辞書で返す

        Returns:
            Dict[str, Any]: 送信直前の論理キー付きデータ
        """
        zenkakuDate = "" if self.emptyZenkakuCheck.isChecked() else self.zenkakuDateEdit.date().toString("yyyy-MM-dd")
        return {
            "routeKey": self.destCombo.currentData() or "",
            "routeLabel": self.destCombo.currentText() or "",
            "kanKatsu": self.kanKatsuCombo.currentText().strip(),
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


def load_destinations_from_settings() -> Tuple[List[Dict[str, str]], str]:
    """設定ファイルから Googleフォーム転送先の候補を読み込む。

    探索順: exe直下 → CWD → ソース直下 → _MEIPASS
    ファイル名の優先順: gform_settings.json → settings.json → setteings.json

    Returns:
        (items, default_key)
        items: [{label, routeKey}, ...]
        default_key: 優先既定（設定値 defaultRouteKey があれば優先、無ければ ZAI_HOME→DEV→先頭）
    """
    def _candidate_files() -> List[Path]:
        names = ("gform_settings.json", "settings.json", "setteings.json")
        dirs: List[Path] = []
        try:
            dirs.append(Path(sys.argv[0]).resolve().parent)
        except Exception:
            pass
        try:
            dirs.append(Path.cwd())
        except Exception:
            pass
        try:
            dirs.append(Path(__file__).resolve().parents[1])
        except Exception:
            pass
        try:
            if hasattr(sys, "_MEIPASS"):
                mp = Path(sys._MEIPASS)
                dirs += [mp, mp / "TelephoneTool", mp / "config"]
        except Exception:
            pass
        out: List[Path] = []
        seen: set[str] = set()
        for d in dirs:
            for n in names:
                p = (d / n).resolve()
                ps = str(p)
                if ps in seen:
                    continue
                seen.add(ps)
                if p.exists():
                    out.append(p)
        return out

    # デフォルト
    items: List[Dict[str, str]] = []
    default_key = ""
    for p in _candidate_files():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            gfp = (data or {}).get("googleFormPosting", {})
            dests = list(gfp.get("destinations") or [])
            if not dests:
                continue
            items = []
            for d in dests:
                rk = str((d or {}).get("routeKey", "")).strip()
                if not rk:
                    continue
                label = str((d or {}).get("label", "")).strip() or rk
                items.append({"label": label, "routeKey": rk})
            # defaultRouteKey 優先
            default_key = str(gfp.get("defaultRouteKey", "")).strip()
            if not default_key:
                keys = [i["routeKey"] for i in items]
                for cand in ("ZAI_HOME", "DEV"):
                    if cand in keys:
                        default_key = cand
                        break
                if not default_key and keys:
                    default_key = keys[0]
            logging.info(f"[GForm:UI] destinations loaded from: {p}")
            return items, default_key
        except Exception as e:
            logging.warning(f"[GForm:UI] settings load failed at {p}: {e}")
            continue
    return items, default_key


