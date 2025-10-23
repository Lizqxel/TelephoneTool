"""
初回起動時セットアップ

・settings.json が存在しない場合に作成を担当
・共有トークンをマスク入力で必須取得（未入力なら転記機能部分は生成しない＝転記不可）
・仕様は既存の動作に影響を与えない（キー名や既定値は既存に準拠）

このモジュールは UI あり/なし双方から呼び出せるよう、PySide6 が利用可能なら
QInputDialog + Password エコーモードで入力ダイアログを表示し、
GUIがない場合はコンソールで getpass 風に入力促す。
"""
from __future__ import annotations

import json
import os
import sys
from typing import Optional, Dict, Any


def _app_root_dir() -> str:
    """実行形態（exe/スクリプト）に応じた設定配置ディレクトリを返す"""
    try:
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _masked_input(prompt: str) -> Optional[str]:
    """GUI優先でパスワード入力を促す。GUI不可ならコンソールで入力。

    Returns: 入力値（空文字は空とみなす）。キャンセルは None。
    """
    # GUI（PySide6）での入力
    try:
        from PySide6.QtWidgets import QApplication, QInputDialog, QLineEdit
        app = QApplication.instance() or QApplication([])
        text, ok = QInputDialog.getText(None, "共有トークンの入力", 
                                        "初回セットアップです。共有トークンを入力してください:\n"
                                        "（入力はマスク表示されます）",
                                        QLineEdit.EchoMode.Password)
        if ok:
            return str(text)
        return None
    except Exception:
        pass

    # コンソールでの入力
    try:
        import getpass
        return getpass.getpass("共有トークン（入力は表示されません）: ")
    except Exception:
        # どうしても入力できない環境では None を返す
        return None


def _default_settings() -> Dict[str, Any]:
    """転記機能以外のデフォルト設定を返す（既存仕様を変更しない）。"""
    default_format = (
        "対応者（お客様の名前）：{operator}\n"
        "工事希望日\n"
        "★出やすい時間帯：{available_time} \n"
        "★電話取次：アナログ→光電話\n"
        "★電話OP：\n"
        "★無線\n"
        "契約者(書類名義)：{contractor}\n"
        "フリガナ：{furigana}\n"
        "生年月日：{birth_date}\n"
        "郵便番号：{postal_code}\n"
        "住所：{address}\n"
        "リスト名：{list_name}\n"
        "リスト名フリガナ：{list_furigana}\n"
        "電話番号：{list_phone}\n"
        "リスト郵便番号：{list_postal_code}\n"
        "リスト住所：{list_address}\n"
        "現状回線：{current_line}\n"
        "受注日：{order_date}\n"
        "受注者：{order_person}\n"
        "提供判定：{judgment}\n\n"
        "料金認識：{fee}\n"
        "ネット利用：{net_usage}\n"
        "家族了承：{family_approval}\n\n"
        "他番号：{other_number}\n"
        "電話機：{phone_device}\n"
        "禁止回線：{forbidden_line}\n"
        "ND：{nd}\n\n"
        "備考：{relationship}\n"
        "お客様が今使っている回線：アナログ\n"
        "案内料金：3650円\n"
    )
    return {
        "format_template": default_format,
        "font_size": 9,
        "delay_seconds": 0,
        "browser_settings": {
            "headless": True,
            "disable_images": True,
            "show_popup": True,
            "auto_close": True,
            "page_load_timeout": 30,
            "script_timeout": 30
        },
        "mode": "simple",
        "show_mode_selection": False,
        "enable_cti_monitoring": True,
        "enable_auto_cti_processing": True,
        "cti_monitor_interval": 0.5,
        "cti_auto_processing_cooldown": 3.0,
        "call_duration_threshold": 0
    }


def ensure_settings_file() -> str:
    """settings.json を生成（なければ）。共有トークンは必須入力。

    - 入力キャンセル/未入力の場合は、googleFormPosting ブロックを生成しない。
      この場合は転記機能だけが利用不可になる。
    - 入力がある場合は tokenValue と entryMap を含む最小構成を生成。

    Returns: 設定ファイルの絶対パス
    """
    root = _app_root_dir()
    settings_path = os.path.join(root, 'settings.json')
    if os.path.exists(settings_path):
        return settings_path

    # 共有トークンの入力
    token = _masked_input("共有トークン")
    # デフォルト設定を構築
    data: Dict[str, Any] = _default_settings()

    # トークンが入力された場合のみ転記関連を生成
    if token:
        data["googleFormPosting"] = {
            "formUrl": "https://docs.google.com/forms/d/e/1FAIpQLSfoDjvD0gyYxqcmADbwBtJ9CSsVB7QAHX8gSRG9sLXzsSNXYQ/formResponse",
            # フォームURLや destinations は後から設定可能。まずはトークンと entryMap を配置
            "tokenValue": token,
            "timezone": "Asia/Tokyo",
            "retryPolicy": {"maxAttempts": 3, "backoffSeconds": [1, 3, 10]},
            "defaults": {
                "kanKatsu": "岩田管轄",
                "shozai": "NA光",
                "kubun": "新規",
                "zenkakuResult": "前確入力待ち"
            },
            "choices": {},
            "entryMap": {
                "kanKatsu": "entry.1653774096",
                "kakutokuSha": "entry.990044023",
                "kakutokuId": "entry.20053472",
                "listName": "entry.311805903",
                "shozai": "entry.444975123",
                "kubun": "entry.858726565",
                "kadenTime": "entry.1423382151",
                "freeBox": "entry.1331784643",
                "tosDate": "entry.1643496135",
                "zenkakuCallDate": "entry.2129109581",
                "zenkakuResult": "entry.1899186247",
                "sharedToken": "entry.447700198",
                "spreadsheetUrl": "entry.574347119",
                "sheetName": "entry.556971462"
            }
        }
    # ファイル保存
    try:
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        # 書けない環境の場合は実行ディレクトリに書く
        fallback = os.path.join(os.getcwd(), 'settings.json')
        with open(fallback, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return fallback

    return settings_path


__all__ = ["ensure_settings_file"]
