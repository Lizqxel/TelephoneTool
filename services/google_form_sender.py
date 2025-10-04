"""
Googleフォーム送信サービス

このモジュールは、アプリ内で入力された営業コメントの要約情報を
Googleフォーム（formResponse）へPOSTし、スプレッドシートに転記させる
機能を提供します。

概要:
- 設定ファイル（settings.json）の`googleFormPosting`セクションを参照し、
  送信先URL・entryマッピング・既定値・リトライ方針を適用します。
- 値のバリデーション（必須・形式）を行い、エラー時は詳細情報を含む
  例外を送出します。
- ネットワーク障害等に備え、指数バックオフのリトライを行います。

制限事項:
- Googleフォームの仕様変更（entry.<id>やバリデーション）に依存します。
- HTTP 200でもフォーム側エラーがあり得るため、真正な成功は
  スプレッドシート側のApps Scriptによって最終的に担保されます。

推奨運用:
- Apps Script側でユニークキー（例: C+M+J）重複検出・整形転記を実施。
- 共有トークンをフォーム必須項目にして防御を行う。
"""

from __future__ import annotations

import re
import time
import logging
from typing import Dict, Any, Optional, Tuple

import requests

from utils.settings import settings


class GoogleFormSender:
    """Googleフォームへの送信を担当するサービスクラス"""

    def __init__(self) -> None:
        """初期化

        設定ファイルから送信設定を読み込みます。
        """
        # 設定の探索順: gform_settings.json → settings.json → setteings.json → utils.settings
        self.config, loaded_from = self._load_google_form_config()
        if loaded_from:
            logging.info(f"[GForm] config loaded from: {loaded_from}")
        else:
            logging.info("[GForm] config loaded from: utils.settings")

        # formUrl / url の両対応
        self.formUrl: str = self.config.get("formUrl") or self.config.get("url")
        if not self.formUrl:
            raise ValueError("Googleフォーム送信設定に必須キー 'formUrl' が存在しません")

        self.tokenValue: str = self._get_required("tokenValue")
        self.timezone: str = self.config.get("timezone", "Asia/Tokyo")
        self.retryPolicy: Dict[str, Any] = self.config.get("retryPolicy", {"maxAttempts": 3, "backoffSeconds": [1, 3, 10]})
        self.defaults: Dict[str, str] = self.config.get("defaults", {})
        self.choices: Dict[str, Any] = self.config.get("choices", {})
        self.entryMap: Dict[str, str] = self._get_required("entryMap")

    def _get_required(self, key: str) -> Any:
        """必須設定の取得

        Args:
            key (str): 設定キー

        Raises:
            ValueError: 設定が存在しない場合

        Returns:
            Any: 設定値
        """
        if key not in self.config:
            raise ValueError(f"Googleフォーム送信設定に必須キー '{key}' が存在しません")
        return self.config[key]

    def send(self, payload: Dict[str, Any]) -> None:
        """Googleフォームへ送信

        Args:
            payload (Dict[str, Any]): アプリ内の論理キーで構成されたデータ
                期待キー:
                    - kanKatsu (管轄) 例: 岩田管轄
                    - kakutokuSha (獲得者名)
                    - kakutokuId (獲得時管理番号) 例: 0171_241009_00039508
                    - listName (リスト名)
                    - shozai (商材) 例: NA光/NP光
                    - kubun (新規/見込み) 例: 新規
                    - kadenTime (架電時間) 例: HH:mm
                    - freeBox (フリーボックス)
                    - tosDate (トス日) 例: yyyy-MM-dd
                    - zenkakuCallDate (前確コール日) 例: yyyy-MM-dd or ''
                    - zenkakuResult (前確コール結果) 例: 前確待ち

        Raises:
            ValueError: バリデーションエラーの場合（詳細メッセージ含む）
            RuntimeError: HTTPエラーや未期待レスポンスの場合
        """
        # 既定値の補完
        data: Dict[str, Any] = dict(payload)
        data.setdefault("kanKatsu", self.defaults.get("kanKatsu", "岩田管轄"))
        data.setdefault("shozai", self.defaults.get("shozai", "NA光"))
        data.setdefault("kubun", self.defaults.get("kubun", "新規"))
        data.setdefault("zenkakuResult", self.defaults.get("zenkakuResult", "前確待ち"))

        # バリデーション
        self._validate(data)

        # エントリマッピングへ変換
        formBody: Dict[str, str] = {}
        for logicalKey, value in data.items():
            entryKey = self.entryMap.get(logicalKey)
            if entryKey:
                # Noneは送らない（空文字に）
                formBody[entryKey] = "" if value is None else str(value)
        # 共有トークン
        formBody[self.entryMap["sharedToken"]] = self.tokenValue
        # ルーティングキー（Apps Script 側で複数スプレッドシート振り分け）
        route_key_entry = self.entryMap.get("routeKey")
        if route_key_entry and data.get("routeKey"):
            formBody[route_key_entry] = str(data["routeKey"])  # 例: DEV / ZAI_HOME

        # フォーム短文化後: choicesが空なら“その他”は適用しない
        if self.choices:
            self._apply_other_option(formBody, data)

        # 送信（リトライ付き）
        maxAttempts = int(self.retryPolicy.get("maxAttempts", 3))
        backoff = list(self.retryPolicy.get("backoffSeconds", [1, 3, 10]))
        attempt = 0
        lastErr: Optional[Exception] = None
        while attempt < maxAttempts:
            attempt += 1
            try:
                resp = requests.post(self.formUrl, data=formBody, headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=20)
                status = resp.status_code
                if status != 200:
                    raise RuntimeError(f"Googleフォーム送信失敗: HTTP {status} (attempt={attempt})")
                logging.info(f"Googleフォーム送信成功 (attempt={attempt}) status={status}")
                return
            except Exception as e:  # ネットワーク系含む
                lastErr = e
                logging.warning(f"Googleフォーム送信エラー (attempt={attempt}): {e}")
                if attempt >= maxAttempts:
                    break
                sleepSec = backoff[min(attempt - 1, len(backoff) - 1)]
                time.sleep(sleepSec)

        # リトライ枯渇
        raise RuntimeError(f"Googleフォーム送信エラー: {lastErr}")

    # ===== 内部ヘルパ =====
    def _settings_candidates(self) -> list[Tuple[str, Dict[str, Any]]]:
        """gform_settings.json / settings.json / setteings.json を実行形態に応じて探索し、
        見つかった全候補の (path, data) をリストで返す。

        先に見つかった空の設定に引っ張られないよう、呼び出し側で要件を満たすものを選別する。
        """
        results: list[Tuple[str, Dict[str, Any]]] = []
        try:
            import sys, json
            from pathlib import Path

            names = ("gform_settings.json", "settings.json", "setteings.json")
            dirs = []
            # exe と同じフォルダ
            try:
                dirs.append(Path(sys.argv[0]).resolve().parent)
            except Exception:
                pass
            # カレント
            try:
                dirs.append(Path.cwd())
            except Exception:
                pass
            # 開発時のソース直下（.../TelephoneTool）
            try:
                dirs.append(Path(__file__).resolve().parents[1])
            except Exception:
                pass
            # PyInstaller onefile 展開先
            try:
                if hasattr(sys, "_MEIPASS"):
                    mp = Path(sys._MEIPASS)
                    dirs += [mp, mp / "TelephoneTool", mp / "config"]
            except Exception:
                pass

            seen: set[str] = set()
            for d in dirs:
                for n in names:
                    try:
                        p = (d / n).resolve()
                    except Exception:
                        continue
                    ps = str(p)
                    if ps in seen:
                        continue
                    seen.add(ps)
                    if p.exists():
                        try:
                            txt = p.read_text(encoding="utf-8")
                            data = json.loads(txt)
                            if isinstance(data, dict):
                                results.append((ps, data))
                        except Exception:
                            # 壊れた JSON はスキップ
                            continue
        except Exception:
            # 何らかの理由で探索自体に失敗した場合は空を返す
            return []
        return results

    def _load_google_form_config(self) -> Tuple[Dict[str, Any], Optional[str]]:
        """googleFormPosting 設定を外部ファイル優先で読み込む。

        優先順: gform_settings.json → settings.json → setteings.json → utils.settings。
        さらに、googleFormPosting が存在しても必須キー（url/formUrl と entryMap）を満たさないものはスキップし、
        次の候補を評価する。
        """
        candidates = self._settings_candidates()
        for path, data in candidates:
            try:
                g = data.get("googleFormPosting") or {}
                if not isinstance(g, dict) or not g:
                    continue
                # 必須キーの存在確認（url/formUrl と entryMap のいずれも必須）
                if (g.get("formUrl") or g.get("url")) and g.get("entryMap"):
                    return g, path
            except Exception:
                continue

        # フォールバック: utils.settings（こちらも最低限の整合性チェックを行う）
        g = settings.get("googleFormPosting", {}) or {}
        return g, None

    def _validate(self, d: Dict[str, Any]) -> None:
        """入力値の検証

        Args:
            d (Dict[str, Any]): 入力データ

        Raises:
            ValueError: いずれかの検証に失敗した場合
        """
        errors = []

        def required(name: str) -> None:
            if not d.get(name):
                errors.append(f"必須項目が未入力: {name}")

        # 必須
        for key in ["kanKatsu", "kakutokuSha", "kakutokuId", "shozai", "kubun", "kadenTime", "tosDate", "zenkakuResult"]:
            required(key)

        # 形式チェック
        if d.get("kakutokuId") and not re.match(r"^\d{4}_\d{6}_\d{8}$", str(d["kakutokuId"])):
            errors.append(f"獲得時管理番号の形式不正: {d.get('kakutokuId')}")

        if d.get("kadenTime") and not re.match(r"^\d{2}:\d{2}$", str(d["kadenTime"])):
            errors.append(f"架電時間の形式はHH:mmで指定してください: {d.get('kadenTime')}")

        def is_date_fmt(val: str) -> bool:
            return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", val))

        if d.get("tosDate") and not is_date_fmt(str(d["tosDate"])):
            errors.append(f"トス日の形式はyyyy-MM-ddで指定してください: {d.get('tosDate')}")

        if d.get("zenkakuCallDate") and d["zenkakuCallDate"] != "" and not is_date_fmt(str(d["zenkakuCallDate"])):
            errors.append(f"前確コール日の形式はyyyy-MM-ddで指定してください: {d.get('zenkakuCallDate')}")

        # 選択肢チェック
        if d.get("shozai") and str(d["shozai"]) not in ["NA光", "NP光"]:
            errors.append(f"商材は 'NA光' または 'NP光' を指定してください: {d.get('shozai')}")

        # 前確のバリデーション（GASで最終正規化するため、ここでは候補緩め）
        if d.get("zenkakuResult") and str(d["zenkakuResult"]) not in ["前確待ち", "トス対象外", "再コール", "前確NG", "前確OK", "前確入力待ち"]:
            errors.append(f"前確コール結果が不正です: {d.get('zenkakuResult')}")

        if errors:
            raise ValueError("; ".join(errors))

    def _apply_other_option(self, formBody: Dict[str, str], d: Dict[str, Any]) -> None:
        """フォームの“その他”送信を適用

        - Googleフォームの仕様では、`entry.<id>=__other_option__` と
          `entry.<id>.other_option_response=自由入力` の2つを送る
        - 対象: 管轄(kanKatsu), 獲得者名(kakutokuSha)
        """
        # 管轄
        kan = str(d.get("kanKatsu", ""))
        if kan and self.choices.get("kanKatsu") and kan not in self.choices["kanKatsu"]:
            base = self.entryMap.get("kanKatsu")
            other = self.entryMap.get("kanKatsuOther")
            if base and other:
                formBody[base] = "__other_option__"
                formBody[other] = kan

        # 獲得者名
        name = str(d.get("kakutokuSha", ""))
        if name and self.choices.get("kakutokuSha") and name not in self.choices["kakutokuSha"]:
            base = self.entryMap.get("kakutokuSha")
            other = self.entryMap.get("kakutokuShaOther")
            if base and other:
                formBody[base] = "__other_option__"
                formBody[other] = name


