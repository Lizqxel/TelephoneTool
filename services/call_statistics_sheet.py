"""時間帯別の通話集計をスプレッドシートへ送る（送信用Googleフォーム経由）。

会社の設定でGASのウェブアプリを「全員」に公開できないため、ログイン不要の送信用フォームへ送り、
10分ごとの時間トリガー（gas/call_statistics.gs）がまとめてシートの 日付×名前×時間帯 のセルを上書きする。
毎回「本人・当日」の全時間帯を送るので、同じ内容を何度送っても重複せず、失敗しても次回の送信で回復する。
認証トークンは既存の googleFormPosting.tokenValue を使い、各PCで追加の設定を不要にする。
"""

import json

import requests


# gas/call_statistics.gs の setup() が実行ログに出す値。
# settings.json の callStatisticsSheet に同じキーがあればそちらを優先する。
DEFAULT_FORM = {
    "formResponseUrl": "https://docs.google.com/forms/d/e/1FAIpQLSemkvlrLeeoMs86CFFq_igvQaqRAbos24jIpkxffRuln4tMrg/formResponse",
    "tokenEntry": "entry.10799235",
    "dataEntry": "entry.2027761291",
}


def _settings_candidates():
    from services.google_form_sender import GoogleFormSender
    # 設定ファイルの探索順はフォーム送信と同じにする（初期化処理は不要なので呼ばない）。
    return GoogleFormSender._settings_candidates(GoogleFormSender.__new__(GoogleFormSender))


def load_sheet_settings(candidates=None):
    form, token = {}, ""
    for _, data in (candidates if candidates is not None else _settings_candidates()):
        section = data.get("callStatisticsSheet") or {}
        shared = data.get("googleFormPosting") or {}
        if isinstance(section, dict):
            for key in DEFAULT_FORM:
                value = str(section.get(key) or "").strip()
                if value and key not in form:
                    form[key] = value
            token = token or str(section.get("token") or "").strip()
        if isinstance(shared, dict):
            token = token or str(shared.get("tokenValue") or "").strip()
    return {key: form.get(key) or DEFAULT_FORM[key] for key in DEFAULT_FORM}, token


def sheet_hours(rows):
    """シートの「コール・対話・トス」へ送る値。未取得は None（シート側では上書きしない）。"""
    result = []
    for row in rows:
        values = {"hour": row["hour"], "calls": row.get("total_calls"),
                  "dialogues": row.get("dialogues"), "toss": row.get("toss")}
        if any(values[k] is not None for k in ("calls", "dialogues", "toss")):
            result.append(values)
    return result


class SheetSender:
    def __init__(self, form, token, timeout=30, post=None):
        self.form = form
        self.token = token
        self.timeout = timeout
        self.post = post or requests.post

    @classmethod
    def from_settings(cls):
        """送信先が未設定なら None。トークンはリポジトリが公開のためコードに置かず、各PCの設定から読む。"""
        form, token = load_sheet_settings()
        if not all(form.values()):
            return None
        if not token:
            raise RuntimeError("共有トークンが未設定です。設定画面で転記用のトークンを入力すると送信します。")
        return cls(form, token)

    def send(self, day, staff_id, staff_name, hours):
        payload = {"date": day, "staffId": staff_id, "staffName": staff_name, "hours": hours}
        body = {self.form["tokenEntry"]: self.token,
                self.form["dataEntry"]: json.dumps(payload, ensure_ascii=False, separators=(",", ":"))}
        try:
            response = self.post(self.form["formResponseUrl"], data=body, timeout=self.timeout)
        except requests.RequestException as exc:
            raise RuntimeError(f"シートへ接続できませんでした（{exc.__class__.__name__}）") from exc
        if "accounts.google.com" in str(getattr(response, "url", "")):
            raise RuntimeError("送信用フォームがログイン必須になっています。フォームの設定を確認してください。")
        if response.status_code != 200:
            raise RuntimeError(f"送信用フォームの応答が異常です（HTTP {response.status_code}）")
        # フォームは書き込み結果を返さない。シートへの反映結果はGASの実行ログで確認する。
        return {"ok": True}
