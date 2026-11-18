"""全員分の通話集計を、Googleアカウントでログインしてスプレッドシートから読む（読み取り専用）。

一度ブラウザでGoogleにログインすれば、以後はPCに保存した更新用トークンで読み続ける。
シートを読めるアカウント（社内アカウント）なら全員分を表示できる。書き込みはしない。

OAuthクライアント（Google Cloudの「デスクトップアプリ」）の設定は、リポジトリが公開のためコードに置かず、
google_oauth_client.json（gitignore済み・ビルド時にexeへ同梱）か settings.json の callStatisticsSheet.oauthClient から読む。
"""

import base64
from dataclasses import dataclass, field
import json
import logging
import os
from pathlib import Path
import re
import sys
import time
import unicodedata
import webbrowser
import wsgiref.simple_server
import wsgiref.util

SCOPES = ["openid", "https://www.googleapis.com/auth/userinfo.email",
          "https://www.googleapis.com/auth/spreadsheets.readonly"]
SPREADSHEET_ID = "1Asghaq_Hp3tqWQpirbBoBnQpXsKuWiTi33Hzfq-0qRU"
SHEET_NAME_PREFIX = "コール集計打ち込み用_"
CLIENT_FILE = "google_oauth_client.json"
LABELS = ("コール", "対話", "トス")
API = "https://sheets.googleapis.com/v4/spreadsheets"
SUCCESS_PAGE = ("<!doctype html><meta charset='utf-8'><title>TelephoneTool</title>"
                "<p style='font:18px sans-serif;margin:40px'>TelephoneToolのログインが完了しました。このページは閉じてかまいません。</p>")


class LoginCancelled(RuntimeError):
    pass


class _RedirectReceiver:
    """Googleのログイン後に戻ってくる http://localhost:PORT/?code=... を受け取る。"""

    def __init__(self):
        self.uri = None

    def __call__(self, environ, start_response):
        query = environ.get("QUERY_STRING", "")
        # favicon などの別のリクエストは無視し、code か error が付いたものだけを結果にする
        if "code=" in query or "error=" in query:
            self.uri = wsgiref.util.request_uri(environ)
            body = SUCCESS_PAGE if "code=" in query else SUCCESS_PAGE.replace("完了しました", "中止されました")
        else:
            body = ""
        start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
        return [body.encode("utf-8")]


class _QuietHandler(wsgiref.simple_server.WSGIRequestHandler):
    def log_message(self, *args):  # 配布版（コンソール無し）で標準エラーに書かない
        pass


def normalize(value):
    return "".join(unicodedata.normalize("NFKC", str(value or "")).split())


def number(value):
    text = normalize(value).replace(",", "")
    return int(text) if re.fullmatch(r"-?\d+", text) else None


def _search_dirs():
    dirs = []
    for get in (lambda: Path(sys.argv[0]).resolve().parent, Path.cwd, lambda: Path(__file__).resolve().parents[1]):
        try:
            dirs.append(get())
        except Exception:
            pass
    if hasattr(sys, "_MEIPASS"):
        dirs.append(Path(sys._MEIPASS))
    return dirs


def client_config(settings_candidates=None):
    """OAuthクライアント設定（{"installed": {...}} 形式）。見つからなければ None。"""
    for directory in _search_dirs():
        path = directory / CLIENT_FILE
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(data, dict) and ("installed" in data or "web" in data):
                    return data
            except (OSError, ValueError):
                continue
    if settings_candidates is None:
        from services.call_statistics_sheet import _settings_candidates
        settings_candidates = _settings_candidates()
    for _, data in settings_candidates:
        section = (data.get("callStatisticsSheet") or {}).get("oauthClient") if isinstance(data, dict) else None
        if isinstance(section, dict) and section.get("client_id") and section.get("client_secret"):
            return {"installed": {"client_id": section["client_id"], "client_secret": section["client_secret"],
                                  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                                  "token_uri": "https://oauth2.googleapis.com/token",
                                  "redirect_uris": ["http://localhost"]}}
    return None


def default_token_path():
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "TelephoneTool"
    return base / "google_token.json"


def email_from_id_token(id_token):
    """表示用。署名は検証しない（認可には使わない）。"""
    try:
        payload = id_token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload)).get("email", "")
    except Exception:
        return ""


class GoogleLogin:
    def __init__(self, token_path=None, client=None):
        self.token_path = Path(token_path) if token_path else default_token_path()
        self._client = client

    @property
    def client(self):
        if self._client is None:
            self._client = client_config()
        return self._client

    def stored(self):
        try:
            return json.loads(self.token_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    @property
    def email(self):
        return (self.stored() or {}).get("email", "")

    @property
    def logged_in(self):
        return bool((self.stored() or {}).get("refresh_token"))

    def _save(self, credentials, email=None):
        data = json.loads(credentials.to_json())
        data["email"] = email if email is not None else (self.stored() or {}).get("email", "")
        self.token_path.parent.mkdir(parents=True, exist_ok=True)
        self.token_path.write_text(json.dumps(data), encoding="utf-8")

    def login(self, timeout_seconds=300, cancel=None, on_url=None, open_browser=True):
        """既定のブラウザでGoogleのログイン画面を開き、許可されるまで待つ。

        cancel（threading.Event）が立てば中止する。on_url にはログイン画面のURLを渡すので、
        別のブラウザ（社内アカウントのブラウザ）に貼り付けてもログインできる。
        """
        if not self.client:
            raise RuntimeError(f"Googleログインの設定ファイル（{CLIENT_FILE}）がありません。")
        from google_auth_oauthlib.flow import InstalledAppFlow
        os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")
        flow = InstalledAppFlow.from_client_config(self.client, SCOPES)
        receiver = _RedirectReceiver()
        server = wsgiref.simple_server.make_server("localhost", 0, receiver, handler_class=_QuietHandler)
        try:
            flow.redirect_uri = f"http://localhost:{server.server_port}/"
            url, _ = flow.authorization_url(prompt="select_account")
            if on_url:
                on_url(url)
            if open_browser:
                webbrowser.open(url, new=1)
            server.timeout = 0.5  # 0.5秒ごとに中止・時間切れを確かめる
            deadline = time.monotonic() + timeout_seconds
            while receiver.uri is None:
                if cancel is not None and cancel.is_set():
                    raise LoginCancelled("ログインを中止しました。")
                if time.monotonic() > deadline:
                    raise RuntimeError("ログインが時間内に完了しませんでした。もう一度やり直してください。")
                server.handle_request()
        finally:
            server.server_close()
        if "code=" not in receiver.uri:
            raise LoginCancelled("Googleの画面でログインが中止されました。")
        # oauthlib は https 以外を受け付けないため（ライブラリ標準の run_local_server と同じ扱い）
        flow.fetch_token(authorization_response=receiver.uri.replace("http://", "https://", 1))
        credentials = flow.credentials
        if not credentials or not credentials.refresh_token:
            raise RuntimeError("Googleのログインが完了しませんでした。")
        email = email_from_id_token(getattr(credentials, "id_token", "") or "")
        self._save(credentials, email)
        return email

    def logout(self):
        try:
            self.token_path.unlink()
        except FileNotFoundError:
            pass

    def credentials(self):
        data = self.stored()
        if not data or not data.get("refresh_token"):
            raise RuntimeError("Googleにログインしていません。")
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        credentials = Credentials.from_authorized_user_info(data, SCOPES)
        if not credentials.valid:
            try:
                credentials.refresh(Request())
            except Exception as exc:
                raise RuntimeError("Googleのログインの有効期限が切れました。もう一度ログインしてください。") from exc
            self._save(credentials)
        return credentials

    def session(self):
        from google.auth.transport.requests import AuthorizedSession
        return AuthorizedSession(self.credentials())


@dataclass
class Member:
    name: str
    department: str
    unit: str
    total: tuple
    hours: dict = field(default_factory=dict)


@dataclass
class TeamDay:
    sheet: str
    day: str
    hours: list
    members: list
    interval_members: list = field(default_factory=list)
    interval_error: str = ""


@dataclass
class IntervalMember:
    staff_id: str
    name: str
    department: str
    unit: str
    intervals: dict = field(default_factory=dict)  # 開始分 -> (コール, 対話, トス, 対話秒数)
    observed_at: str = ""


INTERVAL_HEADERS = ["日付", "スタッフID", "氏名", "部署", "GPユニット", "開始分", "コール数", "対話数", "トス数", "対話時間（秒）", "取得日時"]


def parse_intervals(rows, day):
    """非表示シートの数値を読む。空欄は0と区別し、重複キーや不正な値は採用しない。"""
    import math
    from datetime import datetime
    from services.call_statistics import JST

    if not rows:
        return []
    if rows[0] != INTERVAL_HEADERS:
        raise RuntimeError("30分集計シートの見出しが一致しません。GASのバージョンを確認してください。")
    people = {}
    for row in rows[1:]:
        if _cell(row, 0) != day:
            continue
        staff_id = str(_cell(row, 1))
        name = str(_cell(row, 2))
        minute = number(_cell(row, 5))
        stamp = str(_cell(row, 10))
        try:
            parsed = datetime.fromisoformat(stamp)
            if parsed.tzinfo is None or parsed.astimezone(JST).date().isoformat() != day:
                raise ValueError()
        except ValueError as exc:
            raise RuntimeError("30分集計シートの取得日時が不正です。") from exc
        if not staff_id or not name or minute not in range(0, 1440, 30):
            raise RuntimeError("30分集計シートのスタッフまたは時間帯が不正です。")
        acquired = parsed.astimezone(JST)
        if minute > acquired.hour * 60 + acquired.minute:
            raise RuntimeError("30分集計シートに取得日時より後の時間帯があります。")
        values = []
        for index in range(6, 10):
            cell = _cell(row, index)
            if cell in ("", None):
                values.append(None)
                continue
            try:
                value = float(cell)
                if isinstance(cell, bool) or not math.isfinite(value) or value < 0 or (index < 9 and not value.is_integer()):
                    raise ValueError()
            except (TypeError, ValueError) as exc:
                raise RuntimeError("30分集計シートの件数または時間が不正です。") from exc
            values.append(int(value) if index < 9 else value)
        if ((values[0] is not None and values[1] is not None and values[1] > values[0])
                or (values[1] is None and values[3] is not None)):
            raise RuntimeError("30分集計シートの対話数・対話時間が矛盾しています。")
        member = people.setdefault(staff_id, IntervalMember(staff_id, name, str(_cell(row, 3)), str(_cell(row, 4))))
        if minute in member.intervals:
            raise RuntimeError("30分集計シートに同じスタッフ・時間帯が重複しています。")
        member.intervals[minute] = tuple(values)
        if not member.observed_at or parsed > datetime.fromisoformat(member.observed_at):
            member.observed_at = stamp
    return list(people.values())


def _cell(row, index):
    return row[index] if index < len(row) else ""


def find_block(left, month, day):
    """「【11/1（日）】」の見出し行（0始まり）と、次の日付の見出し行。"""
    marker = normalize(f"【{month}/{day}（")
    block = next((i for i, row in enumerate(left) if any(normalize(c).startswith(marker) for c in row)), None)
    if block is None:
        raise RuntimeError(f"{month}/{day}の表がシートにありません。")
    end = next((i for i in range(block + 1, len(left))
                if any(re.match(r"^【\d+/\d+\(", normalize(c)) for c in left[i])), len(left))
    return block, end


def parse_block(rows):
    """見出し2行＋名前の行から、全員の 日計・時間帯別の コール/対話/トス を取り出す。"""
    if len(rows) < 2:
        raise RuntimeError("表の見出しを読み取れませんでした。")
    hour_row, label_row = rows[0], rows[1]
    width = max(len(hour_row), len(label_row))
    name_column = next((i for i in range(min(4, len(hour_row))) if normalize(hour_row[i]) == "名前"), None)
    if name_column is None:
        raise RuntimeError("表に「名前」の列がありません。")
    groups = {}
    for index in range(width):
        title = normalize(_cell(hour_row, index))
        labels = tuple(normalize(_cell(label_row, index + k)) for k in range(3))
        if labels != LABELS:
            continue
        match = re.fullmatch(r"(\d+)時台", title)
        if match:
            groups[int(match.group(1))] = index
        elif title == "日計":
            groups["total"] = index
    if "total" not in groups:
        raise RuntimeError("表に「日計」の列がありません。")
    hours = sorted(h for h in groups if h != "total")
    members = []
    for row in rows[2:]:
        name = str(_cell(row, name_column)).strip()
        if not name:
            continue
        values = lambda start: tuple(number(_cell(row, start + k)) for k in range(3))
        members.append(Member(name=name, department=str(_cell(row, name_column + 1)).strip(),
                              unit=str(_cell(row, name_column + 2)).strip(), total=values(groups["total"]),
                              hours={h: values(groups[h]) for h in hours}))
    return hours, members


def _sheet_read_error(response, sheet, range_name):
    """Googleのエラー項目だけを記録する（ヘッダーや認証情報は記録しない）。"""
    try:
        body = response.json()
    except ValueError:
        body = {}
    error = body.get("error", {}) if isinstance(body, dict) else {}
    error = error if isinstance(error, dict) else {}
    status = str(error.get("status") or "")
    message = str(error.get("message") or "")
    reasons = set()
    for key in ("errors", "details"):
        items = error.get(key)
        if isinstance(items, list):
            reasons.update(str(item["reason"]) for item in items if isinstance(item, dict) and item.get("reason"))
    logging.getLogger(__name__).warning(
        "全員の通話集計: シート読み取り失敗 HTTP=%s sheet=%s range=%s status=%s reasons=%s message=%s",
        response.status_code, sheet, range_name, status, ",".join(sorted(reasons)), message,
    )
    if reasons & {"SERVICE_DISABLED", "accessNotConfigured", "apiNotEnabled"}:
        return RuntimeError("Googleログインに使用しているプロジェクトでGoogle Sheets APIが無効です。管理者にAPIの有効化を依頼してください。")
    if reasons & {"ACCESS_TOKEN_SCOPE_INSUFFICIENT", "insufficientPermissions"}:
        return RuntimeError("Googleログインでシート読み取りの許可が不足しています。ログアウトし、再ログイン時にシートの読み取りを許可してください。")
    if response.status_code == 401:
        return RuntimeError("Googleの認証が無効です。もう一度ログインしてください。")
    if response.status_code == 403:
        return RuntimeError("シートを読めません（HTTP 403）。閲覧権限またはGoogle側の利用制限を確認してください。詳細はapp.logに記録しました。")
    if response.status_code == 400 and "Unable to parse range" in message:
        return RuntimeError(f"シート「{sheet}」がありません。")
    return RuntimeError(f"シートを読み取れませんでした（HTTP {response.status_code}）。詳細はapp.logに記録しました。")


def fetch_day(session, day):
    """day は 'YYYY-MM-DD'。シートの該当日の表を2回の読み取りで取得する。"""
    year, month, date = (int(x) for x in day.split("-"))
    sheet = f"{SHEET_NAME_PREFIX}{month}月"

    def read(range_name):
        try:
            response = session.get(f"{API}/{SPREADSHEET_ID}/values:batchGet",
                                   params={"ranges": f"'{sheet}'!{range_name}", "valueRenderOption": "FORMATTED_VALUE"},
                                   timeout=60)
        except Exception as exc:
            logging.getLogger(__name__).warning(
                "全員の通話集計: シート接続失敗 sheet=%s range=%s error_type=%s",
                sheet, range_name, exc.__class__.__name__,
            )
            raise RuntimeError(f"シートへ接続できませんでした（{exc.__class__.__name__}）") from exc
        if response.status_code != 200:
            raise _sheet_read_error(response, sheet, range_name)
        ranges = response.json().get("valueRanges") or [{}]
        return ranges[0].get("values", [])

    left = read("A1:D")
    if not left or f"{year}年{month}月" not in normalize("".join(left[0])):
        raise RuntimeError(f"シート「{sheet}」が{year}年{month}月の表ではありません。")
    block, end = find_block(left, month, date)
    # block・end は0始まりの行番号。見出し（時台）の行から、次の日付の見出しの直前までを行ごと読む。
    hours, members = parse_block(read(f"{block + 2}:{max(end, block + 3)}"))
    data = TeamDay(sheet=sheet, day=day, hours=hours, members=members)
    hidden = f"通話集計データ_{year:04}-{month:02}"
    try:
        response = session.get(f"{API}/{SPREADSHEET_ID}/values:batchGet",
                               params={"ranges": f"'{hidden}'!A1:K", "valueRenderOption": "UNFORMATTED_VALUE"}, timeout=60)
        if response.status_code != 200:
            raise _sheet_read_error(response, hidden, "A1:K")
        values = (response.json().get("valueRanges") or [{}])[0].get("values", [])
        data.interval_members = parse_intervals(values, day)
        if not data.interval_members:
            data.interval_error = "この日の30分集計はまだありません。各PCから集計を送信してください。"
    except Exception as exc:
        logging.getLogger(__name__).warning("全員の通話集計: 30分データ取得失敗 error_type=%s", type(exc).__name__)
        data.interval_error = f"30分集計を取得できませんでした：{exc}"
    # 未送信の人も対象人数に含め、取得済みの人だけで全員分と誤認させない。
    sent_names = {normalize(member.name) for member in data.interval_members}
    for member in members:
        if normalize(member.name) not in sent_names:
            data.interval_members.append(IntervalMember("", member.name, member.department, member.unit))
    return data
