"""時間帯別通話集計の計算と永続化（CTIやQtに依存しない）。"""

from dataclasses import dataclass
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import sqlite3
import unicodedata


JST = timezone(timedelta(hours=9))
COUNTERS = ("total_calls", "new_prospects", "ng_with_dialogue", "ng_without_dialogue")


def normalize_name(value):
    return "".join(unicodedata.normalize("NFKC", value or "").split())


def usable_call_id(value):
    """未設定の表示文字列は通話の識別子として使わない。実IDの大小文字は維持する。"""
    text = str(value).strip() if value is not None else ""
    placeholder = unicodedata.normalize("NFKC", text).casefold()
    return "" if placeholder in ("", "-", "0", "null", "(null)") else text


def parse_call_duration(value, numeric_unit="seconds"):
    """数値は選択した単位、H:MM:SS/M:SSは時刻表記として秒に変換する。"""
    if numeric_unit not in ("seconds", "minutes"):
        raise ValueError("通話時間の単位が不正です")
    text = str(value).strip().replace(",", "")
    if text in ("", "-", "―", "—"):
        return 0.0
    try:
        if ":" in text:
            parts = text.split(":")
            if len(parts) not in (2, 3):
                raise ValueError("時間の書式が不正です")
            numbers = [Decimal(p) for p in parts]
            if any(n < 0 or not n.is_finite() for n in numbers) or any(n >= 60 for n in numbers[1:]):
                raise ValueError("時間の範囲が不正です")
            seconds = sum(n * 60 ** i for i, n in enumerate(reversed(numbers)))
        elif re.fullmatch(r"\d+(?:\.\d+)?\s*秒", text):
            seconds = Decimal(text[:-1].strip())
        elif re.fullmatch(r"\d+(?:\.\d+)?\s*分", text):
            seconds = Decimal(text[:-1].strip()) * 60
        else:
            seconds = Decimal(text) * (60 if numeric_unit == "minutes" else 1)
        if not seconds.is_finite() or seconds < 0:
            raise ValueError("通話時間は0以上である必要があります")
        return float(seconds)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"IP電話通話時間の形式を確認してください: {text!r}") from exc


def history_outcome_counters(row):
    """履歴に明記された結果だけを分類する。曖昧な見込・NGは未判定にする。"""
    outcome = normalize_name(row.get("架電結果"))
    detail = normalize_name(row.get("架電結果詳細"))
    if outcome == "NG" and detail == "アプローチNG":
        detail = "対話有NG"
    if outcome in ("留守", "不在", "担当者不在", "受注", "禁止", "削除"):
        return dict(new_prospects=0, ng_with_dialogue=0, ng_without_dialogue=0)
    if outcome in ("新規見込", "継続見込") or (outcome == "見込" and detail in ("新規", "新規見込", "継続", "継続見込")):
        return dict(new_prospects=int(outcome == "新規見込" or detail in ("新規", "新規見込")), ng_with_dialogue=0, ng_without_dialogue=0)
    if outcome in ("対話有NG", "対話無NG") or (outcome == "NG" and detail in ("対話有NG", "対話無NG")):
        result = detail if outcome == "NG" else outcome
        return dict(new_prospects=0, ng_with_dialogue=int(result == "対話有NG"), ng_without_dialogue=int(result == "対話無NG"))
    return dict(new_prospects=None, ng_with_dialogue=None, ng_without_dialogue=None)


def hourly_history_totals(rows, day, staff_name, numeric_unit="seconds"):
    totals = {hour: 0.0 for hour in range(24)}
    counters = {hour: {key: 0 for key in COUNTERS} for hour in range(24)}
    first_hour = None
    seen = {}
    for row in rows:
        if not {"対応日時", "対応スタッフ", "IP電話通話時間"}.issubset(row):
            raise ValueError("集計に必要なカラムを取得できませんでした。")
        if normalize_name(row.get("対応スタッフ")) != normalize_name(staff_name):
            raise ValueError("検索結果に別のスタッフが含まれています。本人だけをチェックしてください。")
        try:
            when = datetime.fromisoformat(row["対応日時"].replace("/", "-"))
        except (ValueError, KeyError) as exc:
            raise ValueError("履歴の対応日時を読み取れませんでした。") from exc
        if when.date().isoformat() != day:
            raise ValueError("検索結果に別の日付が含まれています。架電履歴を当日で検索してください。")
        seconds = parse_call_duration(row.get("IP電話通話時間", ""), numeric_unit)
        unique = usable_call_id(row.get("通話ユニークID"))
        outcome_counts = history_outcome_counters(row)
        identity = (when, seconds, tuple(outcome_counts.items()))
        if unique:
            if unique in seen:
                if seen[unique] != identity:
                    raise ValueError("同じ通話IDの値が異なるため、再取得してください。")
                continue
            seen[unique] = identity
        first_hour = when.hour if first_hour is None else min(first_hour, when.hour)
        totals[when.hour] += seconds
        counters[when.hour]["total_calls"] += 1
        for key, value in outcome_counts.items():
            if value is None or counters[when.hour][key] is None:
                counters[when.hour][key] = None
            else:
                counters[when.hour][key] += value
    return {"talk_seconds": totals, "counters": counters, "first_hour": first_hour}


def hourly_talk_seconds(rows, day, staff_name, numeric_unit="seconds"):
    return hourly_history_totals(rows, day, staff_name, numeric_unit)["talk_seconds"]


def format_duration(seconds):
    if seconds is None:
        return "未取得"
    total = round(seconds)
    return f"{total // 3600:02}:{total // 60 % 60:02}:{total % 60:02}"


def next_half_hour(now):
    return now.replace(minute=0 if now.minute < 30 else 30, second=0, microsecond=0) + timedelta(minutes=30)


@dataclass(frozen=True)
class StatisticsSnapshot:
    acquired_at: datetime
    staff_id: str
    staff_name: str
    session: str
    counters: dict
    talk_seconds: dict
    history_rows: int
    history_counters: dict | None = None
    first_history_hour: int | None = None


class StatisticsStore:
    """日・スタッフごとに保存する。累計差分と履歴再集計を混ぜて加算しない。"""

    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS snapshots (
                    staff_id TEXT NOT NULL, day TEXT NOT NULL, staff_name TEXT NOT NULL,
                    observed_at TEXT NOT NULL, session TEXT NOT NULL,
                    total_calls INTEGER, new_prospects INTEGER,
                    ng_with_dialogue INTEGER, ng_without_dialogue INTEGER,
                    PRIMARY KEY(staff_id,day));
                CREATE TABLE IF NOT EXISTS hourly (
                    staff_id TEXT NOT NULL, day TEXT NOT NULL, hour INTEGER NOT NULL,
                    talk_seconds REAL, total_calls INTEGER, new_prospects INTEGER,
                    ng_with_dialogue INTEGER, ng_without_dialogue INTEGER,
                    counter_status TEXT NOT NULL DEFAULT '未取得',
                    talk_updated_at TEXT, counter_updated_at TEXT,
                    PRIMARY KEY(staff_id,day,hour));
                CREATE TABLE IF NOT EXISTS preferences (key TEXT PRIMARY KEY, value TEXT);
                CREATE TABLE IF NOT EXISTS history_bounds (
                    staff_id TEXT NOT NULL, day TEXT NOT NULL, first_hour INTEGER NOT NULL,
                    PRIMARY KEY(staff_id,day));
                CREATE TABLE IF NOT EXISTS history_hourly (
                    staff_id TEXT NOT NULL, day TEXT NOT NULL, hour INTEGER NOT NULL,
                    total_calls INTEGER, new_prospects INTEGER,
                    ng_with_dialogue INTEGER, ng_without_dialogue INTEGER,
                    PRIMARY KEY(staff_id,day,hour));
            """)

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        try:
            with db:
                yield db
        finally:
            db.close()

    def preference(self, key, default):
        with self.connect() as db:
            row = db.execute("SELECT value FROM preferences WHERE key=?", (key,)).fetchone()
            return row[0] if row else default

    def set_preference(self, key, value):
        with self.connect() as db:
            db.execute("INSERT OR REPLACE INTO preferences VALUES (?,?)", (key, str(value)))

    def save(self, snapshot):
        stamp = snapshot.acquired_at.astimezone(JST)
        day = stamp.date().isoformat()
        with self.connect() as db:
            previous = db.execute("SELECT * FROM snapshots WHERE staff_id=? AND day=?", (snapshot.staff_id, day)).fetchone()
            if previous and stamp <= datetime.fromisoformat(previous["observed_at"]):
                raise ValueError("前回より古い集計結果は保存できません。")
            for hour in range(24):
                db.execute("INSERT OR IGNORE INTO hourly(staff_id,day,hour) VALUES (?,?,?)", (snapshot.staff_id, day, hour))
                # Future hours are unknown, rather than observed zeros.
                if hour <= stamp.hour or (snapshot.history_counters and snapshot.history_counters.get(hour, {}).get("total_calls", 0)):
                    db.execute("UPDATE hourly SET talk_seconds=?,talk_updated_at=? WHERE staff_id=? AND day=? AND hour=?",
                               (snapshot.talk_seconds.get(hour, 0), stamp.isoformat(), snapshot.staff_id, day, hour))
            if previous:
                start = datetime.fromisoformat(previous["observed_at"])
                # Automatic runs near :00/:30 close the preceding interval.
                end = stamp
                if stamp.minute in (0, 30) and stamp.second < 60:
                    boundary = stamp.replace(second=0, microsecond=0)
                    if boundary > start:
                        end = boundary
                last = end - timedelta(microseconds=1)
                same_hour = start.hour == last.hour
                reset = previous["session"] != snapshot.session or any(snapshot.counters[k] < previous[k] for k in COUNTERS)
                if same_hour and not reset:
                    hour = start.hour
                    status = "差分" if start.minute == 0 and start.second < 60 else "一部の差分"
                    row = db.execute("SELECT * FROM hourly WHERE staff_id=? AND day=? AND hour=?", (snapshot.staff_id, day, hour)).fetchone()
                    if row["counter_status"] != "未取得":
                        status = row["counter_status"]
                    sums = [(row[k] or 0) + snapshot.counters[k] - previous[k] for k in COUNTERS]
                    db.execute("UPDATE hourly SET total_calls=?,new_prospects=?,ng_with_dialogue=?,ng_without_dialogue=?,counter_status=?,counter_updated_at=? WHERE staff_id=? AND day=? AND hour=?",
                               (*sums, status, stamp.isoformat(), snapshot.staff_id, day, hour))
                else:
                    status = "リセットあり" if reset else "時間またぎ・差分不明"
                    for hour in range(start.hour, last.hour + 1):
                        db.execute("UPDATE hourly SET counter_status=?,counter_updated_at=? WHERE staff_id=? AND day=? AND hour=?",
                                   (status, stamp.isoformat(), snapshot.staff_id, day, hour))
            db.execute("INSERT OR REPLACE INTO snapshots VALUES (?,?,?,?,?,?,?,?,?)",
                       (snapshot.staff_id, day, snapshot.staff_name, stamp.isoformat(), snapshot.session,
                        *(snapshot.counters[k] for k in COUNTERS)))
            if snapshot.history_counters is not None:
                if snapshot.first_history_hour is not None:
                    db.execute("INSERT INTO history_bounds VALUES (?,?,?) ON CONFLICT(staff_id,day) DO UPDATE SET first_hour=MIN(first_hour,excluded.first_hour)",
                               (snapshot.staff_id, day, snapshot.first_history_hour))
                bound = db.execute("SELECT first_hour FROM history_bounds WHERE staff_id=? AND day=?", (snapshot.staff_id, day)).fetchone()
                if bound:
                    for hour in range(bound[0], 24):
                        counts = snapshot.history_counters.get(hour, {})
                        if hour <= stamp.hour or counts.get("total_calls", 0):
                            db.execute("INSERT OR REPLACE INTO history_hourly VALUES (?,?,?,?,?,?,?)",
                                       (snapshot.staff_id, day, hour, *(counts.get(k) for k in COUNTERS)))

    def staff(self):
        with self.connect() as db:
            return [dict(r) for r in db.execute("SELECT staff_id,staff_name,MAX(observed_at) AS last_seen FROM snapshots GROUP BY staff_id ORDER BY last_seen DESC")]

    def hours(self, staff_id, day):
        with self.connect() as db:
            rows = {r["hour"]: dict(r) for r in db.execute("SELECT * FROM hourly WHERE staff_id=? AND day=? ORDER BY hour", (staff_id, day))}
            for history in db.execute("SELECT * FROM history_hourly WHERE staff_id=? AND day=?", (staff_id, day)):
                if history["hour"] not in rows:
                    continue
                row = rows[history["hour"]]
                missing = []
                for key in COUNTERS:
                    if history[key] is not None:
                        row[key] = history[key]
                    else:
                        missing.append(key)
                row["counter_status"] = "履歴集計" if not missing else "履歴＋差分／未判定"
            return [rows.get(h, {"hour": h, "talk_seconds": None, **{k: None for k in COUNTERS}, "counter_status": "未取得"}) for h in range(24)]

    def display_hours(self, staff_id, day):
        rows = self.hours(staff_id, day)
        with self.connect() as db:
            bound = db.execute("SELECT first_hour FROM history_bounds WHERE staff_id=? AND day=?", (staff_id, day)).fetchone()
        if bound:
            return rows[bound[0]:]
        latest = self.latest(staff_id, day)
        if not latest:
            return []
        observed = [r["hour"] for r in rows if any(r.get(k) is not None for k in COUNTERS)]
        start = min(observed) if observed else datetime.fromisoformat(latest["observed_at"]).hour
        return rows[start:]

    def latest(self, staff_id, day):
        with self.connect() as db:
            row = db.execute("SELECT * FROM snapshots WHERE staff_id=? AND day=?", (staff_id, day)).fetchone()
            return dict(row) if row else None
