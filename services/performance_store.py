"""Local, recoverable storage for CTI performance statistics.

The database is intentionally independent from the existing Google Form flow.
It stores no customer name, phone number, address, or management number.
"""

from __future__ import annotations

import os
import sqlite3
import sys
import threading
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


def default_performance_db_path() -> str:
    """Return a writable per-user database path."""
    candidates = []
    for env_name in ("LOCALAPPDATA", "APPDATA"):
        base = os.environ.get(env_name)
        if base:
            candidates.append(Path(base) / "TelephoneTool")
    candidates.append(Path.home() / ".telephonetool")
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / "data")
    else:
        candidates.append(Path(__file__).resolve().parents[1] / "data")
    candidates.append(Path.cwd() / "data")
    for root in candidates:
        try:
            root.mkdir(parents=True, exist_ok=True)
            probe = root / ".performance-write-test"
            probe.touch(exist_ok=True)
            probe.unlink()
            return str(root / "performance.db")
        except OSError:
            continue
    raise OSError("実績データを保存できるフォルダが見つかりません")


class PerformanceStore:
    """Thread-safe SQLite store used by the collector and sync worker."""

    def __init__(self, db_path: Optional[str] = None) -> None:
        self.db_path = db_path or default_performance_db_path()
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA busy_timeout=15000")
        return conn

    def _initialize(self) -> None:
        with self._lock, self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS hourly_stats (
                    business_date TEXT NOT NULL,
                    hour INTEGER NOT NULL,
                    staff_id TEXT NOT NULL,
                    staff_name TEXT NOT NULL,
                    calls INTEGER NOT NULL DEFAULT 0,
                    dialogues INTEGER NOT NULL DEFAULT 0,
                    tosses INTEGER NOT NULL DEFAULT 0,
                    talk_seconds INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (business_date, hour, staff_id, source)
                );

                CREATE TABLE IF NOT EXISTS call_details (
                    event_id TEXT PRIMARY KEY,
                    business_date TEXT NOT NULL,
                    hour INTEGER NOT NULL,
                    staff_id TEXT NOT NULL,
                    staff_name TEXT NOT NULL,
                    called_at TEXT NOT NULL,
                    result TEXT NOT NULL DEFAULT '',
                    is_dialogue INTEGER NOT NULL DEFAULT 0,
                    is_toss INTEGER NOT NULL DEFAULT 0,
                    talk_seconds INTEGER NOT NULL DEFAULT 0,
                    source TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS daily_totals (
                    business_date TEXT NOT NULL,
                    staff_id TEXT NOT NULL,
                    staff_name TEXT NOT NULL,
                    calls INTEGER NOT NULL,
                    dialogues INTEGER NOT NULL,
                    tosses INTEGER NOT NULL,
                    captured_hour INTEGER NOT NULL,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY (business_date, staff_id)
                );

                CREATE INDEX IF NOT EXISTS idx_hourly_date
                    ON hourly_stats (business_date, staff_id, hour);
                CREATE INDEX IF NOT EXISTS idx_detail_date
                    ON call_details (business_date, staff_id, called_at);
                """
            )

    @staticmethod
    def _now_iso() -> str:
        return datetime.now(timezone(timedelta(hours=9))).isoformat(timespec="seconds")

    def get_or_create_client_id(self) -> str:
        existing = self.get_metadata("client_id")
        if existing:
            return existing
        value = str(uuid.uuid4())
        self.set_metadata("client_id", value)
        return value

    def get_metadata(self, key: str, default: str = "") -> str:
        with self._lock, self._connect() as conn:
            row = conn.execute("SELECT value FROM metadata WHERE key = ?", (key,)).fetchone()
            return str(row["value"]) if row else default

    def set_metadata(self, key: str, value: str) -> None:
        with self._lock, self._connect() as conn:
            conn.execute(
                "INSERT INTO metadata(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, str(value)),
            )

    def add_delta(
        self,
        business_date: str,
        hour: int,
        staff_id: str,
        staff_name: str,
        calls: int = 0,
        dialogues: int = 0,
        tosses: int = 0,
        talk_seconds: int = 0,
        source: str = "live",
    ) -> None:
        now = self._now_iso()
        values = (
            business_date,
            int(hour),
            staff_id,
            staff_name,
            max(0, int(calls)),
            max(0, int(dialogues)),
            max(0, int(tosses)),
            max(0, int(talk_seconds)),
            source,
            now,
        )
        with self._lock, self._connect() as conn:
            conn.execute(
                """
                INSERT INTO hourly_stats(
                    business_date, hour, staff_id, staff_name, calls, dialogues,
                    tosses, talk_seconds, source, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(business_date, hour, staff_id, source) DO UPDATE SET
                    staff_name = excluded.staff_name,
                    calls = hourly_stats.calls + excluded.calls,
                    dialogues = hourly_stats.dialogues + excluded.dialogues,
                    tosses = hourly_stats.tosses + excluded.tosses,
                    talk_seconds = hourly_stats.talk_seconds + excluded.talk_seconds,
                    updated_at = excluded.updated_at
                """,
                values,
            )

    def add_call_detail(self, detail: Dict[str, Any]) -> None:
        now = self._now_iso()
        with self._lock, self._connect() as conn:
            conn.execute(
                """
                INSERT INTO call_details(
                    event_id, business_date, hour, staff_id, staff_name, called_at,
                    result, is_dialogue, is_toss, talk_seconds, source, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(event_id) DO UPDATE SET
                    staff_name = excluded.staff_name,
                    result = excluded.result,
                    is_dialogue = excluded.is_dialogue,
                    is_toss = excluded.is_toss,
                    talk_seconds = excluded.talk_seconds,
                    updated_at = excluded.updated_at
                """,
                (
                    detail["event_id"],
                    detail["business_date"],
                    int(detail["hour"]),
                    detail["staff_id"],
                    detail["staff_name"],
                    detail["called_at"],
                    detail.get("result", ""),
                    1 if detail.get("is_dialogue") else 0,
                    1 if detail.get("is_toss") else 0,
                    max(0, int(detail.get("talk_seconds", 0))),
                    detail.get("source", "live"),
                    now,
                ),
            )

    def set_daily_totals(
        self,
        business_date: str,
        staff_id: str,
        staff_name: str,
        calls: int,
        dialogues: int,
        tosses: int,
        captured_hour: int,
    ) -> None:
        with self._lock, self._connect() as conn:
            conn.execute(
                """
                INSERT INTO daily_totals(
                    business_date, staff_id, staff_name, calls, dialogues,
                    tosses, captured_hour, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(business_date, staff_id) DO UPDATE SET
                    staff_name = excluded.staff_name,
                    calls = excluded.calls,
                    dialogues = excluded.dialogues,
                    tosses = excluded.tosses,
                    captured_hour = excluded.captured_hour,
                    updated_at = excluded.updated_at
                """,
                (
                    business_date, staff_id, staff_name, max(0, int(calls)),
                    max(0, int(dialogues)), max(0, int(tosses)),
                    max(0, min(23, int(captured_hour))), self._now_iso(),
                ),
            )

    def migrate_staff_identity(
        self,
        old_staff_id: str,
        new_staff_id: str,
        new_staff_name: str,
        business_date: Optional[str] = None,
    ) -> None:
        """Merge rows recorded before CTI identity detection into the real staff."""
        if not old_staff_id or old_staff_id == new_staff_id:
            return
        where_date = " AND business_date = ?" if business_date else ""
        date_args = (business_date,) if business_date else ()
        now = self._now_iso()
        with self._lock, self._connect() as conn:
            old_rows = conn.execute(
                "SELECT * FROM hourly_stats WHERE staff_id = ?" + where_date,
                (old_staff_id, *date_args),
            ).fetchall()
            for row in old_rows:
                conn.execute(
                    """
                    INSERT INTO hourly_stats(
                        business_date, hour, staff_id, staff_name, calls, dialogues,
                        tosses, talk_seconds, source, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(business_date, hour, staff_id, source) DO UPDATE SET
                        staff_name = excluded.staff_name,
                        calls = hourly_stats.calls + excluded.calls,
                        dialogues = hourly_stats.dialogues + excluded.dialogues,
                        tosses = hourly_stats.tosses + excluded.tosses,
                        talk_seconds = hourly_stats.talk_seconds + excluded.talk_seconds,
                        updated_at = excluded.updated_at
                    """,
                    (
                        row["business_date"], row["hour"], new_staff_id,
                        new_staff_name, row["calls"], row["dialogues"],
                        row["tosses"], row["talk_seconds"], row["source"], now,
                    ),
                )
            conn.execute(
                "DELETE FROM hourly_stats WHERE staff_id = ?" + where_date,
                (old_staff_id, *date_args),
            )
            conn.execute(
                "UPDATE call_details SET staff_id = ?, staff_name = ?, updated_at = ? "
                "WHERE staff_id = ?" + where_date,
                (new_staff_id, new_staff_name, now, old_staff_id, *date_args),
            )
            # The authoritative CTI value will be written immediately afterwards.
            conn.execute(
                "DELETE FROM daily_totals WHERE staff_id = ?" + where_date,
                (old_staff_id, *date_args),
            )

    def replace_history(self, records: Iterable[Dict[str, Any]]) -> int:
        """Replace CSV-derived records for each included date/staff pair."""
        rows = list(records)
        if not rows:
            return 0
        pairs = sorted({(r["business_date"], r["staff_id"]) for r in rows})
        aggregates: Dict[tuple, Dict[str, Any]] = {}
        for row in rows:
            key = (row["business_date"], int(row["hour"]), row["staff_id"])
            item = aggregates.setdefault(
                key,
                {
                    "business_date": row["business_date"],
                    "hour": int(row["hour"]),
                    "staff_id": row["staff_id"],
                    "staff_name": row["staff_name"],
                    "calls": 0,
                    "dialogues": 0,
                    "tosses": 0,
                    "talk_seconds": 0,
                },
            )
            item["calls"] += 1
            item["dialogues"] += 1 if row.get("is_dialogue") else 0
            item["tosses"] += 1 if row.get("is_toss") else 0
            item["talk_seconds"] += max(0, int(row.get("talk_seconds", 0)))

        now = self._now_iso()
        with self._lock, self._connect() as conn:
            for business_date, staff_id in pairs:
                conn.execute(
                    "DELETE FROM hourly_stats WHERE business_date = ? AND staff_id = ? AND source IN ('cti_history', 'live')",
                    (business_date, staff_id),
                )
                conn.execute(
                    "DELETE FROM call_details WHERE business_date = ? AND staff_id = ? AND source IN ('cti_history', 'live')",
                    (business_date, staff_id),
                )
            for item in aggregates.values():
                conn.execute(
                    """
                    INSERT INTO hourly_stats(
                        business_date, hour, staff_id, staff_name, calls, dialogues,
                        tosses, talk_seconds, source, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'cti_history', ?)
                    """,
                    (
                        item["business_date"], item["hour"], item["staff_id"],
                        item["staff_name"], item["calls"], item["dialogues"],
                        item["tosses"], item["talk_seconds"], now,
                    ),
                )
            for row in rows:
                conn.execute(
                    """
                    INSERT INTO call_details(
                        event_id, business_date, hour, staff_id, staff_name, called_at,
                        result, is_dialogue, is_toss, talk_seconds, source, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'cti_history', ?)
                    """,
                    (
                        row["event_id"], row["business_date"], int(row["hour"]),
                        row["staff_id"], row["staff_name"], row["called_at"],
                        row.get("result", ""), 1 if row.get("is_dialogue") else 0,
                        1 if row.get("is_toss") else 0,
                        max(0, int(row.get("talk_seconds", 0))), now,
                    ),
                )
        return len(rows)

    def _resolved_rows(self, business_date: str) -> tuple[List[sqlite3.Row], List[sqlite3.Row]]:
        with self._lock, self._connect() as conn:
            stats = conn.execute(
                "SELECT * FROM hourly_stats WHERE business_date = ? ORDER BY staff_name, hour",
                (business_date,),
            ).fetchall()
            details = conn.execute(
                "SELECT * FROM call_details WHERE business_date = ? ORDER BY called_at",
                (business_date,),
            ).fetchall()

        return list(stats), list(details)

    def dashboard(self, business_date: str) -> Dict[str, Any]:
        stats, details = self._resolved_rows(business_date)
        with self._lock, self._connect() as conn:
            authoritative_rows = conn.execute(
                "SELECT * FROM daily_totals WHERE business_date = ?",
                (business_date,),
            ).fetchall()
        staff_map: Dict[str, Dict[str, Any]] = {}
        for row in stats:
            staff = staff_map.setdefault(
                row["staff_id"],
                {
                    "staff_id": row["staff_id"],
                    "staff_name": row["staff_name"],
                    "calls": 0,
                    "dialogues": 0,
                    "tosses": 0,
                    "talk_seconds": 0,
                    "corrected": False,
                    "hours": [],
                    "_hours_by_number": {},
                    "details": [],
                },
            )
            hour_number = int(row["hour"])
            hour_row = staff["_hours_by_number"].get(hour_number)
            if hour_row is None:
                hour_row = {
                    "hour": hour_number,
                    "calls": 0,
                    "dialogues": 0,
                    "tosses": 0,
                    "talk_seconds": 0,
                }
                staff["_hours_by_number"][hour_number] = hour_row
                staff["hours"].append(hour_row)
            for key in ("calls", "dialogues", "tosses", "talk_seconds"):
                value = int(row[key])
                hour_row[key] += value
                staff[key] += value

        for row in details:
            staff = staff_map.get(row["staff_id"])
            if staff is None:
                continue
            staff["details"].append(
                {
                    "event_id": row["event_id"],
                    "called_at": row["called_at"],
                    "hour": int(row["hour"]),
                    "result": row["result"],
                    "is_dialogue": bool(row["is_dialogue"]),
                    "is_toss": bool(row["is_toss"]),
                    "talk_seconds": int(row["talk_seconds"]),
                }
            )

        for row in authoritative_rows:
            staff = staff_map.setdefault(
                row["staff_id"],
                {
                    "staff_id": row["staff_id"], "staff_name": row["staff_name"],
                    "calls": 0, "dialogues": 0, "tosses": 0, "talk_seconds": 0,
                    "corrected": False, "hours": [], "_hours_by_number": {}, "details": [],
                },
            )
            staff["staff_name"] = row["staff_name"]
            captured_hour = int(row["captured_hour"])
            if captured_hour not in staff["_hours_by_number"]:
                hour_row = {
                    "hour": captured_hour, "calls": 0, "dialogues": 0,
                    "tosses": 0, "talk_seconds": 0,
                }
                staff["_hours_by_number"][captured_hour] = hour_row
                staff["hours"].append(hour_row)
            for metric in ("calls", "dialogues", "tosses"):
                self._adjust_hours_to_total(
                    staff["hours"], metric, int(row[metric]), captured_hour
                )
                staff[metric] = int(row[metric])

        staff_list = sorted(staff_map.values(), key=lambda item: item["staff_name"])
        totals = {"calls": 0, "dialogues": 0, "tosses": 0, "talk_seconds": 0}
        for staff in staff_list:
            staff.pop("_hours_by_number", None)
            staff["hours"].sort(key=lambda item: item["hour"])
            for key in totals:
                totals[key] += int(staff[key])
        return {"date": business_date, "staff": staff_list, "totals": totals, "source": "local"}

    @staticmethod
    def _adjust_hours_to_total(
        hours: List[Dict[str, Any]], metric: str, target: int, captured_hour: int
    ) -> None:
        current = sum(int(item.get(metric, 0)) for item in hours)
        difference = int(target) - current
        if difference > 0:
            bucket = next(item for item in hours if int(item["hour"]) == captured_hour)
            bucket[metric] = int(bucket.get(metric, 0)) + difference
        elif difference < 0:
            remaining = -difference
            for bucket in sorted(hours, key=lambda item: int(item["hour"]), reverse=True):
                value = int(bucket.get(metric, 0))
                reduction = min(value, remaining)
                bucket[metric] = value - reduction
                remaining -= reduction
                if remaining <= 0:
                    break

    def sync_payload(self, business_date: str) -> Dict[str, List[Dict[str, Any]]]:
        dashboard = self.dashboard(business_date)
        _, details = self._resolved_rows(business_date)
        stats: List[Dict[str, Any]] = []
        for staff in dashboard.get("staff", []):
            for hour in staff.get("hours", []):
                stats.append(
                    {
                        "business_date": business_date,
                        "hour": int(hour["hour"]),
                        "staff_id": staff["staff_id"],
                        "staff_name": staff["staff_name"],
                        "calls": int(hour.get("calls", 0)),
                        "dialogues": int(hour.get("dialogues", 0)),
                        "tosses": int(hour.get("tosses", 0)),
                        "talk_seconds": int(hour.get("talk_seconds", 0)),
                        "source": "resolved",
                        "updated_at": self._now_iso(),
                    }
                )
        return {
            "stats": stats,
            "details": [dict(row) for row in details],
        }

