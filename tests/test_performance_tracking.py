import csv
from pathlib import Path
import uuid
from datetime import datetime, timedelta, timezone

from services.cti_performance import (
    CTIHistoryCSVImporter,
    parse_duration_seconds,
    parse_identity_text,
)
from services.performance_store import PerformanceStore
from services.performance_manager import PerformanceManager


def test_parse_cti_identity_from_status_bar():
    assert parse_identity_text("初山　健人さん：bb0102399") == ("bb0102399", "初山 健人")


def test_parse_ip_phone_duration():
    assert parse_duration_seconds("125") == 125
    assert parse_duration_seconds("04:31") == 271
    assert parse_duration_seconds("1:02:03") == 3723
    assert parse_duration_seconds("-") == 0


def _work_path(name: str) -> Path:
    root = Path.cwd() / "work" / "test-performance-files"
    root.mkdir(parents=True, exist_ok=True)
    return root / f"{uuid.uuid4().hex}-{name}"


def test_import_history_csv_without_retaining_customer_fields():
    path = _work_path("history.csv")
    rows = [
        {
            "対応日時": "2026/11/13 14:42:00",
            "管理番号": "0171_240304_00011467",
            "電話番号": "0298926160",
            "対応スタッフ": "初山 健人",
            "IP電話通話時間": "06:18",
            "架電結果": "見込",
            "架電結果詳細": "新規見込",
        }
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    records = CTIHistoryCSVImporter().parse(str(path), "bb0102399", "初山 健人")
    assert len(records) == 1
    record = records[0]
    assert record["business_date"] == "2026-11-13"
    assert record["hour"] == 14
    assert record["talk_seconds"] == 378
    assert record["is_dialogue"] is True
    assert record["is_toss"] is True
    serialized = repr(record)
    assert "0298926160" not in serialized
    assert "0171_240304_00011467" not in serialized


def test_history_replaces_live_counts_for_same_staff_and_day():
    store = PerformanceStore(str(_work_path("performance.db")))
    store.add_delta("2026-11-13", 14, "bb01", "担当者", calls=5, dialogues=2, source="live")
    store.replace_history(
        [
            {
                "event_id": "csv-one",
                "business_date": "2026-11-13",
                "hour": 14,
                "staff_id": "bb01",
                "staff_name": "担当者",
                "called_at": "2026-11-13T14:10:00+09:00",
                "result": "留守",
                "is_dialogue": False,
                "is_toss": False,
                "talk_seconds": 0,
            }
        ]
    )
    dashboard = store.dashboard("2026-11-13")
    assert dashboard["totals"]["calls"] == 1
    assert dashboard["totals"]["dialogues"] == 0
    assert dashboard["staff"][0]["details"][0]["result"] == "留守"


def test_hourly_delta_is_idempotent_only_for_absolute_sync_not_local_events():
    store = PerformanceStore(str(_work_path("performance.db")))
    store.add_delta("2026-11-13", 10, "bb01", "担当者", calls=1)
    store.add_delta("2026-11-13", 10, "bb01", "担当者", calls=1, dialogues=1, talk_seconds=60)
    dashboard = store.dashboard("2026-11-13")
    assert dashboard["totals"] == {"calls": 2, "dialogues": 1, "tosses": 0, "talk_seconds": 60}


def test_live_calls_after_history_import_are_added_to_same_hour():
    store = PerformanceStore(str(_work_path("performance.db")))
    store.replace_history(
        [{
            "event_id": "csv-one", "business_date": "2026-11-13", "hour": 14,
            "staff_id": "bb01", "staff_name": "担当者",
            "called_at": "2026-11-13T14:10:00+09:00", "result": "留守",
            "is_dialogue": False, "is_toss": False, "talk_seconds": 0,
        }]
    )
    store.add_delta("2026-11-13", 14, "bb01", "担当者", calls=1, dialogues=1, talk_seconds=90)
    dashboard = store.dashboard("2026-11-13")
    assert dashboard["totals"]["calls"] == 2
    assert dashboard["totals"]["dialogues"] == 1
    assert dashboard["staff"][0]["hours"] == [
        {"hour": 14, "calls": 2, "dialogues": 1, "tosses": 0, "talk_seconds": 90}
    ]


def test_cti_daily_totals_override_incomplete_live_detection():
    store = PerformanceStore(str(_work_path("performance.db")))
    store.add_delta("2026-11-13", 14, "bb01", "担当者", calls=2)
    store.set_daily_totals("2026-11-13", "bb01", "担当者", 3, 1, 1, 14)

    dashboard = store.dashboard("2026-11-13")

    assert dashboard["totals"]["calls"] == 3
    assert dashboard["totals"]["dialogues"] == 1
    assert dashboard["totals"]["tosses"] == 1
    assert dashboard["staff"][0]["hours"][0]["calls"] == 3


def test_pending_identity_rows_are_merged_into_detected_staff():
    store = PerformanceStore(str(_work_path("performance.db")))
    store.add_delta("2026-11-13", 14, "device:abc", "担当者取得待ち", calls=2)
    store.migrate_staff_identity(
        "device:abc", "bb0102399", "初山 健人", business_date="2026-11-13"
    )

    dashboard = store.dashboard("2026-11-13")

    assert len(dashboard["staff"]) == 1
    assert dashboard["staff"][0]["staff_id"] == "bb0102399"
    assert dashboard["staff"][0]["staff_name"] == "初山 健人"
    assert dashboard["totals"]["calls"] == 2


def test_manager_maps_cti_status_and_mikomi_to_defined_metrics():
    manager = PerformanceManager(
        {"performanceTracking": {"enabled": False, "fallbackStaffId": "bb01", "fallbackStaffName": "担当者"}},
        None,
        db_path=str(_work_path("performance.db")),
    )
    manager.staff_id = "bb01"
    manager.staff_name = "担当者"
    manager._handle_status_change("待ち受け中", "発信中")
    manager._handle_status_change("発信中", "通話中")
    manager._handle_status_change("通話中", "待ち受け中")
    manager.on_cti_action("見込")
    day = datetime.now(timezone(timedelta(hours=9))).date().isoformat()
    totals = manager.store.dashboard(day)["totals"]
    assert totals["calls"] == 1
    assert totals["dialogues"] == 1
    assert totals["tosses"] == 1

