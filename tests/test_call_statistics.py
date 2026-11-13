from datetime import datetime

import pytest

from services.call_statistics import (JST, StatisticsSnapshot, StatisticsStore,
                                      hourly_history_totals, hourly_talk_seconds, next_half_hour, parse_call_duration)
from services.cti_statistics_service import counter_context, read_counters, validate_phone_filter, validate_staff_selection


def snapshot(when, total, seconds=60, session="cti-1"):
    stamp = datetime.fromisoformat(when).replace(tzinfo=JST)
    return StatisticsSnapshot(stamp, "staff-1", "姓　名", session,
                              {"total_calls": total, "new_prospects": total // 5,
                               "ng_with_dialogue": total // 3, "ng_without_dialogue": 0},
                              {stamp.hour: seconds}, 1)


def test_half_hour_scheduler_aligns_wall_clock_and_rolls_day():
    assert next_half_hour(datetime(2026, 11, 13, 13, 9, tzinfo=JST)).hour == 13
    assert next_half_hour(datetime(2026, 11, 13, 13, 9, tzinfo=JST)).minute == 30
    assert next_half_hour(datetime(2026, 11, 13, 13, 30, tzinfo=JST)).hour == 14
    assert next_half_hour(datetime(2026, 11, 13, 23, 59, tzinfo=JST)).date().isoformat() == "2026-11-14"


def test_duration_formats_and_units():
    assert parse_call_duration("1:02:03") == 3723
    assert parse_call_duration("02:03") == 123
    assert parse_call_duration("1.5", "minutes") == 90
    assert parse_call_duration("90秒") == 90
    assert parse_call_duration("2分") == 120
    for invalid in ("-10", "NaN", "Infinity", "1:90", "未取得"):
        with pytest.raises(ValueError):
            parse_call_duration(invalid)


def test_history_hours_deduplicate_call_ids_but_not_unidentified_calls():
    base = {"対応スタッフ": "姓 名", "対応日時": "2026/11/13 13:20", "IP電話通話時間": "90", "通話ユニークID": "call-1"}
    second = {**base, "対応日時": "2026/11/13 14:00", "IP電話通話時間": "01:00", "通話ユニークID": ""}
    totals = hourly_talk_seconds([base, base, second, second], "2026-11-13", "姓　名")
    assert totals[13] == 90
    assert totals[14] == 120


def test_history_rejects_different_staff_day_and_missing_columns():
    row = {"対応スタッフ": "姓　名", "対応日時": "2026/11/13 13:20", "IP電話通話時間": "90"}
    for changed in ({**row, "対応スタッフ": "別の人"}, {**row, "対応日時": "2026/11/12 13:20"}, {"対応スタッフ": "姓　名"}):
        with pytest.raises(ValueError):
            hourly_talk_seconds([changed], "2026-11-13", "姓　名")


def test_first_snapshot_is_baseline_not_historical_counter_totals(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:10:00", 100))
    rows = store.hours("staff-1", "2026-11-13")
    assert all(r["total_calls"] is None for r in rows)
    assert rows[13]["talk_seconds"] == 60
    assert rows[14]["talk_seconds"] is None
    assert store.latest("staff-1", "2026-11-13")["total_calls"] == 100


def test_manual_and_half_hour_refresh_do_not_double_count(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:00:04", 100, 60))
    store.save(snapshot("2026-11-13T13:30:04", 110, 120))
    store.save(snapshot("2026-11-13T13:40:00", 113, 150))
    store.save(snapshot("2026-11-13T14:00:05", 120, 0))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["total_calls"] == 20
    # Later complete-day history reads replace each hour, rather than adding again.
    assert row["talk_seconds"] == 0
    assert row["counter_status"] == "差分"


def test_first_partial_hour_is_marked_partial_and_future_reads_overwrite_talk(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:10:00", 100))
    store.save(snapshot("2026-11-13T13:30:00", 110, 90))
    store.save(snapshot("2026-11-13T13:45:00", 112, 100))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["total_calls"] == 12
    assert row["talk_seconds"] == 100
    assert row["counter_status"] == "一部の差分"


def test_missing_hour_boundary_is_not_arbitrarily_allocated(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:30:00", 100))
    store.save(snapshot("2026-11-13T14:15:00", 120))
    rows = store.hours("staff-1", "2026-11-13")
    assert rows[13]["total_calls"] is None and rows[14]["total_calls"] is None
    assert "時間またぎ" in rows[13]["counter_status"]


def test_counter_reset_does_not_record_negative_or_false_zero(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:00:00", 100))
    store.save(snapshot("2026-11-13T13:30:00", 2))
    store.save(snapshot("2026-11-13T14:00:00", 7))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["total_calls"] == 5
    assert row["counter_status"] == "リセットあり"


def test_database_survives_restart_and_separates_dates_and_staff(tmp_path):
    path = tmp_path / "stats.sqlite3"
    store = StatisticsStore(path)
    store.save(snapshot("2026-11-13T23:30:00", 100))
    store = StatisticsStore(path)
    store.save(snapshot("2026-11-14T00:00:05", 0))
    assert store.latest("staff-1", "2026-11-13")["total_calls"] == 100
    assert store.hours("staff-1", "2026-11-14")[0]["total_calls"] is None
    assert store.latest("other-staff", "2026-11-13") is None
    with pytest.raises(ValueError):
        store.save(snapshot("2026-11-14T00:00:01", 0))


def test_counters_use_label_column_and_ignore_percentages():
    nodes = []
    for index, (label, text) in enumerate(zip(["総コール数", "新規見込", "対話有NG", "対話無NG"], ["1,234", "12(1.0%)", "2(0.2%)", "3(0.3%)"])):
        x = index * 50
        for offset, (value, y) in enumerate([(label, 80), (text, 94)]):
            nodes.append({"hwnd": index * 2 + offset, "parent_hwnd": 100, "class_name": "WindowsForms10.STATIC", "text": value, "rect": [x, y, x + 48, y + 12]})
    result = read_counters(nodes)
    assert result == {"total_calls": 1234, "new_prospects": 12, "ng_with_dialogue": 2, "ng_without_dialogue": 3}


def test_staff_selection_requires_only_the_logged_in_staff():
    own = {"スタッフ": "姓　名", "選択": "True"}
    other = {"スタッフ": "別の人", "選択": "False"}
    validate_staff_selection({"rows": [own, other]}, "姓 名")
    with pytest.raises(RuntimeError):
        validate_staff_selection({"rows": [own, {**other, "選択": "True"}]}, "姓 名")


def test_phone_filter_uses_its_label_and_ignores_unrelated_memo():
    nodes = [
        {"text": "電話番号", "parent_hwnd": 1, "class_name": "STATIC", "rect": [10, 10, 60, 24]},
        {"text": "", "parent_hwnd": 1, "class_name": "EDIT", "rect": [62, 10, 150, 24]},
        {"text": "顧客メモ", "parent_hwnd": 2, "class_name": "RICHEDIT", "rect": [62, 10, 150, 24]},
    ]
    validate_phone_filter(nodes)
    nodes[1]["text"] = "09012345678"
    with pytest.raises(RuntimeError, match="空に"):
        validate_phone_filter(nodes)


def test_new_counter_context_does_not_subtract_unrelated_larger_totals(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(snapshot("2026-11-13T13:00:00", 10, session="list-a"))
    store.save(snapshot("2026-11-13T13:30:00", 100, session="list-b"))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["total_calls"] is None
    assert row["counter_status"] == "リセットあり"


def test_list_context_ignores_customer_fields():
    nodes = [
        {"text": "リスト", "visible": True, "parent_hwnd": 1, "class_name": "STATIC", "rect": [10, 10, 60, 24]},
        {"text": "架電リストA", "visible": True, "parent_hwnd": 1, "class_name": "COMBOBOX", "rect": [62, 10, 150, 24]},
        {"text": "顧客の選択値", "visible": True, "parent_hwnd": 2, "class_name": "COMBOBOX", "rect": [62, 100, 150, 120]},
    ]
    old = counter_context(nodes)
    nodes[2]["text"] = "別の顧客の選択値"
    assert counter_context(nodes) == old
    nodes[1]["text"] = "架電リストB"
    assert counter_context(nodes) != old


def history_snapshot(when, rows):
    original = snapshot(when, len(rows))
    totals = hourly_history_totals(rows, "2026-11-13", original.staff_name)
    return StatisticsSnapshot(original.acquired_at, original.staff_id, original.staff_name, original.session,
                              original.counters, totals["talk_seconds"], len(rows), totals["counters"], totals["first_hour"])


def history_row(when="13:53", outcome="留守", seconds="0", call_id="call-1"):
    return {"対応スタッフ": "姓　名", "対応日時": f"2026/11/13 {when}", "IP電話通話時間": seconds,
            "架電結果": outcome, "架電結果詳細": "-", "通話ユニークID": call_id}


def test_historical_zero_duration_call_is_counted_before_first_collection(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    data = [history_row()]
    store.save(history_snapshot("2026-11-13T16:10:00", data))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["talk_seconds"] == 0
    assert row["total_calls"] == 1
    assert row["new_prospects"] == row["ng_with_dialogue"] == 0
    assert row["counter_status"] == "履歴集計"
    store.save(history_snapshot("2026-11-13T16:30:00", data))
    assert store.hours("staff-1", "2026-11-13")[13]["total_calls"] == 1


def test_display_starts_at_first_log_even_when_it_has_zero_talk_duration(tmp_path):
    path = tmp_path / "stats.sqlite3"
    store = StatisticsStore(path)
    store.save(history_snapshot("2026-11-13T16:10:00", [history_row("12:00"), history_row(call_id="call-2")]))
    store = StatisticsStore(path)
    rows = store.display_hours("staff-1", "2026-11-13")
    assert rows[0]["hour"] == 12
    assert len(rows) == 12
    assert rows[0]["total_calls"] == rows[1]["total_calls"] == 1
    assert all(row["hour"] >= 12 for row in rows)


def test_historical_outcomes_require_explicit_labels():
    rows = [history_row(outcome="新規見込"), history_row(outcome="対話有NG", call_id="call-2"),
            history_row(when="14:10", outcome="NG", call_id="call-3")]
    totals = hourly_history_totals(rows, "2026-11-13", "姓　名")
    assert totals["counters"][13]["new_prospects"] == 1
    assert totals["counters"][13]["ng_with_dialogue"] == 1
    assert totals["counters"][14]["total_calls"] == 1
    assert totals["counters"][14]["ng_with_dialogue"] is None


def test_backfills_existing_database_without_discarding_counter_baseline(tmp_path):
    path = tmp_path / "stats.sqlite3"
    store = StatisticsStore(path)
    store.save(snapshot("2026-11-13T16:10:00", 1, 0))
    assert store.hours("staff-1", "2026-11-13")[13]["total_calls"] is None
    store = StatisticsStore(path)
    store.save(history_snapshot("2026-11-13T16:30:00", [history_row()]))
    assert store.hours("staff-1", "2026-11-13")[13]["total_calls"] == 1
    assert store.latest("staff-1", "2026-11-13")["total_calls"] == 1
    assert store.display_hours("staff-1", "2026-11-13")[0]["hour"] == 13


@pytest.mark.parametrize("missing_id", [None, "", "-", "0", "(Null)", "(null)", " (NULL) ", "null", "（Ｎｕｌｌ）"])
def test_unset_call_ids_do_not_collapse_distinct_history_rows(missing_id):
    rows = [history_row(call_id=missing_id),
            {**history_row("16:34", "NG", "7", missing_id), "架電結果詳細": "アプローチNG"}]
    totals = hourly_history_totals(rows, "2026-11-13", "姓　名")
    assert totals["counters"][13]["total_calls"] == 1
    assert totals["counters"][16]["total_calls"] == 1
    assert totals["counters"][16]["ng_with_dialogue"] == 1
    assert totals["talk_seconds"][16] == 7


def test_real_call_ids_still_deduplicate_and_reject_conflicting_data():
    row = history_row(seconds="7", call_id="valid-call-123")
    totals = hourly_history_totals([row, row], "2026-11-13", "姓　名")
    assert totals["counters"][13]["total_calls"] == 1
    assert totals["talk_seconds"][13] == 7
    with pytest.raises(ValueError, match="同じ通話ID"):
        hourly_history_totals([row, {**row, "IP電話通話時間": "10"}], "2026-11-13", "姓　名")
    distinct = {**row, "通話ユニークID": "VALID-call-123"}
    assert hourly_history_totals([row, distinct], "2026-11-13", "姓　名")["counters"][13]["total_calls"] == 2


def test_approach_ng_reaggregation_updates_old_totals_without_double_counting(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    old = history_row(call_id="(Null)")
    ng = {**history_row("16:34", "NG", "7", "(Null)"), "架電結果詳細": "アプローチNG"}
    store.save(history_snapshot("2026-11-13T16:30:00", [old]))
    store.save(history_snapshot("2026-11-13T16:40:00", [ng, old]))
    store.save(history_snapshot("2026-11-13T16:50:00", [ng, old]))
    rows = store.display_hours("staff-1", "2026-11-13")
    hour = next(row for row in rows if row["hour"] == 16)
    assert sum(row["total_calls"] or 0 for row in rows) == 2
    assert hour["total_calls"] == hour["ng_with_dialogue"] == 1
    assert hour["talk_seconds"] == 7
    assert hour["ng_without_dialogue"] == 0


def test_unknown_ng_detail_remains_unclassified():
    row = {**history_row(outcome="NG"), "架電結果詳細": "未確認の理由"}
    counts = hourly_history_totals([row], "2026-11-13", "姓　名")["counters"][13]
    assert counts["ng_with_dialogue"] is None
    assert counts["ng_without_dialogue"] is None
