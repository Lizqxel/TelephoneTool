from datetime import datetime

import pytest

from services.call_statistics import (JST, StatisticsSnapshot, StatisticsStore,
                                      hourly_history_totals, hourly_talk_seconds, next_half_hour, next_hour,
                                      parse_call_duration)
from services.cti_accessibility import msaa_tables
from services.cti_statistics_service import (CollectionSkipped, counter_context, history_total, parse_staff, read_counters,
                                             validate_phone_filter, validate_search_dates, validate_staff_selection)


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


def test_history_excludes_other_staff_and_days_but_rejects_missing_columns():
    row = {"対応スタッフ": "姓　名", "対応日時": "2026/11/13 13:20", "IP電話通話時間": "90"}
    others = [{**row, "対応スタッフ": "別の人"}, {**row, "対応日時": "2026/11/12 13:20"}]
    totals = hourly_history_totals([row, *others], "2026-11-13", "姓 名")
    assert totals["talk_seconds"][13] == 90
    assert totals["counters"][13]["total_calls"] == 1
    assert sum(totals["counters"][h]["total_calls"] for h in range(24)) == 1
    with pytest.raises(ValueError):
        hourly_talk_seconds([{"対応スタッフ": "姓　名"}], "2026-11-13", "姓　名")


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


def counter_nodes(labels, texts):
    nodes = []
    for index, (label, text) in enumerate(zip(labels, texts)):
        x = index * 50
        for offset, (value, y) in enumerate([(label, 80), (text, 94)]):
            nodes.append({"hwnd": index * 2 + offset, "parent_hwnd": 100, "class_name": "WindowsForms10.STATIC", "text": value, "rect": [x, y, x + 48, y + 12]})
    return nodes


def test_counters_use_label_column_and_ignore_percentages():
    nodes = counter_nodes(["総コール数", "新規見込", "対話有NG", "有効コール", "対話無NG"], ["1,234", "12(1.0%)", "2(0.2%)", "40", "3(0.3%)"])
    result = read_counters(nodes)
    assert result == {"total_calls": 1234, "new_prospects": 12, "ng_with_dialogue": 2, "ng_without_dialogue": 3, "valid_calls": 40}


def test_blank_counter_is_unknown_but_missing_label_is_an_error():
    nodes = counter_nodes(["総コール数", "新規見込", "対話有NG", "有効コール", "対話無NG"], ["", "", "", "", ""])
    assert read_counters(nodes) == {k: None for k in ("total_calls", "new_prospects", "ng_with_dialogue", "ng_without_dialogue", "valid_calls")}
    with pytest.raises(RuntimeError, match="有効コール"):
        read_counters(counter_nodes(["総コール数", "新規見込", "対話有NG", "対話無NG"], ["1", "1", "1", "1"]))


def test_staff_selection_allows_other_checked_staff_but_requires_self():
    own = {"スタッフ": "姓　名", "選択": "True"}
    other = {"スタッフ": "別の人", "選択": "False"}
    validate_staff_selection({"rows": [own, other]}, "姓 名")
    validate_staff_selection({"rows": [own, {**other, "選択": "True"}]}, "姓 名")
    with pytest.raises(CollectionSkipped, match="本人にチェック"):
        validate_staff_selection({"rows": [{**own, "選択": "False"}, {**other, "選択": "True"}]}, "姓 名")
    with pytest.raises(RuntimeError, match="特定できません"):
        validate_staff_selection({"rows": [other]}, "姓 名")


def date_nodes(start, end):
    return [{"class_name": "WindowsForms10.SysDateTimePick32", "text": start},
            {"class_name": "WindowsForms10.SysDateTimePick32", "text": end}]


def test_search_dates_may_cover_more_than_today_but_must_include_the_whole_day():
    validate_search_dates(date_nodes("2026-11-13 00:00:00", "2026-11-13 23:59:59"), "2026-11-13")
    validate_search_dates(date_nodes("2026-11-01 00:00:00", "2026-11-30 23:59:59"), "2026-11-13")
    for start, end in (("2026-10-01 00:00:00", "2026-10-31 23:59:59"), ("2026-11-13 09:00:00", "2026-11-13 23:59:59"),
                       ("2026-11-13 00:00:00", "2026-11-13 18:00:00")):
        with pytest.raises(CollectionSkipped, match="当日"):
            validate_search_dates(date_nodes(start, end), "2026-11-13")


def test_history_total_reads_the_whole_search_count_not_the_page_count():
    nodes = [
        {"hwnd": 1, "parent_hwnd": 9, "class_name": "WindowsForms10.STATIC", "text": "検索結果件数", "rect": [0, 0, 990, 20]},
        {"hwnd": 2, "parent_hwnd": 9, "class_name": "WindowsForms10.STATIC", "text": "2,043", "rect": [88, 2, 173, 18]},
        {"hwnd": 3, "parent_hwnd": 9, "class_name": "WindowsForms10.STATIC", "text": "464", "rect": [115, 22, 165, 40]},
        {"hwnd": 4, "parent_hwnd": 8, "class_name": "WindowsForms10.STATIC", "text": "50 件", "rect": [845, 52, 962, 68]},
    ]
    assert history_total(nodes) == 2043


def test_phone_filter_uses_its_label_and_ignores_unrelated_memo():
    nodes = [
        {"text": "電話番号", "parent_hwnd": 1, "class_name": "STATIC", "rect": [10, 10, 60, 24]},
        {"text": "", "parent_hwnd": 1, "class_name": "EDIT", "rect": [62, 10, 150, 24]},
        {"text": "顧客メモ", "parent_hwnd": 2, "class_name": "RICHEDIT", "rect": [62, 10, 150, 24]},
    ]
    validate_phone_filter(nodes)
    nodes[1]["text"] = "09012345678"
    with pytest.raises(CollectionSkipped, match="空に"):
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
    # NGは現アナNG・多重留守NG以外を対話有NGとして数える。
    assert totals["counters"][14]["ng_with_dialogue"] == 1


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


def test_other_ng_details_count_as_ng_with_dialogue():
    row = {**history_row(outcome="NG"), "架電結果詳細": "未確認の理由"}
    counts = hourly_history_totals([row], "2026-11-13", "姓　名")["counters"][13]
    assert counts["ng_with_dialogue"] == 1
    assert counts["ng_without_dialogue"] == 0
    assert counts["dialogues"] == 1


def prospect_row(when, number, outcome="見込", detail="-", call_id=None):
    return {**history_row(when, outcome, "30", call_id or f"call-{number}-{when}"), "架電結果詳細": detail, "管理番号": number}


def test_toss_counts_new_prospects_whose_current_rank_is_toss_once_per_number():
    rows = [prospect_row("10:05", "A"), prospect_row("11:40", "A"),  # 同じ見込への再架電は最初の時間帯に1件
            prospect_row("10:20", "B"), prospect_row("11:10", "C"),
            prospect_row("12:00", "D", "継続見込"), prospect_row("12:30", "E", "見込", "継続"),
            prospect_row("13:00", "F", "新規見込"), history_row("14:00", "留守", call_id="x")]
    ranks = {"A": "トス", "B": "C", "C": "ﾄｽ", "D": "トス", "E": "トス", "F": "トス"}
    totals = hourly_history_totals(rows, "2026-11-13", "姓　名", ranks=ranks)
    toss = {h: totals["counters"][h]["toss"] for h in (10, 11, 12, 13, 14)}
    assert toss == {10: 1, 11: 1, 12: 0, 13: 1, 14: 0}
    assert totals["unmatched_prospects"] == 0


def test_toss_is_unknown_only_in_hours_with_prospects_when_ranks_are_unavailable():
    rows = [prospect_row("10:05", "A"), history_row("11:00", "留守", call_id="x")]
    totals = hourly_history_totals(rows, "2026-11-13", "姓　名", ranks=None)
    assert totals["counters"][10]["toss"] is None
    assert totals["counters"][11]["toss"] == 0
    missing = hourly_history_totals(rows, "2026-11-13", "姓　名", ranks={})
    assert missing["counters"][10]["toss"] == 0
    assert missing["unmatched_prospects"] == 1


def counter_snapshot(when, valid, total=10, session="cti-1"):
    stamp = datetime.fromisoformat(when).replace(tzinfo=JST)
    return StatisticsSnapshot(stamp, "staff-1", "姓　名", session,
                              {"total_calls": total, "new_prospects": 0, "ng_with_dialogue": 0, "ng_without_dialogue": 0,
                               "valid_calls": valid}, None, 0)


def test_valid_calls_accumulate_per_hour_from_five_minute_counter_reads(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    for when, valid in (("09:00:02", 3), ("09:05:02", 5), ("09:55:02", 9), ("10:00:02", 10), ("10:05:02", 14)):
        store.save(counter_snapshot(f"2026-11-13T{when}", valid))
    rows = store.hours("staff-1", "2026-11-13")
    assert rows[9]["valid_calls"] == 7 and rows[9]["counter_status"] == "差分"
    assert rows[10]["valid_calls"] == 4
    # 履歴の無い回は応答時間を上書きしない。
    assert rows[9]["talk_seconds"] is None


def test_blank_valid_counter_skips_only_that_counter(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(counter_snapshot("2026-11-13T09:00:02", None, 10))
    store.save(counter_snapshot("2026-11-13T09:05:02", 4, 12))
    store.save(counter_snapshot("2026-11-13T09:10:02", 6, 15))
    row = store.hours("staff-1", "2026-11-13")[9]
    assert row["total_calls"] == 5
    assert row["valid_calls"] == 2


def test_history_and_valid_counter_status_are_both_kept(tmp_path):
    store = StatisticsStore(tmp_path / "stats.sqlite3")
    store.save(counter_snapshot("2026-11-13T13:00:02", 1))
    store.save(counter_snapshot("2026-11-13T13:30:02", 4))
    base = history_snapshot("2026-11-13T14:00:02", [prospect_row("13:10", "A")])
    totals = hourly_history_totals([prospect_row("13:10", "A")], "2026-11-13", "姓　名", ranks={"A": "トス"})
    store.save(StatisticsSnapshot(base.acquired_at, base.staff_id, base.staff_name, "cti-1",
                                  {**counter_snapshot("2026-11-13T14:00:02", 6).counters}, totals["talk_seconds"], 1, totals["counters"], totals["first_hour"]))
    row = store.hours("staff-1", "2026-11-13")[13]
    assert row["valid_calls"] == 5 and row["toss"] == 1 and row["total_calls"] == 1 and row["dialogues"] == 1
    assert row["counter_status"] == "履歴＋差分／未判定"  # 見込/- は新規・継続を区別できない


def test_existing_database_without_new_columns_is_migrated_in_place(tmp_path):
    import sqlite3
    path = tmp_path / "old.sqlite3"
    db = sqlite3.connect(path)
    db.executescript("""
        CREATE TABLE snapshots (staff_id TEXT NOT NULL, day TEXT NOT NULL, staff_name TEXT NOT NULL,
            observed_at TEXT NOT NULL, session TEXT NOT NULL, total_calls INTEGER, new_prospects INTEGER,
            ng_with_dialogue INTEGER, ng_without_dialogue INTEGER, PRIMARY KEY(staff_id,day));
        CREATE TABLE hourly (staff_id TEXT NOT NULL, day TEXT NOT NULL, hour INTEGER NOT NULL,
            talk_seconds REAL, total_calls INTEGER, new_prospects INTEGER, ng_with_dialogue INTEGER,
            ng_without_dialogue INTEGER, counter_status TEXT NOT NULL DEFAULT '未取得',
            talk_updated_at TEXT, counter_updated_at TEXT, PRIMARY KEY(staff_id,day,hour));
        CREATE TABLE history_hourly (staff_id TEXT NOT NULL, day TEXT NOT NULL, hour INTEGER NOT NULL,
            total_calls INTEGER, new_prospects INTEGER, ng_with_dialogue INTEGER, ng_without_dialogue INTEGER,
            PRIMARY KEY(staff_id,day,hour));
        INSERT INTO snapshots VALUES ('staff-1','2026-11-13','姓　名','2026-11-13T13:00:02+09:00','cti-1',5,0,0,0);
        INSERT INTO hourly(staff_id,day,hour,total_calls,counter_status) VALUES ('staff-1','2026-11-13',12,3,'差分');
    """)
    db.commit()
    db.close()
    store = StatisticsStore(path)
    assert store.hours("staff-1", "2026-11-13")[12]["total_calls"] == 3
    store.save(counter_snapshot("2026-11-13T13:05:02", 2, 7))
    assert store.latest("staff-1", "2026-11-13")["valid_calls"] == 2


def test_hour_schedule_aligns_to_the_wall_clock():
    assert next_hour(datetime(2026, 11, 13, 13, 57, 30, tzinfo=JST)) == datetime(2026, 11, 13, 14, 0, tzinfo=JST)
    assert next_hour(datetime(2026, 11, 13, 14, 0, 2, tzinfo=JST)) == datetime(2026, 11, 13, 15, 0, tzinfo=JST)
    assert next_hour(datetime(2026, 11, 13, 23, 58, tzinfo=JST)).date().isoformat() == "2026-11-14"


def msaa_result(headers, rows):
    nodes = [{"id": 0, "parent_id": -1, "role": 24, "name": "", "value": None}]
    header_row = {"id": 1, "parent_id": 0, "role": 28, "name": "", "value": None}
    nodes.append(header_row)
    for name in headers:
        nodes.append({"id": len(nodes), "parent_id": 1, "role": 25, "name": name, "value": name})
    for index, values in enumerate(rows):
        row = {"id": len(nodes), "parent_id": 0, "role": 28, "name": "", "value": None}
        nodes.append(row)
        for name, value in zip(headers, values):
            nodes.append({"id": len(nodes), "parent_id": row["id"], "role": 29, "name": f"{name} 行 {index}", "value": value})
    return {"hwnd": 5, "nodes": nodes, "truncated": False, "errors": []}


def test_duplicate_column_names_are_read_by_position():
    headers = ["選択", "管理番号", "スタッフ", "ランク", "スタッフ"]
    table = msaa_tables([msaa_result(headers, [["False", "0171_1", "姓 名", "トス", "bb01"]])])[0]
    assert table["headers"] == ["選択", "管理番号", "スタッフ", "ランク", "スタッフ#2"]
    assert table["rows"] == [{"選択": "False", "管理番号": "0171_1", "スタッフ": "姓 名", "ランク": "トス", "スタッフ#2": "bb01"}]
    unique = msaa_tables([msaa_result(["対応日時", "管理番号"], [["2026/11/13 10:00", "0171_1"]])])[0]
    assert unique["rows"] == [{"対応日時": "2026/11/13 10:00", "管理番号": "0171_1"}]


def test_dialogues_are_prospects_plus_ng_with_dialogue_from_history():
    rows = [prospect_row("10:05", "A"), prospect_row("10:10", "B", "継続見込"),
            {**history_row("10:20", "NG", call_id="n1"), "架電結果詳細": "アプローチNG"},
            {**history_row("10:30", "NG", call_id="n2"), "架電結果詳細": "対話無NG"},
            history_row("10:40", "留守", call_id="r1"), history_row("11:00", "留守", call_id="r2"),
            {**history_row("12:00", "NG", call_id="n3"), "架電結果詳細": "現アナNG"},
            {**history_row("12:10", "NG", call_id="n4"), "架電結果詳細": "多重留守NG"},
            {**history_row("12:20", "NG", call_id="n5"), "架電結果詳細": "申込意思NG"},
            history_row("13:00", "未知の結果", call_id="u1")]
    totals = hourly_history_totals(rows, "2026-11-13", "姓　名", ranks={})
    assert totals["counters"][10]["dialogues"] == 3
    assert totals["counters"][11]["dialogues"] == 0
    assert totals["counters"][12]["dialogues"] == 1
    assert totals["counters"][12]["ng_with_dialogue"] == 1 and totals["counters"][12]["ng_without_dialogue"] == 2
    # 未知の架電結果は推測せず、その時間帯だけ未判定にして結果名を返す。
    assert totals["counters"][13]["dialogues"] is None
    assert totals["unknown_dialogue_results"] == {"未知の結果": 1}


def test_ng_detail_matching_ignores_width_and_spaces():
    rows = [{**history_row("12:00", "NG", call_id="n3"), "架電結果詳細": "現ｱﾅNG"},
            {**history_row("12:10", "NG", call_id="n4"), "架電結果詳細": "多重留守 ＮＧ"}]
    counts = hourly_history_totals(rows, "2026-11-13", "姓　名", ranks={})["counters"][12]
    assert counts["dialogues"] == 0 and counts["ng_without_dialogue"] == 2


def test_login_name_is_parsed_from_status_bar_items():
    assert parse_staff(["管理メニュー", "姓　名 さん：bb0110326", "【Ver.00.00.507】"]) == ("姓　名", "bb0110326")
    with pytest.raises(RuntimeError):
        parse_staff(["管理メニュー"])


def test_sheet_sender_posts_token_and_json_to_the_form():
    import json
    from services.call_statistics_sheet import SheetSender, load_sheet_settings
    form, token = load_sheet_settings([
        ("gform_settings.json", {"googleFormPosting": {"tokenValue": "shared"}}),
        ("settings.json", {"callStatisticsSheet": {"formResponseUrl": "https://docs.google.com/forms/d/e/X/formResponse",
                                                   "tokenEntry": "entry.1", "dataEntry": "entry.2"}}),
    ])
    assert token == "shared" and form["tokenEntry"] == "entry.1"
    sent = {}

    class Response:
        status_code, url = 200, "https://docs.google.com/forms/d/e/X/formResponse"

    def post(url, data, timeout):
        sent.update(url=url, data=data)
        return Response()

    SheetSender(form, token, post=post).send("2026-11-13", "bb1", "姓 名", [{"hour": 10, "calls": 3, "dialogues": None, "toss": 1}])
    assert sent["url"].endswith("/formResponse") and sent["data"]["entry.1"] == "shared"
    assert json.loads(sent["data"]["entry.2"]) == {"date": "2026-11-13", "staffId": "bb1", "staffName": "姓 名",
                                                   "hours": [{"hour": 10, "calls": 3, "dialogues": None, "toss": 1}]}
    Response.url = "https://accounts.google.com/v3/signin"
    with pytest.raises(RuntimeError, match="ログイン必須"):
        SheetSender(form, token, post=post).send("2026-11-13", "bb1", "姓 名", [])
    Response.url, Response.status_code = "https://docs.google.com/forms/x", 400
    with pytest.raises(RuntimeError, match="HTTP 400"):
        SheetSender(form, token, post=post).send("2026-11-13", "bb1", "姓 名", [])


def test_sheet_sender_is_disabled_until_the_form_is_configured(monkeypatch):
    from services import call_statistics_sheet as module
    monkeypatch.setattr(module, "_settings_candidates", lambda: [("s", {"googleFormPosting": {"tokenValue": "t"}})])
    assert module.SheetSender.from_settings() is not None  # 送信先は既定値で設定済み
    monkeypatch.setattr(module, "_settings_candidates", lambda: [("s", {"googleFormPosting": {"tokenValue": ""}})])
    with pytest.raises(RuntimeError, match="共有トークンが未設定"):
        module.SheetSender.from_settings()
    monkeypatch.setattr(module, "DEFAULT_FORM", {"formResponseUrl": "", "tokenEntry": "", "dataEntry": ""})
    assert module.SheetSender.from_settings() is None
