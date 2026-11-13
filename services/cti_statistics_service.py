"""CTIを表示切替せず検索し、集計に必要な値だけを返す。"""

import ctypes
from datetime import datetime
import hashlib
import re
import threading
import time
from types import SimpleNamespace

from services.call_statistics import JST, StatisticsSnapshot, hourly_history_totals, normalize_name
from services.cti_accessibility import Windows, msaa_tables, read_msaa, uia_snapshot


def descendants(nodes, root):
    ids = {root}
    for node in nodes:
        if node["parent_hwnd"] in ids:
            ids.add(node["hwnd"])
    return [n for n in nodes if n["hwnd"] in ids]


def read_staff(nodes):
    parents = {n["id"] for n in nodes if n["control_type"] == "ControlType.StatusBar"}
    matches = []
    for node in nodes:
        if node["parent_id"] in parents:
            match = re.fullmatch(r"(.+?)\s*さん\s*[:：]\s*(\S+)", node["name"].strip())
            if match:
                matches.append((match.group(1).strip(), match.group(2)))
    if len(matches) != 1:
        raise RuntimeError("CTI下部のログインスタッフ名を特定できませんでした。")
    return matches[0]


def read_counters(nodes):
    mapping = {"総コール数": "total_calls", "新規見込": "new_prospects",
               "対話有NG": "ng_with_dialogue", "対話無NG": "ng_without_dialogue"}
    result = {}
    for label, key in mapping.items():
        labels = [n for n in nodes if n["text"].strip() == label and "STATIC" in n["class_name"]]
        if len(labels) != 1:
            raise RuntimeError(f"CTIの「{label}」を特定できませんでした。")
        source = labels[0]
        left, top, right, bottom = source["rect"]
        candidates = []
        for node in nodes:
            if node["parent_hwnd"] != source["parent_hwnd"] or node["hwnd"] == source["hwnd"]:
                continue
            x1, y1, x2, y2 = node["rect"]
            match = re.fullmatch(r"\s*([\d,]+)(?:\s*\([^)]*%\))?\s*", node["text"])
            if match and abs((x1 + x2) - (left + right)) <= max(8, right - left) and bottom <= y1 <= bottom + max(30, (bottom - top) * 3):
                candidates.append((y1 - bottom, abs(x1 - left), int(match.group(1).replace(",", ""))))
        candidates.sort()
        if not candidates or (len(candidates) > 1 and candidates[0][:2] == candidates[1][:2]):
            raise RuntimeError(f"CTIの「{label}」の数値を読み取れませんでした。")
        result[key] = candidates[0][2]
    return result


def counter_context(nodes):
    """リスト切替の前後で累計値を差し引かないための識別値。"""
    labels = sorted((n for n in nodes if n["visible"] and n["text"] == "リスト" and "STATIC" in n["class_name"]),
                    key=lambda n: n["rect"][1])
    if not labels:
        return ""
    label = labels[0]
    rect = label["rect"]
    combos = [n for n in nodes if "COMBOBOX" in n["class_name"] and n["parent_hwnd"] == label["parent_hwnd"]
              and n["rect"][0] >= rect[2] and abs(n["rect"][1] - rect[1]) <= 8]
    if not combos:
        return ""
    value = min(combos, key=lambda n: n["rect"][0])["text"]
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:16]


def validate_staff_selection(table, staff_name):
    own = [row for row in table["rows"] if normalize_name(row.get("スタッフ")) == normalize_name(staff_name)]
    selected = [row for row in table["rows"] if row.get("選択", "").strip().lower() == "true"]
    if len(own) != 1 or len(selected) != 1 or selected[0] is not own[0]:
        raise RuntimeError("CTIの架電履歴で、ログイン中の本人だけをチェックしてください。チェックの自動切替は行いません。")


def result_count(nodes, history_hwnd):
    """結果グリッドに最も近い祖先パネル内の『N 件』表示を優先する。"""
    indexed = {n["hwnd"]: n for n in nodes}
    current = history_hwnd
    for _ in range(8):
        if current not in indexed:
            break
        family = descendants(nodes, current)
        counts = [int(m.group(1).replace(",", "")) for n in family
                  if "STATIC" in n["class_name"] and (m := re.fullmatch(r"\s*([\d,]+)\s*件\s*", n["text"]))]
        if len(counts) == 1:
            return counts[0]
        if len(counts) > 1:
            break
        current = indexed[current]["parent_hwnd"]
    raise RuntimeError("履歴の検索結果件数を確認できませんでした。")


def validate_search_dates(nodes, day):
    values = [n["text"].strip() for n in nodes if "SysDateTimePick32" in n["class_name"]]
    if len(values) != 2:
        raise RuntimeError("架電履歴の検索日時を確認できませんでした。")
    try:
        dates = sorted(datetime.fromisoformat(v.replace("/", "-")) for v in values)
    except ValueError as exc:
        raise RuntimeError("架電履歴の検索日時を読み取れませんでした。") from exc
    if dates[0].date().isoformat() != day or dates[1].date().isoformat() != day or dates[0].time().isoformat() != "00:00:00" or dates[1].time().isoformat() != "23:59:59":
        raise RuntimeError("架電履歴の日時を当日の00:00:00～23:59:59に設定してください。")


def validate_phone_filter(nodes):
    labels = [n for n in nodes if n["text"] == "電話番号" and "STATIC" in n["class_name"]]
    if len(labels) != 1:
        raise RuntimeError("架電履歴の電話番号検索欄を特定できませんでした。")
    label = labels[0]
    rect = label["rect"]
    editors = [n for n in nodes if "EDIT" in n["class_name"] and n["parent_hwnd"] == label["parent_hwnd"]
               and n["rect"][0] >= rect[2]
               and abs((n["rect"][1] + n["rect"][3]) - (rect[1] + rect[3])) <= max(10, rect[3] - rect[1])]
    if len(editors) != 1:
        raise RuntimeError("架電履歴の電話番号検索欄を特定できませんでした。")
    if editors[0]["text"].strip():
        raise RuntimeError("架電履歴の電話番号検索欄を空にしてください。")


class CtiStatisticsCollector:
    def __init__(self, cancel=None):
        self.cancel = cancel or threading.Event()

    def check_cancelled(self):
        if self.cancel.is_set():
            raise RuntimeError("集計を中止しました。")

    def collect(self, numeric_unit="seconds"):
        started = datetime.now(JST)
        day = started.date().isoformat()
        windows = Windows()
        ctis = windows.windows("CTIメイン")
        if len(ctis) != 1:
            raise RuntimeError("CTIメインを1つだけ起動してください。")
        cti = ctis[0]
        before = windows.snapshot(cti["hwnd"], 3000)
        if before["errors"]:
            raise RuntimeError("CTIの画面構造をすべて取得できませんでした。")
        # Read the footer directly so a large active grid cannot exhaust the scan.
        uia = uia_snapshot(cti["hwnd"], SimpleNamespace(max_nodes=100, max_depth=5, control_type="StatusBar"))
        self.check_cancelled()
        staff_name, staff_id = read_staff(uia.get("nodes", []))
        pages = [n for n in before["nodes"] if n["text"] == "架電履歴" and n["class_name"].startswith("WindowsForms10.Window.")]
        if len(pages) != 1:
            raise RuntimeError("CTIの架電履歴ページを特定できませんでした。")
        controls = descendants(before["nodes"], pages[0]["hwnd"])
        validate_search_dates(controls, day)
        validate_phone_filter(controls)
        tables = msaa_tables(read_msaa([n["hwnd"] for n in controls if n["class_name"].startswith("WindowsForms10.Window.")]))
        self.check_cancelled()
        staff_tables = [t for t in tables if {"選択", "スタッフ"}.issubset(t["headers"])]
        if len(staff_tables) != 1:
            raise RuntimeError("スタッフ一覧を特定できませんでした。")
        validate_staff_selection(staff_tables[0], staff_name)
        buttons = [n for n in controls if n["text"] == "検索" and "BUTTON" in n["class_name"]]
        if len(buttons) != 1:
            raise RuntimeError("CTIの履歴検索ボタンを特定できませんでした。")
        self.check_cancelled()
        result = ctypes.c_size_t()
        button = buttons[0]
        ok = windows.api.SendMessageTimeoutW(button["parent_hwnd"], 0x111, button["control_id"] & 0xffff,
                                             button["hwnd"], 2, 15000, ctypes.byref(result))
        if not ok:
            raise RuntimeError("CTIの検索応答がありません。検索を連続実行せず、完了を確認してください。")
        self.check_cancelled()
        refreshed = windows.snapshot(cti["hwnd"], 3000)
        refreshed_controls = descendants(refreshed["nodes"], pages[0]["hwnd"])
        table_handles = [n["hwnd"] for n in refreshed_controls if n["class_name"].startswith("WindowsForms10.Window.")]
        # Read twice after the command returns; never treat a truncated table as all rows.
        stable = None
        for _ in range(5):
            if self.cancel.wait(0.3):
                self.check_cancelled()
            snapshots = msaa_tables(read_msaa(table_handles))
            self.check_cancelled()
            histories = [t for t in snapshots if {"対応日時", "IP電話通話時間", "管理番号"}.issubset(t["headers"])]
            if not histories:
                # Clearing a search can remove the grid schema until the next search.
                histories = [t for t in snapshots if t["hwnd"] != staff_tables[0]["hwnd"] and not t["headers"] and not t["rows"]]
            history = histories[0] if len(histories) == 1 else None
            staff = next((t for t in snapshots if t["hwnd"] == staff_tables[0]["hwnd"]), None)
            if not history or not staff:
                raise RuntimeError("検索後の履歴を取得できませんでした。")
            validate_staff_selection(staff, staff_name)
            # CTI leaves the search button disabled after a completed search.
            # Use a stable table and the result-count check below instead.
            if stable == history["rows"]:
                break
            stable = history["rows"]
        else:
            raise RuntimeError("履歴の更新が完了しませんでした。")
        after = windows.snapshot(cti["hwnd"], 3000)
        if after["errors"]:
            raise RuntimeError("検索後の画面を取得できませんでした。")
        if datetime.now(JST).date().isoformat() != day:
            raise RuntimeError("集計中に日付が変わりました。再実行してください。")
        expected = result_count(after["nodes"], history["hwnd"])
        if expected != len(history["rows"]):
            raise RuntimeError(f"履歴は{expected}件ですが{len(history['rows'])}件しか読めませんでした。ページ分割や検索条件を確認してください。")
        # Filters can be changed by the operator while reading.
        after_controls = descendants(after["nodes"], pages[0]["hwnd"])
        validate_search_dates(after_controls, day)
        validate_phone_filter(after_controls)
        history_totals = hourly_history_totals(history["rows"], day, staff_name, numeric_unit)
        counters = read_counters(after["nodes"])
        context = counter_context(after["nodes"])
        if context != counter_context(before["nodes"]):
            raise RuntimeError("集計中にCTIのリストが切り替わりました。再実行してください。")
        self.check_cancelled()
        return StatisticsSnapshot(datetime.now(JST), staff_id, staff_name, f"{cti['pid']}:{cti['hwnd']}:{context}",
                                  counters, history_totals["talk_seconds"], len(history["rows"]),
                                  history_totals["counters"], history_totals["first_hour"])
