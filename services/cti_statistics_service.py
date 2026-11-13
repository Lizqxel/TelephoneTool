"""CTIを表示切替せず検索し、集計に必要な値だけを返す。"""

import ctypes
from datetime import datetime
import hashlib
import re
import threading
from types import SimpleNamespace

from services.call_statistics import JST, StatisticsSnapshot, hourly_history_totals, may_be_new_prospect, normalize_name
from services.cti_accessibility import Windows, msaa_tables, read_msaa, read_status_bar, uia_snapshot


WM_COMMAND = 0x111
SMTO_ABORTIFHUNG = 0x0002
MAX_HISTORY_PAGES = 60
COUNTER_LABELS = {"総コール数": "total_calls", "新規見込": "new_prospects", "対話有NG": "ng_with_dialogue",
                  "対話無NG": "ng_without_dialogue", "有効コール": "valid_calls"}


class CollectionSkipped(RuntimeError):
    """CTI側の検索条件が今は集計に使えない。保存済みの値を維持し、次回に取り直す。"""


class CollectionCancelled(RuntimeError):
    pass


def descendants(nodes, root):
    ids = {root}
    for node in nodes:
        if node["parent_hwnd"] in ids:
            ids.add(node["hwnd"])
    return [n for n in nodes if n["hwnd"] in ids]


def parse_staff(names):
    """『苗字 名前 さん：ログインID』からスタッフ名とIDを取り出す。"""
    matches = []
    for name in names:
        match = re.fullmatch(r"(.+?)\s*さん\s*[:：]\s*(\S+)", (name or "").strip())
        if match:
            matches.append((match.group(1).strip(), match.group(2)))
    if len(matches) != 1:
        raise RuntimeError("CTI下部のログインスタッフ名を特定できませんでした。")
    return matches[0]


def read_staff(nodes):
    parents = {n["id"] for n in nodes if n["control_type"] == "ControlType.StatusBar"}
    return parse_staff(node["name"] for node in nodes if node["parent_id"] in parents)


def status_bar_handles(nodes, root):
    """WinFormsのステータスバー（statusStrip1）を優先し、無ければメイン画面直下の部品を候補にする。"""
    named = [n["hwnd"] for n in nodes if n["text"].lower().startswith("statusstrip")]
    return named or [n["hwnd"] for n in nodes if n["parent_hwnd"] == root]


def read_counters(nodes):
    """上部の累計カウンター。値の欄が空白のものは None（未取得）にする。"""
    result = {}
    for label, key in COUNTER_LABELS.items():
        labels = [n for n in nodes if n["text"].strip() == label and "STATIC" in n["class_name"]]
        if len(labels) != 1:
            raise RuntimeError(f"CTIの「{label}」を特定できませんでした。")
        source = labels[0]
        left, top, right, bottom = source["rect"]
        candidates, blank = [], False
        for node in nodes:
            if node["parent_hwnd"] != source["parent_hwnd"] or node["hwnd"] == source["hwnd"]:
                continue
            x1, y1, x2, y2 = node["rect"]
            if not (abs((x1 + x2) - (left + right)) <= max(8, right - left) and bottom <= y1 <= bottom + max(30, (bottom - top) * 3)):
                continue
            match = re.fullmatch(r"\s*([\d,]+)(?:\s*\([^)]*%\))?\s*", node["text"])
            if match:
                candidates.append((y1 - bottom, abs(x1 - left), int(match.group(1).replace(",", ""))))
            elif not node["text"].strip() and "STATIC" in node["class_name"]:
                blank = True
        candidates.sort()
        if not candidates and blank:
            result[key] = None
            continue
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


def validate_staff_selection(table, staff_name, page="架電履歴"):
    """本人にチェックが入っていればよい。他のスタッフのチェックは結果の行を名前で絞り込んで除く。"""
    own = [row for row in table["rows"] if normalize_name(row.get("スタッフ")) == normalize_name(staff_name)]
    if len(own) != 1:
        raise RuntimeError(f"CTIの{page}のスタッフ一覧で、ログイン中の本人を特定できませんでした。")
    if own[0].get("選択", "").strip().lower() != "true":
        raise CollectionSkipped(f"CTIの{page}のスタッフ一覧で本人にチェックが入っていないため、今回はスキップしました。")


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


def history_total(nodes):
    """架電履歴の『検索結果件数』（全ページの合計）。表は50件ごとのページに分かれる。"""
    labels = [n for n in nodes if n["text"].strip() == "検索結果件数" and "STATIC" in n["class_name"]]
    if len(labels) != 1:
        raise RuntimeError("架電履歴の検索結果件数を確認できませんでした。")
    label = labels[0]
    left, top, right, bottom = label["rect"]
    values = [int(n["text"].replace(",", "")) for n in nodes
              if n["hwnd"] != label["hwnd"] and n["parent_hwnd"] == label["parent_hwnd"] and "STATIC" in n["class_name"]
              and re.fullmatch(r"\s*[\d,]+\s*", n["text"])
              and top <= (n["rect"][1] + n["rect"][3]) / 2 <= bottom and left <= n["rect"][0] < right]
    if len(values) != 1:
        raise RuntimeError("架電履歴の検索結果件数を読み取れませんでした。")
    return values[0]


def validate_search_dates(nodes, day):
    """検索期間が当日の00:00:00～23:59:59を含んでいればよい（他の日の行は集計で除く）。"""
    values = [n["text"].strip() for n in nodes if "SysDateTimePick32" in n["class_name"]]
    if len(values) != 2:
        raise RuntimeError("架電履歴の検索日時を確認できませんでした。")
    try:
        start, end = sorted(datetime.fromisoformat(v.replace("/", "-")) for v in values)
    except ValueError as exc:
        raise RuntimeError("架電履歴の検索日時を読み取れませんでした。") from exc
    if start > datetime.fromisoformat(f"{day}T00:00:00") or end < datetime.fromisoformat(f"{day}T23:59:59"):
        raise CollectionSkipped("架電履歴の検索日時が当日（00:00:00～23:59:59）を含んでいないため、今回は履歴の取得をスキップしました。")


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
        raise CollectionSkipped("架電履歴の電話番号欄に入力があるため、今回は履歴の取得をスキップしました。空にすると次回から取得します。")


def find_button(controls, text):
    buttons = [n for n in controls if n["text"] == text and "BUTTON" in n["class_name"]]
    if len(buttons) != 1:
        raise RuntimeError(f"CTIの「{text}」ボタンを特定できませんでした。")
    return buttons[0]


def find_page(nodes, title):
    pages = [n for n in nodes if n["text"] == title and n["class_name"].startswith("WindowsForms10.Window.")]
    if len(pages) > 1:
        raise RuntimeError(f"CTIの{title}ページを特定できませんでした。")
    return pages[0] if pages else None


def grid_handles(nodes, page):
    return [n["hwnd"] for n in descendants(nodes, page["hwnd"]) if n["class_name"].startswith("WindowsForms10.Window.")]


def choose_history(tables, staff_hwnd):
    histories = [t for t in tables if {"対応日時", "IP電話通話時間", "管理番号"}.issubset(t["headers"])]
    if not histories:
        # Clearing a search can remove the grid schema until the next search.
        histories = [t for t in tables if t["hwnd"] != staff_hwnd and not t["headers"] and not t["rows"]]
    return histories[0] if len(histories) == 1 else None


def choose_prospects(tables, staff_hwnd):
    grids = [t for t in tables if {"管理番号", "ランク"}.issubset(t["headers"])]
    if not grids:
        grids = [t for t in tables if t["hwnd"] != staff_hwnd and not t["headers"] and not t["rows"]]
    return grids[0] if len(grids) == 1 else None


def staff_table(tables):
    found = [t for t in tables if {"選択", "スタッフ"}.issubset(t["headers"]) and "管理番号" not in t["headers"]]
    if len(found) != 1:
        raise RuntimeError("スタッフ一覧を特定できませんでした。")
    return found[0]


class CtiStatisticsCollector:
    def __init__(self, cancel=None):
        self.cancel = cancel or threading.Event()
        self.windows = None

    def check_cancelled(self):
        if self.cancel.is_set():
            raise CollectionCancelled("集計を中止しました。")

    def press(self, button, timeout_ms=15000):
        """非表示のボタンへ BN_CLICKED を送る。画面の切替やマウス操作はしない。"""
        self.check_cancelled()
        result = ctypes.c_size_t()
        return bool(self.windows.api.SendMessageTimeoutW(button["parent_hwnd"], WM_COMMAND, button["control_id"] & 0xffff,
                                                         button["hwnd"], SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(result)))

    def read_grid(self, handles, choose, staff_hwnd, previous=None):
        """2回続けて同じ内容になるまで読む。previous があれば、それと違う表（次のページ）を待つ。"""
        stable = None
        for _ in range(10):
            if self.cancel.wait(0.3):
                self.check_cancelled()
            tables = msaa_tables(read_msaa(handles))
            self.check_cancelled()
            grid = choose(tables, staff_hwnd)
            if grid is None:
                raise RuntimeError("検索後の表を取得できませんでした。")
            if previous is not None and grid["rows"] == previous:
                continue
            if stable is not None and stable == grid["rows"]:
                return grid, tables
            stable = grid["rows"]
        raise RuntimeError("表の更新が完了しませんでした。")

    def collect(self, numeric_unit="seconds", full=True):
        """上部カウンターは毎回、full のときは架電履歴と見込検索も取得する。

        履歴を取れない回（条件が合わない・エラー）もカウンターは返し、理由を note に入れる。
        """
        started = datetime.now(JST)
        day = started.date().isoformat()
        self.windows = Windows()
        ctis = self.windows.windows("CTIメイン")
        if len(ctis) != 1:
            raise RuntimeError("CTIメインを1つだけ起動してください。")
        cti = ctis[0]
        before = self.windows.snapshot(cti["hwnd"], 5000)
        observed = datetime.now(JST)
        if before["errors"]:
            raise RuntimeError("CTIの画面構造をすべて取得できませんでした。")
        counters = read_counters(before["nodes"])
        session = f"{cti['pid']}:{cti['hwnd']}:{counter_context(before['nodes'])}"
        staff_name, staff_id = self.read_login_staff(cti, before["nodes"])
        base = dict(acquired_at=observed, staff_id=staff_id, staff_name=staff_name, session=session,
                    counters=counters, talk_seconds=None, history_rows=0)
        if not full:
            return StatisticsSnapshot(**base)
        try:
            rows = self.collect_history(cti, before, day, staff_name)
            ranks, rank_note = None, ""
            if any(normalize_name(r.get("対応スタッフ")) == normalize_name(staff_name)
                   and r.get("対応日時", "").replace("/", "-").startswith(day) and may_be_new_prospect(r) for r in rows):
                ranks, rank_note = self.collect_ranks(cti, before, staff_name)
            else:
                ranks = {}
            totals = hourly_history_totals(rows, day, staff_name, numeric_unit, ranks)
        except CollectionCancelled:
            raise
        except CollectionSkipped as exc:
            return StatisticsSnapshot(**base, note=str(exc))
        except (RuntimeError, ValueError) as exc:
            return StatisticsSnapshot(**base, note=f"履歴を取得できませんでした：{exc}")
        if totals["unmatched_prospects"]:
            rank_note = (rank_note + " " if rank_note else "") + f"見込検索に無い見込{totals['unmatched_prospects']}件はトスに数えていません。"
        if totals["unknown_dialogue_results"]:
            listed = "、".join(f"{k or '空欄'} {v}件" for k, v in totals["unknown_dialogue_results"].items())
            rank_note = (rank_note + " " if rank_note else "") + f"対話の有無が分からない架電結果があるため、その時間帯の対話は未判定です（{listed}）。"
        own_rows = sum(totals["counters"][h]["total_calls"] for h in range(24))
        return StatisticsSnapshot(**{**base, "talk_seconds": totals["talk_seconds"], "history_rows": own_rows},
                                  history_counters=totals["counters"], first_history_hour=totals["first_hour"], note=rank_note)

    def read_login_staff(self, cti, nodes):
        """ステータスバーのログイン名。MSAAは最小化中も読めるので優先し、読めなければUI Automationで読む。"""
        try:
            return parse_staff(read_status_bar(status_bar_handles(nodes, cti["hwnd"])))
        except RuntimeError:
            pass
        self.check_cancelled()
        # Read the footer directly so a large active grid cannot exhaust the scan.
        uia = uia_snapshot(cti["hwnd"], SimpleNamespace(max_nodes=100, max_depth=5, control_type="StatusBar"))
        self.check_cancelled()
        if not uia.get("nodes") and self.windows.api.IsIconic(cti["hwnd"]):
            raise RuntimeError("CTIが最小化されているため、ログインスタッフを確認できません。CTIを元のサイズに戻すと次回から取得します。")
        return read_staff(uia.get("nodes", []))

    def collect_history(self, cti, before, day, staff_name):
        page = find_page(before["nodes"], "架電履歴")
        if page is None:
            raise CollectionSkipped("CTIの架電履歴タブがまだ読み込まれていません。CTI起動後に一度、架電履歴タブを開いてください。")
        if page["visible"]:
            raise CollectionSkipped("架電履歴を表示中のため、今回は履歴の取得をスキップしました。")
        controls = descendants(before["nodes"], page["hwnd"])
        validate_search_dates(controls, day)
        validate_phone_filter(controls)
        tables = msaa_tables(read_msaa(grid_handles(before["nodes"], page)))
        self.check_cancelled()
        staff = staff_table(tables)
        validate_staff_selection(staff, staff_name)
        if not self.press(find_button(controls, "検索")):
            raise RuntimeError("CTIの検索応答がありません。検索を連続実行せず、完了を確認してください。")
        refreshed = self.windows.snapshot(cti["hwnd"], 5000)
        refreshed_page = find_page(refreshed["nodes"], "架電履歴") or page
        handles = grid_handles(refreshed["nodes"], refreshed_page)
        # Read twice after the command returns; never treat a truncated table as all rows.
        grid, tables = self.read_grid(handles, choose_history, staff["hwnd"])
        after = self.windows.snapshot(cti["hwnd"], 5000)
        if after["errors"]:
            raise RuntimeError("検索後の画面を取得できませんでした。")
        after_controls = descendants(after["nodes"], refreshed_page["hwnd"])
        total = history_total(after_controls)
        rows = list(grid["rows"])
        if len(rows) < total:
            next_button = find_button(after_controls, "次へ")
            for _ in range(MAX_HISTORY_PAGES):
                if len(rows) >= total:
                    break
                if not self.press(next_button):
                    raise RuntimeError("架電履歴の次のページへ進めませんでした。")
                grid, tables = self.read_grid(handles, choose_history, staff["hwnd"], previous=grid["rows"])
                if not grid["rows"]:
                    break
                rows.extend(grid["rows"])
            # 担当者が次に開いたときのため、1ページ目へ戻す（失敗しても集計には影響しない）。
            try:
                self.press(find_button(after_controls, "最初"), 5000)
            except RuntimeError:
                pass
        if total != len(rows):
            raise RuntimeError(f"履歴は{total}件ですが{len(rows)}件しか読めませんでした。")
        staff_after = next((t for t in tables if t["hwnd"] == staff["hwnd"]), None)
        if staff_after is not None:
            validate_staff_selection(staff_after, staff_name)
        if datetime.now(JST).date().isoformat() != day:
            raise RuntimeError("集計中に日付が変わりました。再実行してください。")
        # Filters can be changed by the operator while reading.
        final = self.windows.snapshot(cti["hwnd"], 5000)
        final_controls = descendants(final["nodes"], refreshed_page["hwnd"])
        validate_search_dates(final_controls, day)
        validate_phone_filter(final_controls)
        return rows

    def collect_ranks(self, cti, before, staff_name):
        """見込検索から本人の {管理番号: ランク}。取れなければ (None, 理由) でトスを未取得にする。"""
        try:
            page = find_page(before["nodes"], "見込検索")
            if page is None:
                return None, "見込検索タブがまだ読み込まれていないため、トスは未取得です。CTI起動後に一度、見込検索タブを開いてください。"
            controls = descendants(before["nodes"], page["hwnd"])
            tables = msaa_tables(read_msaa(grid_handles(before["nodes"], page)))
            self.check_cancelled()
            staff = staff_table(tables)
            validate_staff_selection(staff, staff_name, "見込検索")
            if page["visible"]:
                # 担当者が見ている画面は検索し直さず、表示中の結果を使う。
                grid, nodes = choose_prospects(tables, staff["hwnd"]), before["nodes"]
            else:
                if not self.press(find_button(controls, "検索")):
                    return None, "見込検索の検索応答がないため、トスは未取得です。"
                refreshed = self.windows.snapshot(cti["hwnd"], 5000)
                refreshed_page = find_page(refreshed["nodes"], "見込検索") or page
                grid, _ = self.read_grid(grid_handles(refreshed["nodes"], refreshed_page), choose_prospects, staff["hwnd"])
                nodes = self.windows.snapshot(cti["hwnd"], 5000)["nodes"]
            if grid is None:
                return None, "見込検索の結果を特定できないため、トスは未取得です。"
            if grid["rows"]:
                expected = result_count(nodes, grid["hwnd"])
                if expected != len(grid["rows"]):
                    return None, f"見込検索は{expected}件ですが{len(grid['rows'])}件しか読めないため、トスは未取得です。"
            return {str(r.get("管理番号", "")).strip(): r.get("ランク", "") for r in grid["rows"]
                    if normalize_name(r.get("スタッフ")) == normalize_name(staff_name)}, ""
        except CollectionCancelled:
            raise
        except (RuntimeError, ValueError) as exc:
            return None, f"トスは未取得です（{exc}）"
