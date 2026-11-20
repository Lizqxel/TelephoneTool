"""CTIを表示切替せず検索し、集計に必要な値だけを返す。"""

import ctypes
from datetime import datetime
import hashlib
import re
import threading
from types import SimpleNamespace

from services.call_statistics import JST, StatisticsSnapshot, hourly_history_totals, may_be_new_prospect, normalize_name
from services.cti_accessibility import Windows, msaa_tables, read_msaa, read_status_bar, uia_snapshot
from services.cti_remote import set_date_time


WM_COMMAND = 0x111
WM_GETTEXT = 0x000D
SMTO_ABORTIFHUNG = 0x0002
CB_GETCOUNT = 0x0146
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149
CB_SETCURSEL = 0x014E
CBN_SELCHANGE = 1
# 架電履歴の検索条件の「未選択」。CTIは商材が未選択だと架電結果詳細（現アナNGなど）を空欄で返す。
UNSELECTED_CHOICE = "-"
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


def parse_search_date(text):
    return datetime.fromisoformat(text.strip().replace("/", "-"))


def search_date_pickers(nodes):
    """架電履歴の検索日時の欄（開始・終了）。CTIでは上が開始、下が終了。2つ無ければ None。"""
    pickers = sorted((n for n in nodes if "SysDateTimePick32" in n["class_name"]), key=lambda n: (n["rect"][1], n["rect"][0]))
    return pickers if len(pickers) == 2 else None


def validate_search_dates(nodes, day):
    """検索期間が対象日の00:00:00～23:59:59を含んでいればよい（他の日の行は集計で除く）。"""
    values = [n["text"].strip() for n in nodes if "SysDateTimePick32" in n["class_name"]]
    if len(values) != 2:
        raise RuntimeError("架電履歴の検索日時を確認できませんでした。")
    try:
        start, end = sorted(parse_search_date(v) for v in values)
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


def find_labeled_combo(nodes, label_text):
    """ラベルと同じ行で、すぐ右にあるコンボボックス。見つからなければ None。"""
    labels = [n for n in nodes if n["text"] == label_text and "STATIC" in n["class_name"]]
    if len(labels) != 1:
        return None
    label = labels[0]
    rect = label["rect"]
    combos = [n for n in nodes if "COMBOBOX" in n["class_name"] and n["parent_hwnd"] == label["parent_hwnd"]
              and n["rect"][0] >= rect[2] - 4
              and abs((n["rect"][1] + n["rect"][3]) - (rect[1] + rect[3])) <= max(10, rect[3] - rect[1])]
    return min(combos, key=lambda n: n["rect"][0]) if combos else None


def current_list_product(nodes):
    """架電画面で、今のリストと同じ行にある「商材」欄の値。読めなければ空文字。"""
    lists = sorted((n for n in nodes if n["visible"] and n["text"] == "リスト" and "STATIC" in n["class_name"]),
                   key=lambda n: n["rect"][1])
    if not lists:
        return ""
    top = lists[0]["rect"][1]
    labels = [n for n in nodes if n["visible"] and n["text"].strip() == "商材" and "STATIC" in n["class_name"]
              and abs(n["rect"][1] - top) <= 8]
    if len(labels) != 1:
        return ""
    label = labels[0]
    rect = label["rect"]
    values = [n for n in nodes if n["visible"] and n["hwnd"] != label["hwnd"] and n["parent_hwnd"] == label["parent_hwnd"]
              and any(kind in n["class_name"] for kind in ("EDIT", "STATIC", "COMBOBOX")) and n["text"].strip()
              and n["rect"][0] >= rect[2] - 4
              and abs((n["rect"][1] + n["rect"][3]) - (rect[1] + rect[3])) <= max(10, rect[3] - rect[1])]
    return min(values, key=lambda n: n["rect"][0])["text"].strip() if values else ""


def product_to_select(items, current, preferred=""):
    """架電履歴の商材が未選択なら、選ぶ位置を返す。選ばないときは (None, 理由)。

    架電画面の商材（preferred）と同じ名前を優先し、無ければ選べる商材が1つだけのときにそれを選ぶ。
    """
    if current is None or not 0 <= current < len(items):
        return None, "架電履歴の商材の選択状態を読めないため、商材を自動で選んでいません。"
    if normalize_name(items[current]) != UNSELECTED_CHOICE:
        return None, ""  # 担当者が選んだ商材のまま検索する
    if preferred:
        same = [i for i, item in enumerate(items) if normalize_name(item) == normalize_name(preferred)]
        if len(same) == 1:
            return same[0], ""
    products = [i for i, item in enumerate(items) if normalize_name(item) not in ("", UNSELECTED_CHOICE)]
    if len(products) == 1:
        return products[0], ""
    if preferred:
        return None, f"架電履歴の商材に、架電画面の商材「{preferred}」と同じ選択肢が無いため、自動で選んでいません。"
    return None, f"架電画面の商材を読めず、架電履歴で選べる商材が{len(products)}件あるため、自動で選んでいません。"


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


def search_conditions(nodes):
    return {n["hwnd"]: n["text"] for n in nodes
            if any(kind in n["class_name"] for kind in ("EDIT", "COMBOBOX", "LISTBOX", "SysDateTimePick32"))}


class CtiStatisticsCollector:
    def __init__(self, cancel=None):
        self.cancel = cancel or threading.Event()
        self.windows = None
        self._history_context = None
        self.clear_history_info = None
        self.history_notes = []  # 集計結果と一緒に表示するお知らせ
        self.product_note = ""  # 商材を自動で選べなかった理由（NGの詳細が読めなかったときに表示）
        self.started_day = None  # 集計を始めた日（途中で日付が変わったら集計し直す）

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

    def collect(self, numeric_unit="seconds", day=None):
        """上部カウンターと、架電履歴・見込検索から当日分を取得する。

        day に過去日を渡すと、その日の架電履歴だけを集計する（上部カウンターは当日の累計なので使わない）。
        履歴を取れない回（条件が合わない・エラー）もカウンターは返し、理由を note に入れる。
        """
        self._history_context = None
        self.clear_history_info = None
        self.history_notes = []
        self.product_note = ""
        started = datetime.now(JST)
        self.started_day = started.date().isoformat()
        day = day or self.started_day
        past = day != self.started_day
        if day > self.started_day:
            raise RuntimeError("先の日付は集計できません。")
        self.windows = Windows()
        ctis = self.windows.windows("CTIメイン")
        if len(ctis) != 1:
            raise RuntimeError("CTIメインを1つだけ起動してください。")
        cti = ctis[0]
        before = self.windows.snapshot(cti["hwnd"], 5000)
        # 過去日は、その日の終わりまでの履歴として保存・送信する。
        observed = datetime.fromisoformat(f"{day}T23:59:59").replace(tzinfo=JST) if past else datetime.now(JST)
        if before["errors"]:
            raise RuntimeError("CTIの画面構造をすべて取得できませんでした。")
        counters = {key: None for key in COUNTER_LABELS.values()} if past else read_counters(before["nodes"])
        session = f"{cti['pid']}:{cti['hwnd']}:{counter_context(before['nodes'])}"
        staff_name, staff_id = self.read_login_staff(cti, before["nodes"])
        base = dict(acquired_at=observed, staff_id=staff_id, staff_name=staff_name, session=session,
                    counters=counters, talk_seconds=None, history_rows=0, history_only=past)
        try:
            rows = self.collect_history(cti, before, day, staff_name)
            ranks, rank_note = None, ""
            if any(normalize_name(r.get("対応スタッフ")) == normalize_name(staff_name)
                   and r.get("対応日時", "").replace("/", "-").startswith(day) and may_be_new_prospect(r) for r in rows):
                ranks, rank_note = self.collect_ranks(cti, before, staff_name)
            else:
                ranks = {}
            totals = hourly_history_totals(rows, day, staff_name, numeric_unit, ranks)
            halves = hourly_history_totals(rows, day, staff_name, numeric_unit, ranks, interval_minutes=30)
        except CollectionCancelled:
            raise
        except CollectionSkipped as exc:
            self._history_context = None  # 集計していない一覧はクリアしない
            return StatisticsSnapshot(**base, note=str(exc))
        except (RuntimeError, ValueError) as exc:
            self._history_context = None
            return StatisticsSnapshot(**base, note=f"履歴を取得できませんでした：{exc}")
        if totals["unmatched_prospects"]:
            rank_note = (rank_note + " " if rank_note else "") + f"見込検索に無い見込{totals['unmatched_prospects']}件はトスに数えていません。"
        notes = [rank_note] if rank_note else []
        if totals["unknown_dialogue_results"]:
            listed = "、".join(f"{k or '空欄'} {v}件" for k, v in totals["unknown_dialogue_results"].items())
            notes.append(f"対話の有無が分からない架電結果があるため、その時間帯の対話は未判定です（{listed}）。")
            if any(k.endswith("（詳細なし）") for k in totals["unknown_dialogue_results"]):
                notes.append(self.product_note or "CTIが架電結果詳細を返さなかったため、NGが対話有か対話無か判定できません。")
        rank_note = " ".join(notes + self.history_notes)
        own_rows = sum(totals["counters"][h]["total_calls"] for h in range(24))
        return StatisticsSnapshot(**{**base, "talk_seconds": totals["talk_seconds"], "history_rows": own_rows},
                                  history_counters=totals["counters"], first_history_hour=totals["first_hour"], note=rank_note,
                                  intervals={minute: {**counts, "dialogue_seconds": halves["dialogue_seconds"][minute]}
                                             for minute, counts in halves["counters"].items()})

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
        # 検索日時はその日の00:00:00〜23:59:59にして検索し、読み終えたら元に戻す（過去日の集計と、長い期間の読み込み短縮）。
        dates = self.prepare_dates(cti, controls, day)
        try:
            if dates is None:
                validate_search_dates(controls, day)
            rows, first_page, grid, tables, refreshed_page, total, staff = self.search_own_history(cti, before, page, controls, staff_name)
        except BaseException:
            if dates:
                try:
                    self.restore_dates(cti, dates)
                except Exception:
                    pass  # 集計できなかった理由（元の例外）を優先して知らせる
            raise
        if dates:
            try:
                self.restore_dates(cti, dates)
            except RuntimeError as exc:
                self.history_notes.append(f"架電履歴の検索日時を元に戻せませんでした（{exc}）。")
        if total != len(rows):
            raise RuntimeError(f"履歴は{total}件ですが{len(rows)}件しか読めませんでした。")
        staff_after = next((t for t in tables if t["hwnd"] == staff["hwnd"]), None)
        if staff_after is not None:
            validate_staff_selection(staff_after, staff_name)
        if datetime.now(JST).date().isoformat() != (self.started_day or day):
            raise RuntimeError("集計中に日付が変わりました。再実行してください。")
        # Filters can be changed by the operator while reading.
        final = self.windows.snapshot(cti["hwnd"], 5000)
        final_controls = descendants(final["nodes"], refreshed_page["hwnd"])
        if dates is None:
            validate_search_dates(final_controls, day)
        validate_phone_filter(final_controls)
        # 保存後のクリア用。1ページ目へ戻した状態の一覧と検索条件が変わっていなければクリアする。
        self._history_context = {"cti": cti, "page_hwnd": refreshed_page["hwnd"], "history_hwnd": grid["hwnd"],
                                 "staff_hwnd": staff["hwnd"], "staff_rows": (staff_after or staff)["rows"],
                                 "rows": first_page, "conditions": search_conditions(final_controls)}
        return rows

    def search_own_history(self, cti, before, page, controls, staff_name):
        """条件を確かめ、商材を選んで検索し、商材を元に戻す。検索日時は呼び出し元が戻す。"""
        validate_phone_filter(controls)
        tables = msaa_tables(read_msaa(grid_handles(before["nodes"], page)))
        self.check_cancelled()
        staff = staff_table(tables)
        validate_staff_selection(staff, staff_name)
        product = self.prepare_product(controls, current_list_product(before["nodes"]))
        try:
            if product:
                self.select_choice(product["combo"], product["target"])
            result = self.search_history(cti, page, controls, staff)
        except BaseException:
            if product:
                try:
                    self.restore_choice(product)
                except Exception:
                    pass  # 集計できなかった理由（元の例外）を優先して知らせる
            raise
        if product:
            try:
                self.restore_choice(product)
            except RuntimeError as exc:
                self.history_notes.append(f"架電履歴の商材を元の「{product['items'][product['original']]}」に戻せませんでした（{exc}）。")
        return (*result, staff)

    def read_text(self, hwnd):
        buffer = ctypes.create_unicode_buffer(256)
        if self.send_message(hwnd, WM_GETTEXT, len(buffer), ctypes.addressof(buffer)) is None:
            return None
        return buffer.value

    def prepare_dates(self, cti, controls, day):
        """検索日時をその日の00:00:00〜23:59:59にする。既にそうなら何もしない（None）。

        変えられなかったときは元に戻してお知らせを残し、今の検索日時のまま集計する（None）。
        """
        pickers = search_date_pickers(controls)
        if pickers is None:
            return None  # 検索日時の確認（validate_search_dates）で理由を出す
        try:
            original = [parse_search_date(p["text"]) for p in pickers]
        except ValueError:
            return None
        target = [datetime.fromisoformat(f"{day}T00:00:00"), datetime.fromisoformat(f"{day}T23:59:59")]
        if original == target:
            return None
        self.check_cancelled()
        dates = {"pickers": pickers, "original": original}
        try:
            self.set_dates(cti, pickers, target)
        except RuntimeError as exc:
            try:
                self.restore_dates(cti, dates)
            except RuntimeError:
                pass
            self.history_notes.append(f"架電履歴の検索日時を変えられなかったため、今の検索日時のまま集計しました（{exc}）。")
            return None
        return dates

    def set_dates(self, cti, pickers, values):
        for picker, value in zip(pickers, values):
            set_date_time(self.windows.api, cti["pid"], picker, value)
        shown = [self.read_text(p["hwnd"]) for p in pickers]
        try:
            if [parse_search_date(text or "") for text in shown] != list(values):
                raise ValueError()
        except ValueError:
            raise RuntimeError("架電履歴の検索日時が変わったことを確認できませんでした。") from None

    def restore_dates(self, cti, dates):
        """検索日時を集計前の値へ戻す。"""
        self.set_dates(cti, dates["pickers"], dates["original"])

    def search_history(self, cti, page, controls, staff):
        """架電履歴を裏で検索し、全ページを読んで1ページ目へ戻す。"""
        if not self.press(find_button(controls, "検索")):
            raise RuntimeError("CTIの検索応答がありません。検索を連続実行せず、完了を確認してください。")
        refreshed = self.windows.snapshot(cti["hwnd"], 5000)
        refreshed_page = find_page(refreshed["nodes"], "架電履歴") or page
        handles = grid_handles(refreshed["nodes"], refreshed_page)
        try:
            # Read twice after the command returns; never treat a truncated table as all rows.
            grid, tables = self.read_grid(handles, choose_history, staff["hwnd"])
            after = self.windows.snapshot(cti["hwnd"], 5000)
            if after["errors"]:
                raise RuntimeError("検索後の画面を取得できませんでした。")
            after_controls = descendants(after["nodes"], refreshed_page["hwnd"])
            total = history_total(after_controls)
        except RuntimeError:
            empty = self.empty_search_result(cti, refreshed_page, handles)
            if empty is None:
                raise
            return [], [], empty[0], empty[1], refreshed_page, 0
        rows = list(grid["rows"])
        first_page = list(grid["rows"])
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
        return rows, first_page, grid, tables, refreshed_page, total

    def empty_search_result(self, cti, page, handles):
        """検索結果が0件なら (空の表, 表の一覧)。そうでなければ None。

        CTIは0件のとき、表も件数表示も変えず、メッセージも出さずに終わる。結果があるときは
        検索ボタンを押せなくする（クリアするまで）ので、検索ボタンが押せて表が空なら0件と判断する。
        """
        snap = self.windows.snapshot(cti["hwnd"], 5000)
        if snap["errors"]:
            return None
        try:
            if not find_button(descendants(snap["nodes"], page["hwnd"]), "検索")["enabled"]:
                return None
        except RuntimeError:
            return None
        tables = msaa_tables(read_msaa(handles))
        # クリア直後は列の見出しが消えていることもある（choose_history と同じ扱い）。
        grids = [t for t in tables if {"対応日時", "管理番号"}.issubset(t["headers"]) or (not t["headers"] and not t["rows"])]
        if len(grids) != 1 or grids[0]["rows"]:
            return None
        return grids[0], tables

    def send_message(self, hwnd, message, wparam=0, lparam=0, timeout_ms=5000):
        """応答が無ければ None。CB_ERR（-1）などの負の戻り値は負の数で返す。"""
        result = ctypes.c_size_t()
        if not self.windows.api.SendMessageTimeoutW(hwnd, message, wparam, lparam, SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(result)):
            return None
        bits = ctypes.sizeof(ctypes.c_size_t) * 8
        return result.value - (1 << bits) if result.value >= 1 << (bits - 1) else result.value

    def combo_items(self, combo):
        count = self.send_message(combo["hwnd"], CB_GETCOUNT)
        if count is None or not 0 <= count <= 500:
            return None
        items = []
        for index in range(count):
            length = self.send_message(combo["hwnd"], CB_GETLBTEXTLEN, index)
            if length is None or length < 0:
                return None
            buffer = ctypes.create_unicode_buffer(length + 1)
            if self.send_message(combo["hwnd"], CB_GETLBTEXT, index, ctypes.addressof(buffer)) is None:
                return None
            items.append(buffer.value)
        return items

    def prepare_product(self, controls, preferred=""):
        """商材が未選択なら、検索の間だけ架電画面と同じ商材を選ぶ（選ばないと架電結果詳細が空欄になる）。

        選べないときは理由を残し、今の条件のまま検索する（詳細の無いNGは未判定になる）。
        """
        combo = find_labeled_combo(controls, "商材")
        if combo is None:
            self.product_note = "架電履歴の商材欄を特定できないため、商材を自動で選んでいません。"
            return None
        items = self.combo_items(combo)
        current = self.send_message(combo["hwnd"], CB_GETCURSEL)
        target, reason = product_to_select(items or [], current, preferred)
        self.product_note = reason
        if target is None:
            return None
        self.check_cancelled()
        return {"combo": combo, "items": items, "original": current, "target": target}

    def select_choice(self, combo, index):
        """利用者が選んだときと同じく、選択を変えてから親へ CBN_SELCHANGE を送る（CTIの検索条件が切り替わる）。"""
        if self.send_message(combo["hwnd"], CB_SETCURSEL, index) != index:
            raise RuntimeError("架電履歴の商材を選べませんでした。")
        wparam = (CBN_SELCHANGE << 16) | (combo["control_id"] & 0xffff)
        if self.send_message(combo["parent_hwnd"], WM_COMMAND, wparam, combo["hwnd"], 15000) is None:
            raise RuntimeError("架電履歴の商材の切り替えに応答がありません。")
        if self.send_message(combo["hwnd"], CB_GETCURSEL) != index:
            raise RuntimeError("架電履歴の商材が切り替わりませんでした。")

    def restore_choice(self, product):
        """自動で選んだ商材を、担当者が選んでいた状態（未選択）へ戻す。取消済みでも戻す。"""
        if self.send_message(product["combo"]["hwnd"], CB_GETCURSEL) != product["original"]:
            self.select_choice(product["combo"], product["original"])

    def clear_history(self):
        """呼び出し元が集計を保存した後で、取得した一覧だけをクリアする。"""
        self.check_cancelled()
        context = self._history_context
        if context is None:
            raise RuntimeError("クリア対象の集計済み履歴がありません。")
        if not context["rows"]:
            # 検索結果が0件のときはCTIの一覧は空のまま（クリアボタンも押せない）なので、何もしない。
            self._history_context = None
            self.clear_history_info = {"rows_before": 0, "rows_after": 0, "skipped": "検索結果が0件"}
            return
        windows = Windows()
        ctis = windows.windows("CTIメイン")
        if not any(w["hwnd"] == context["cti"]["hwnd"] and w["pid"] == context["cti"]["pid"] for w in ctis):
            raise RuntimeError("CTIが切り替わったため、履歴はクリアしていません。")
        before = windows.snapshot(context["cti"]["hwnd"], 3000)
        if before["errors"]:
            raise RuntimeError("クリア前の画面を確認できませんでした。")
        controls = descendants(before["nodes"], context["page_hwnd"])
        buttons = [n for n in controls if n["text"] == "クリア" and "BUTTON" in n["class_name"]]
        if len(buttons) != 1:
            raise RuntimeError("架電履歴のクリアボタンを特定できませんでした。")
        tables = msaa_tables(read_msaa([context["history_hwnd"], context["staff_hwnd"]]))
        history = next((t for t in tables if t["hwnd"] == context["history_hwnd"]), None)
        staff = next((t for t in tables if t["hwnd"] == context["staff_hwnd"]), None)
        if not history or not staff or history["rows"] != context["rows"] or staff["rows"] != context["staff_rows"] or search_conditions(controls) != context["conditions"]:
            raise RuntimeError("集計後に履歴または検索条件が変わったため、クリアしていません。")
        self.check_cancelled()
        button = buttons[0]
        result = ctypes.c_size_t()
        ok = windows.api.SendMessageTimeoutW(button["parent_hwnd"], WM_COMMAND, button["control_id"] & 0xffff,
                                             button["hwnd"], SMTO_ABORTIFHUNG, 15000, ctypes.byref(result))
        if not ok:
            raise RuntimeError("クリアの応答を確認できませんでした。CTI側の状態を確認してください。")
        # Do not retry the command: a timeout can still mean that CTI processed it.
        self._history_context = None
        for _ in range(3):
            if self.cancel.wait(0.3):
                self.check_cancelled()
            after = windows.snapshot(context["cti"]["hwnd"], 3000)
            if after["errors"]:
                continue
            after_controls = descendants(after["nodes"], context["page_hwnd"])
            current = msaa_tables(read_msaa([n["hwnd"] for n in after_controls if n["class_name"].startswith("WindowsForms10.Window.")]))
            # Clear recreates the results grid and blanks the count label in CTI.
            # Rediscover the grid; the old handle and a numeric label are not reliable.
            candidates = [t for t in current if {"対応日時", "管理番号"}.issubset(t.get("headers", []))
                          or t["hwnd"] == context["history_hwnd"]]
            if len(candidates) == 1 and not candidates[0]["rows"]:
                staff_after = next((t for t in current if {"選択", "スタッフ"}.issubset(t.get("headers", []))), None)
                self.clear_history_info = {
                    "rows_before": len(context["rows"]), "rows_after": 0,
                    "search_conditions_preserved": search_conditions(after_controls) == context["conditions"],
                    "staff_selection_preserved": bool(staff_after and staff_after["rows"] == context["staff_rows"]),
                    "history_visible_before": next((n["visible"] for n in before["nodes"] if n["hwnd"] == context["page_hwnd"]), None),
                    "history_visible_after": next((n["visible"] for n in after["nodes"] if n["hwnd"] == context["page_hwnd"]), None),
                }
                return
        raise RuntimeError("クリア後に履歴が空になったことを確認できませんでした。")

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
