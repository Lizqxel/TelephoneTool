"""読み込み済みの履歴を読み取って重複IDエラーを再現する。検索・保存DB更新はしない。"""
from collections import defaultdict
from datetime import datetime
import hashlib
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from services.call_statistics import hourly_history_totals, normalize_name
from services.cti_accessibility import Windows, msaa_tables, read_msaa
from services.cti_statistics_service import descendants


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    windows = Windows()
    ctis = windows.windows("CTIメイン")
    if len(ctis) != 1:
        raise RuntimeError("Expected one CTI window")
    native = windows.snapshot(ctis[0]["hwnd"], 3000)
    pages = [n for n in native["nodes"] if n["text"] == "架電履歴" and n["class_name"].startswith("WindowsForms10.Window.")]
    if len(pages) != 1:
        raise RuntimeError("Expected one native history page")
    controls = descendants(native["nodes"], pages[0]["hwnd"])
    tables = msaa_tables(read_msaa([n["hwnd"] for n in controls if n["class_name"].startswith("WindowsForms10.Window.")]))
    history = [t for t in tables if {"対応日時", "IP電話通話時間", "管理番号"}.issubset(t["headers"])]
    if len(history) != 1:
        raise RuntimeError("Expected one history table")
    rows = history[0]["rows"]
    groups = defaultdict(list)
    for index, row in enumerate(rows):
        groups[(row.get("通話ユニークID") or "").strip()].append(index)
    duplicate_groups = []
    for unique, indexes in groups.items():
        if len(indexes) < 2:
            continue
        placeholder = unique.lower() in ("", "-", "0", "(null)", "null", "none", "n/a", "なし", "未設定")
        duplicate_groups.append({"id_digest": hashlib.sha256(unique.encode()).hexdigest()[:10],
                                 "placeholder": placeholder, "placeholder_value": unique if placeholder else None,
                                 "rows": [{"row_index": i, **{k: rows[i].get(k) for k in ("対応日時", "架電結果", "架電結果詳細", "IP電話通話時間")}} for i in indexes]})
    reproduction = []
    staff_names = {normalize_name(r.get("対応スタッフ")) for r in rows}
    for staff in staff_names:
        own = [r for r in rows if normalize_name(r.get("対応スタッフ")) == staff]
        if not own:
            continue
        day = datetime.fromisoformat(own[0]["対応日時"].replace("/", "-")).date().isoformat()
        try:
            totals = hourly_history_totals(own, day, staff)
            reproduction.append({"rows": len(own), "error": None, "total_calls": sum(h["total_calls"] for h in totals["counters"].values())})
        except ValueError as exc:
            reproduction.append({"rows": len(own), "error": str(exc)})
    folder = ROOT / "logs" / "cti_inspection" / ("duplicate_id_diagnosis_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    folder.mkdir(parents=True)
    (folder / "rows.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    report = {"history_rows": len(rows), "duplicate_groups": duplicate_groups, "reproduction": reproduction,
              "history_visible": pages[0]["visible"], "output": str(folder)}
    (folder / "diagnosis.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=True))


if __name__ == "__main__":
    main()
