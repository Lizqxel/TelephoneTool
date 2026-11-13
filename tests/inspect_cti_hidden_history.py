"""架電画面を表示したまま、非表示の架電履歴をMSAAで直接読む診断。

取得対象は履歴一覧に読み込み済みの行。現在の架電対象との一致は条件にしない。
タブ変更・検索・入力は行わない。標準ライブラリとWindows標準APIのみ使用。
"""

import base64
from datetime import datetime
import json
import re
import subprocess
import sys
from types import SimpleNamespace
import unicodedata

from inspect_cti_history import HISTORY_COLUMNS, ROOT, Windows, extract_table_rows, uia_snapshot


SCRIPT = r"""
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName Accessibility
$source=@'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Accessibility;

public class AccNode {
    public int id, parent_id, depth, child_id, role, state, child_count;
    public string name, value, error;
}
public class AccResult {
    public long hwnd;
    public int hresult;
    public bool truncated;
    public List<AccNode> nodes = new List<AccNode>();
    public List<string> errors = new List<string>();
}
public static class HiddenHistoryAcc {
    [DllImport("oleacc.dll")]
    static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out IAccessible result);
    [DllImport("oleacc.dll")]
    static extern int AccessibleChildren(IAccessible container, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] object[] children, out int obtained);

    static void Walk(IAccessible acc, int childId, int parent, int depth, AccResult result, Stopwatch clock) {
        if (result.nodes.Count >= 3000 || depth > 16 || clock.Elapsed.TotalSeconds > 15) {
            result.truncated=true; return;
        }
        var node=new AccNode { id=result.nodes.Count, parent_id=parent, depth=depth, child_id=childId };
        try { node.name=acc.get_accName(childId); } catch {}
        try { node.role=Convert.ToInt32(acc.get_accRole(childId)); } catch {}
        try { node.state=Convert.ToInt32(acc.get_accState(childId)); } catch {}
        if ((node.state & 0x20000000) == 0) {
            try { node.value=acc.get_accValue(childId); } catch (Exception ex) { node.error=ex.Message; }
        }
        if (childId == 0) { try { node.child_count=acc.accChildCount; } catch {} }
        result.nodes.Add(node);
        if (node.child_count <= 0) return;
        int count=Math.Min(node.child_count, 3000-result.nodes.Count);
        if (count < node.child_count) result.truncated=true;
        object[] children=new object[count]; int obtained;
        int hr=AccessibleChildren(acc, 0, count, children, out obtained);
        if (hr < 0) { result.errors.Add("AccessibleChildren HRESULT: " + hr); return; }
        for (int i=0; i<obtained; i++) {
            if (clock.Elapsed.TotalSeconds > 15) { result.truncated=true; break; }
            var child=children[i] as IAccessible;
            if (child != null) Walk(child, 0, node.id, depth+1, result, clock);
            else if (children[i] is int) Walk(acc, (int)children[i], node.id, depth+1, result, clock);
        }
    }
    public static AccResult Read(long hwnd) {
        var result=new AccResult { hwnd=hwnd };
        try {
            IAccessible acc; Guid iid=new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
            result.hresult=AccessibleObjectFromWindow(new IntPtr(hwnd), 0xfffffffc, ref iid, out acc);
            if (result.hresult >= 0 && acc != null) {
                // Container trees may repeat descendants. Only enumerate native tables.
                if (Convert.ToInt32(acc.get_accRole(0)) == 24)
                    Walk(acc, 0, -1, 0, result, Stopwatch.StartNew());
            }
        } catch (Exception ex) { result.errors.Add(ex.ToString()); }
        return result;
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location)
$cfg=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__CONFIG__')) | ConvertFrom-Json
$results=@(foreach($h in $cfg.handles) { [HiddenHistoryAcc]::Read([long]$h) })
$results | ConvertTo-Json -Depth 12 -Compress
"""


def extract_history(results):
    tables = []
    for result in results:
        nodes = result["nodes"]
        if not nodes or nodes[0]["role"] != 24:
            continue
        headers = [n["name"] for n in nodes if n["role"] == 25]
        if not {"対応日時", "架電結果", "管理番号"}.issubset(headers):
            continue
        rows = []
        for row in (n for n in nodes if n["role"] == 28 and n["parent_id"] == nodes[0]["id"]):
            cells = [n for n in nodes if n["parent_id"] == row["id"] and n["role"] == 29]
            if not cells:
                continue  # Header row.
            values, unmapped = {}, []
            for cell in cells:
                labels = [h for h in headers if cell.get("name") == h or (cell.get("name") or "").startswith(h + " 行 ")]
                if len(labels) == 1 and labels[0] not in values and cell["value"] is not None:
                    values[labels[0]] = cell["value"]
                else:
                    unmapped.append(cell["id"])
            missing = [h for h in HISTORY_COLUMNS if h not in values]
            rows.append({"row_index": len(rows), "complete": not missing and not unmapped and not result["truncated"] and not result["errors"],
                         "columns_read": len(HISTORY_COLUMNS) - len(missing),
                         "values": {h: values.get(h) for h in HISTORY_COLUMNS},
                         "missing_columns": missing, "unmapped_cell_ids": unmapped})
        tables.append({"hwnd": result["hwnd"], "headers_read": headers,
                       "truncated": result["truncated"], "rows": rows})
    return {"requested_columns": HISTORY_COLUMNS, "tables": tables}


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    windows = Windows()
    candidates = windows.windows("CTIメイン")
    if len(candidates) != 1:
        raise RuntimeError("Expected exactly one CTI window")
    snapshot = windows.snapshot(candidates[0]["hwnd"], 2000)
    pages = [n for n in snapshot["nodes"] if n["text"] == "架電履歴" and "BUTTON" not in n["class_name"]]
    if len(pages) != 1:
        raise RuntimeError("Expected exactly one history page")
    page = pages[0]
    ids = {page["hwnd"]}
    for node in snapshot["nodes"]:
        if node["parent_hwnd"] in ids:
            ids.add(node["hwnd"])
    # DataGridView has this generic WinForms class; inspect all matching controls,
    # rather than relying on HWNDs saved in a previous run.
    controls = [n for n in snapshot["nodes"] if n["hwnd"] in ids and n["class_name"].startswith("WindowsForms10.Window.")]
    config = base64.b64encode(json.dumps({"handles": [n["hwnd"] for n in controls]}).encode()).decode()
    script = base64.b64encode(SCRIPT.replace("__CONFIG__", config).encode("utf-16-le")).decode()
    process = subprocess.run(["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", script],
                             capture_output=True, timeout=90, creationflags=subprocess.CREATE_NO_WINDOW)
    if process.returncode:
        raise RuntimeError(process.stderr.decode("utf-8", errors="replace"))
    results = json.loads(process.stdout.decode("utf-8-sig"))
    if isinstance(results, dict):
        results = [results]
    active_uia = uia_snapshot(candidates[0]["hwnd"], SimpleNamespace(max_nodes=2000, max_depth=30))
    active_tables = extract_table_rows(active_uia.get("nodes", []))
    normalize_phone = lambda value: re.sub(r"[^0-9]", "", unicodedata.normalize("NFKC", value or ""))
    active_phones = {normalize_phone(row["values"].get("電話番号"))
                     for table in active_tables for row in table["rows"]}
    active_phones.discard("")
    after = windows.snapshot(candidates[0]["hwnd"], 2000)
    def visible_pages(nodes):
        return [{"hwnd": n["hwnd"], "name": n["text"]} for n in nodes
                if n["visible"] and n["class_name"].startswith("WindowsForms10.Window.")
                and n["text"] in ("架電画面", "架電履歴")]
    before_pages = visible_pages(snapshot["nodes"])
    after_pages = visible_pages(after["nodes"])
    history_after = [n for n in after["nodes"] if n["hwnd"] == page["hwnd"]]
    directory = ROOT / "logs" / "cti_inspection" / ("hidden_msaa_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    directory.mkdir(parents=True)
    report = {"captured_at": datetime.now().astimezone().isoformat(), "page": page,
              "visible_pages_before": before_pages, "visible_pages_after": after_pages,
              "display_unchanged": before_pages == after_pages,
              "history_visible_after": history_after[0]["visible"] if history_after else None,
              "controls": controls, "results": results}
    (directory / "structure.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    extracted = {"captured_at": report["captured_at"], "source": "MSAA: hidden history page",
                 "history_visible_before": page["visible"],
                 "history_visible_after": report["history_visible_after"],
                 "display_unchanged": report["display_unchanged"], **extract_history(results)}
    for table in extracted["tables"]:
        for row in table["rows"]:
            phone = normalize_phone(row["values"].get("電話番号"))
            row["phone_matches_visible_table"] = phone in active_phones if phone and active_phones else None
    extracted["data_scope"] = "Previously loaded rows in the hidden history page; no search or refresh was performed."
    (directory / "visible_table_rows.json").write_text(json.dumps(active_tables, ensure_ascii=False, indent=2), encoding="utf-8")
    (directory / "requested_history_rows.json").write_text(json.dumps(extracted, ensure_ascii=False, indent=2), encoding="utf-8")
    summary = {"captured_at": report["captured_at"], "output": str(directory),
               "page_visible": page["visible"], "history_visible_after": report["history_visible_after"],
               "display_unchanged": report["display_unchanged"],
               "visible_pages": after_pages, "controls": len(controls),
               "tables": [{"hwnd": t["hwnd"], "headers_read": t["headers_read"],
                           "rows": [{k: v for k, v in r.items() if k != "values"} for r in t["rows"]]}
                          for t in extracted["tables"]]}
    (directory / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=True))
    return 0 if len(extracted["tables"]) == 1 and extracted["tables"][0]["rows"] and all(r["complete"] for r in extracted["tables"][0]["rows"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())
