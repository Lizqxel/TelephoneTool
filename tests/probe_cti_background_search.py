"""架電画面を維持してスタッフ選択・架電履歴検索を検証する。

既定は調査のみ。--execute でログインスタッフだけを選択して検索する。
--verify-toggle は対象のチェックを一度反転して戻し、非表示で操作できるか検証する。
画面を表示するAPIやグローバルなマウス・キー入力は使用しない。
"""

import argparse
import base64
import ctypes
from datetime import datetime
import json
from pathlib import Path
import re
import subprocess
import sys
import time
from types import SimpleNamespace
import unicodedata

from inspect_cti_history import ROOT, Windows, extract_table_rows, uia_snapshot
from inspect_cti_hidden_history import SCRIPT, extract_history


ACTIONS = r"""
    static List<string> methodTrace=new List<string>();
    static int[] calibratedCenters;
    [StructLayout(LayoutKind.Sequential)] struct Point { public int x, y; }
    [StructLayout(LayoutKind.Sequential)] struct Rect { public int left, top, right, bottom; }
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref Point point);
    [DllImport("user32.dll", SetLastError=true)] static extern IntPtr SendMessageTimeout(
        IntPtr hwnd, uint msg, UIntPtr wp, IntPtr lp, uint flags, uint timeout, out UIntPtr result);
    static void CellClick(long hwnd, IAccessible cell) {
        int x, y, width, height; cell.accLocation(out x, out y, out width, out height, 0);
        Point p;
        if (width > 0 && height > 0) {
            p=new Point { x=x+width/2, y=y+height/2 };
            if (!ScreenToClient(new IntPtr(hwnd), ref p)) throw new Exception("ScreenToClient failed");
        } else {
            var match=System.Text.RegularExpressions.Regex.Match(cell.get_accName(0) ?? "", @" 行 (\d+)");
            if (!match.Success || calibratedCenters == null) throw new Exception("Checkbox accessible bounds unavailable");
            int row=Int32.Parse(match.Groups[1].Value);
            if (row*2+1 >= calibratedCenters.Length) throw new Exception("Calibration row missing");
            p=new Point { x=calibratedCenters[row*2], y=calibratedCenters[row*2+1] };
            methodTrace.Add("validated_layout_reference");
        }
        Rect client;
        if (!GetClientRect(new IntPtr(hwnd), out client) || p.x < 0 || p.y < 0 || p.x >= client.right || p.y >= client.bottom || p.x > 32767 || p.y > 32767)
            throw new Exception("Checkbox position outside client area");
        IntPtr lp=new IntPtr((p.y << 16) | p.x); UIntPtr result;
        if (SendMessageTimeout(new IntPtr(hwnd), 0x201, new UIntPtr(1), lp, 2, 2000, out result)==IntPtr.Zero)
            throw new Exception("Cell mouse-down message timed out");
        if (SendMessageTimeout(new IntPtr(hwnd), 0x202, UIntPtr.Zero, lp, 2, 2000, out result)==IntPtr.Zero)
            throw new Exception("Cell mouse-up message timed out");
    }
    static IAccessible Root(long hwnd) {
        Guid iid=new Guid("618736e0-3c3d-11cf-810c-00aa00389b71"); IAccessible acc;
        int hr=AccessibleObjectFromWindow(new IntPtr(hwnd), 0xfffffffc, ref iid, out acc);
        if (hr < 0 || acc == null) throw new Exception("AccessibleObjectFromWindow failed: " + hr);
        return acc;
    }
    static List<IAccessible> Objects(IAccessible parent) {
        int count=parent.accChildCount;
        if (count > 1000) throw new Exception("Too many children for action probe");
        var result=new List<IAccessible>(); if (count == 0) return result;
        object[] items=new object[count]; int got;
        int hr=AccessibleChildren(parent, 0, count, items, out got);
        if (hr < 0) throw new Exception("AccessibleChildren failed: " + hr);
        for (int i=0; i<got; i++) {
            var child=items[i] as IAccessible;
            if (child == null && items[i] is int) child=parent.get_accChild(items[i]) as IAccessible;
            if (child == null) throw new Exception("Child does not expose IAccessible");
            result.Add(child);
        }
        return result;
    }
    static string Normalize(string s) {
        return System.Text.RegularExpressions.Regex.Replace((s ?? "").Normalize(System.Text.NormalizationForm.FormKC), @"\s+", "");
    }
    static bool Checked(IAccessible cell) {
        bool value;
        if (Boolean.TryParse(cell.get_accValue(0), out value)) return value;
        return (Convert.ToInt32(cell.get_accState(0)) & 16) != 0;
    }
    static void SetChecked(long hwnd, IAccessible cell, bool target) {
        if (Checked(cell) == target) return;
        if (!String.IsNullOrEmpty(cell.get_accDefaultAction(0))) cell.accDoDefaultAction(0);
        if (Checked(cell) == target) return;
        try {
            cell.set_accValue(0, target ? "True" : "False");
            methodTrace.Add("put_accValue");
        } catch(Exception ex) { methodTrace.Add("put_accValue_failed: " + ex.Message); }
        if (Checked(cell) == target) return;
        CellClick(hwnd, cell);
        methodTrace.Add("cell_mouse_messages");
        if (Checked(cell) != target) throw new Exception("Checkbox cell messages returned without changing state");
    }
    public static Dictionary<string, object> SelectStaff(long hwnd, string staff, bool execute, bool verifyToggle, int[] centers) {
        var result=new Dictionary<string, object>(); var names=new List<string>();
        methodTrace.Clear(); result["methods"]=methodTrace;
        calibratedCenters=centers;
        var boxes=new List<IAccessible>(); var original=new List<bool>();
        var actions=new List<string>(); result["actions"]=actions;
        try {
            var root=Root(hwnd);
            if (Convert.ToInt32(root.get_accRole(0)) != 24) throw new Exception("Staff control is not a table");
            foreach (var row in Objects(root)) {
                if (Convert.ToInt32(row.get_accRole(0)) != 28) continue;
                IAccessible name=null, box=null;
                foreach (var cell in Objects(row)) {
                    if (Convert.ToInt32(cell.get_accRole(0)) != 29) continue;
                    string label=cell.get_accName(0) ?? "";
                    if (label.StartsWith("スタッフ 行 ")) name=cell;
                    if (label.StartsWith("選択 行 ")) box=cell;
                }
                if (name != null && box != null) { names.Add(name.get_accValue(0)); boxes.Add(box); original.Add(Checked(box)); }
            }
            result["staff_names"]=names; result["checked_before"]=original;
            int target=-1, matches=0;
            for (int i=0; i<names.Count; i++) if (Normalize(names[i]) == Normalize(staff)) { target=i; matches++; }
            result["match_count"]=matches;
            if (matches != 1) throw new Exception("Logged-in staff name must match exactly one staff row");
            result["target_row"]=target; result["target_default_action"]=boxes[target].get_accDefaultAction(0);
            if (execute) {
                if (verifyToggle) {
                    SetChecked(hwnd, boxes[target], !original[target]); actions.Add("target_toggle_verified");
                    SetChecked(hwnd, boxes[target], original[target]); actions.Add("target_restored_after_probe");
                }
                for (int i=0; i<boxes.Count; i++) {
                    bool desired=i == target;
                    if (Checked(boxes[i]) != desired) { SetChecked(hwnd, boxes[i], desired); actions.Add("set_row_"+i+"_"+desired); }
                }
            }
            var after=new List<bool>(); foreach(var box in boxes) after.Add(Checked(box));
            result["checked_after"]=after; result["ok"]=true;
        } catch(Exception ex) {
            result["ok"]=false; result["error"]=ex.ToString();
            if (execute) {
                try { for(int i=0; i<boxes.Count; i++) SetChecked(hwnd, boxes[i], original[i]); result["restored_after_failure"]=true; }
                catch(Exception restore) { result["restore_error"]=restore.ToString(); }
            }
        }
        return result;
    }
    public static Dictionary<string, object> Search(long hwnd, bool execute) {
        var result=new Dictionary<string, object>();
        try {
            var acc=Root(hwnd);
            result["name"]=acc.get_accName(0); result["role"]=Convert.ToInt32(acc.get_accRole(0));
            result["default_action"]=acc.get_accDefaultAction(0);
            if ((string)result["name"] != "検索" || (int)result["role"] != 43) throw new Exception("Search button identity mismatch");
            if (execute) acc.accDoDefaultAction(0);
            result["invoked"]=execute; result["ok"]=true;
        } catch(Exception ex) { result["ok"]=false; result["error"]=ex.ToString(); }
        return result;
    }
"""


def bridge(config, command=None):
    script = SCRIPT
    if command:
        script = script.replace("public static AccResult Read", ACTIONS + "\n    public static AccResult Read")
        start = script.index("$results=@(foreach")
        script = script[:start] + command + " | ConvertTo-Json -Depth 12 -Compress\n"
    config64 = base64.b64encode(json.dumps(config).encode()).decode()
    encoded = base64.b64encode(script.replace("__CONFIG__", config64).encode("utf-16-le")).decode()
    run = subprocess.run(["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", encoded],
                         capture_output=True, timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
    if run.returncode:
        raise RuntimeError(run.stderr.decode("utf-8", errors="replace"))
    return json.loads(run.stdout.decode("utf-8-sig"))


def descendants(nodes, root):
    ids = {root}
    for node in nodes:
        if node["parent_hwnd"] in ids:
            ids.add(node["hwnd"])
    return [node for node in nodes if node["hwnd"] in ids]


def visible_pages(nodes):
    return [(n["hwnd"], n["text"]) for n in nodes if n["visible"]
            and n["class_name"].startswith("WindowsForms10.Window.")
            and n["text"] in ("架電画面", "架電履歴")]


def calibrated_centers(path, cti, staff_table, native_nodes):
    """Reuse measured cell locations after checking process, size and staff order."""
    reference = json.loads(path.read_text(encoding="utf-8"))
    if reference["window"]["pid"] != cti["pid"] or reference["window"]["hwnd"] != cti["hwnd"]:
        raise RuntimeError("Layout reference belongs to a different CTI process/window")
    nodes = reference["uia"]["nodes"]
    tables = [n for n in nodes if n["hwnd"] == staff_table["hwnd"] and "grid_rows" in n]
    if len(tables) != 1:
        raise RuntimeError("Staff table handle does not match layout reference")
    table = tables[0]
    native = next(n for n in native_nodes if n["hwnd"] == table["hwnd"])
    rect = native["rect"]
    old_native = next(n for n in reference["win32"]["nodes"] if n["hwnd"] == table["hwnd"])
    old_rect = old_native["rect"]
    # Native window bounds include the border; UIA table bounds are the client area.
    if [rect[2] - rect[0], rect[3] - rect[1]] != [old_rect[2] - old_rect[0], old_rect[3] - old_rect[1]]:
        raise RuntimeError("Staff grid size changed; calibration cannot be reused")
    ids = {table["id"]}
    for node in nodes:
        if node["parent_id"] in ids:
            ids.add(node["id"])
    reference_rows = next(t["rows"] for t in extract_table_rows(nodes) if t["hwnd"] == table["hwnd"])
    old_names = [r["values"].get("スタッフ") for r in reference_rows]
    current_names = [n["value"] for n in staff_table["nodes"] if n["role"] == 29 and (n["name"] or "").startswith("スタッフ 行 ")]
    normalize = lambda name: "".join(unicodedata.normalize("NFKC", name or "").split())
    if list(map(normalize, old_names)) != list(map(normalize, current_names)):
        raise RuntimeError("Staff order changed; calibration cannot be reused")
    cells = {n["row"]: n for n in nodes if n["id"] in ids and n.get("column") == 0 and "row" in n}
    centers = []
    for row in range(len(current_names)):
        x, y, width, height = cells[row]["rect"]
        if width <= 0 or height <= 0:
            raise RuntimeError("Layout reference has an empty cell rectangle")
        centers += [int(x - table["rect"][0] + width / 2), int(y - table["rect"][1] + height / 2)]
    return centers


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--verify-toggle", action="store_true")
    parser.add_argument("--search-method", choices=("command", "msaa"), default="command")
    parser.add_argument("--layout-reference", type=Path, help="Measured visible layout; validates process, grid size and staff order before reuse")
    args = parser.parse_args()
    if args.verify_toggle and not args.execute:
        parser.error("--verify-toggle requires --execute")
    windows = Windows()
    windows.api.GetForegroundWindow.restype = ctypes.c_void_p
    foreground_before = windows.api.GetForegroundWindow()
    ctis = windows.windows("CTIメイン")
    if len(ctis) != 1:
        raise RuntimeError("Expected one CTI window")
    hwnd = ctis[0]["hwnd"]
    before = windows.snapshot(hwnd, 2000)
    uia = uia_snapshot(hwnd, SimpleNamespace(max_nodes=2000, max_depth=30))
    status_ids = {n["id"] for n in uia.get("nodes", []) if n["control_type"] == "ControlType.StatusBar"}
    staff_labels = []
    for node in uia.get("nodes", []):
        if node["parent_id"] in status_ids:
            match = re.match(r"^(.+?)\s*さん(?:\s*[:：].*)?$", node["name"].strip())
            if match:
                staff_labels.append(match.group(1).strip())
    if len(staff_labels) != 1:
        raise RuntimeError("Logged-in staff footer must be unique")
    pages = [n for n in before["nodes"] if n["text"] == "架電履歴" and "BUTTON" not in n["class_name"]]
    if len(pages) != 1:
        raise RuntimeError("History page must be unique")
    if args.execute and pages[0]["visible"]:
        raise RuntimeError("Background action probe requires the history page to be hidden; no action performed")
    controls = descendants(before["nodes"], pages[0]["hwnd"])
    handles = [n["hwnd"] for n in controls if n["class_name"].startswith("WindowsForms10.Window.")]
    results = bridge({"handles": handles})
    if isinstance(results, dict):
        results = [results]
    staff_tables = [r for r in results if any(n["role"] == 25 and n["name"] == "スタッフ" for n in r["nodes"])]
    if len(staff_tables) != 1:
        raise RuntimeError("Staff table must be unique")
    centers = calibrated_centers(args.layout_reference, ctis[0], staff_tables[0], before["nodes"]) if args.layout_reference else []
    search_buttons = [n for n in controls if n["text"] == "検索" and "BUTTON" in n["class_name"]]
    if len(search_buttons) != 1:
        raise RuntimeError("History search button must be unique")
    directory = ROOT / "logs" / "cti_inspection" / ("background_search_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    directory.mkdir(parents=True)
    report = {"started_at": datetime.now().astimezone().isoformat(), "staff_name": staff_labels[0],
              "execute": args.execute, "visible_pages_before": visible_pages(before["nodes"]),
              "foreground_before": foreground_before,
              "history_visible_before": pages[0]["visible"], "before_history": extract_history(results)}
    def save():
        (directory / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    save()
    selection = bridge({"hwnd": staff_tables[0]["hwnd"], "staff": staff_labels[0],
                        "execute": args.execute, "verify": args.verify_toggle, "centers": centers},
                       "[HiddenHistoryAcc]::SelectStaff([long]$cfg.hwnd, [string]$cfg.staff, [bool]$cfg.execute, [bool]$cfg.verify, [int[]]$cfg.centers)")
    report["selection"] = selection
    save()
    if selection["ok"]:
        report["search"] = bridge({"hwnd": search_buttons[0]["hwnd"], "execute": args.execute and args.search_method == "msaa"},
                                   "[HiddenHistoryAcc]::Search([long]$cfg.hwnd, [bool]$cfg.execute)")
        if args.execute and args.search_method == "command" and report["search"].get("ok"):
            button = search_buttons[0]
            result = ctypes.c_size_t()
            ok = windows.api.SendMessageTimeoutW(button["parent_hwnd"], 0x111,
                                                 button["control_id"] & 0xffff, button["hwnd"],
                                                 2, 15000, ctypes.byref(result))
            report["search"].update(method="WM_COMMAND/BN_CLICKED", invoked=bool(ok), ok=bool(ok),
                                     message_result=result.value)
        save()
        if args.execute and report["search"].get("ok"):
            time.sleep(1)
            after_history = bridge({"handles": [r["hwnd"] for r in results if r["nodes"] and r["nodes"][0]["role"] == 24]})
            if isinstance(after_history, dict):
                after_history = [after_history]
            report["after_history"] = extract_history(after_history)
            report["history_result_changed"] = report["before_history"] != report["after_history"]
    after = windows.snapshot(hwnd, 2000)
    report["visible_pages_after"] = visible_pages(after["nodes"])
    report["display_unchanged"] = report["visible_pages_before"] == report["visible_pages_after"]
    report["history_visible_after"] = next((n["visible"] for n in after["nodes"] if n["hwnd"] == pages[0]["hwnd"]), None)
    report["foreground_after"] = windows.api.GetForegroundWindow()
    report["foreground_unchanged"] = report["foreground_before"] == report["foreground_after"]
    normalize_name = lambda name: "".join(unicodedata.normalize("NFKC", name or "").split())
    rows = [row for table in report.get("after_history", {}).get("tables", []) for row in table["rows"]]
    report["after_rows_match_staff"] = all(normalize_name(row["values"].get("対応スタッフ")) == normalize_name(staff_labels[0]) for row in rows) if rows else None
    report["checkbox_change_verified"] = bool(selection.get("actions"))
    report["search_refresh_verified"] = bool(report.get("search", {}).get("invoked") and
                                              report.get("history_result_changed") and rows and
                                              all(row["complete"] for row in rows) and report["after_rows_match_staff"])
    save()
    summary = {"output": str(directory), "staff_detected": True, "execute": args.execute,
                      "selection": {k: v for k, v in selection.items() if k != "staff_names"},
                      "search": report.get("search"), "display_unchanged": report["display_unchanged"],
                      "foreground_unchanged": report["foreground_unchanged"],
                      "history_result_changed": report.get("history_result_changed"),
                      "history_visible_before": report["history_visible_before"],
                      "history_visible_after": report["history_visible_after"],
                      "checkbox_change_verified": report["checkbox_change_verified"],
                      "search_refresh_verified": report["search_refresh_verified"],
                      "after_rows_match_staff": report["after_rows_match_staff"],
                      "before_row_count": sum(len(t["rows"]) for t in report["before_history"]["tables"]),
                      "after_row_count": len(rows)}
    (directory / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    if "after_history" in report:
        (directory / "requested_history_rows.json").write_text(json.dumps(report["after_history"], ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=True))
    return 0 if selection["ok"] and report.get("search", {}).get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
