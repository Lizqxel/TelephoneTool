"""CTIのWin32/UI Automation構造を調査する手動実行用ツール（標準ライブラリのみ）。

python tests/inspect_cti_history.py

CTIで「架電履歴」タブを開いてから実行する。
読み取り専用で、タブ切り替え・ボタン押下・入力・スクロールは行わない。
出力は顧客情報を含む可能性があるため、git管理対象外の logs 以下へ保存する。
"""

import argparse
import base64
import ctypes
from ctypes import wintypes
from datetime import datetime
import json
from pathlib import Path
import subprocess
import sys
import time


ROOT = Path(__file__).resolve().parents[1]
HISTORY_COLUMNS = [
    "対応日時", "見込", "管理番号", "IP電話使用時間", "IP電話通話時間", "対話者続柄",
    "架電結果", "架電結果詳細", "会社", "ユニット", "グループ", "対応スタッフ", "商材",
    "電話番号", "IP電話架電番号", "SIPサーバーIPアドレス", "通話ユニークID", "先リスト名", "元リスト名",
]


class Windows:
    def __init__(self):
        self.api = ctypes.WinDLL("user32", use_last_error=True)
        self.callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        signatures = {
            "EnumWindows": ([self.callback_type, wintypes.LPARAM], wintypes.BOOL),
            "EnumChildWindows": ([wintypes.HWND, self.callback_type, wintypes.LPARAM], wintypes.BOOL),
            "GetWindowTextW": ([wintypes.HWND, wintypes.LPWSTR, ctypes.c_int], ctypes.c_int),
            "GetClassNameW": ([wintypes.HWND, wintypes.LPWSTR, ctypes.c_int], ctypes.c_int),
            "GetParent": ([wintypes.HWND], wintypes.HWND),
            "GetWindowRect": ([wintypes.HWND, ctypes.POINTER(wintypes.RECT)], wintypes.BOOL),
            "GetWindowThreadProcessId": ([wintypes.HWND, ctypes.POINTER(wintypes.DWORD)], wintypes.DWORD),
            "GetDlgCtrlID": ([wintypes.HWND], ctypes.c_int),
            "IsWindowVisible": ([wintypes.HWND], wintypes.BOOL),
            "IsWindowEnabled": ([wintypes.HWND], wintypes.BOOL),
            "GetWindowLongW": ([wintypes.HWND, ctypes.c_int], wintypes.LONG),
            "SendMessageTimeoutW": ([wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                                     wintypes.LPARAM, wintypes.UINT, wintypes.UINT,
                                     ctypes.POINTER(ctypes.c_size_t)], ctypes.c_ssize_t),
        }
        for name, (args, result) in signatures.items():
            method = getattr(self.api, name)
            method.argtypes, method.restype = args, result

    def string(self, method, hwnd):
        buffer = ctypes.create_unicode_buffer(4096)
        getattr(self.api, method)(hwnd, buffer, len(buffer))
        return buffer.value

    def windows(self, title):
        found = []

        @self.callback_type
        def visit(hwnd, _):
            text = self.string("GetWindowTextW", hwnd)
            if title in text and self.api.IsWindowVisible(hwnd):
                pid = wintypes.DWORD()
                self.api.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                found.append({"hwnd": hwnd, "title": text, "pid": pid.value,
                              "class_name": self.string("GetClassNameW", hwnd)})
            return True

        ctypes.set_last_error(0)
        ok = self.api.EnumWindows(visit, 0)
        if not ok:
            raise OSError(ctypes.get_last_error(), "EnumWindows failed; desktop access may be restricted")
        return found

    def snapshot(self, hwnd, max_nodes, seconds=20):
        nodes, errors = [], []
        deadline = time.monotonic() + seconds

        def read(handle):
            class_name = self.string("GetClassNameW", handle)
            # ES_PASSWORD is meaningful for edit controls, not arbitrary window classes.
            password = "edit" in class_name.lower() and bool(self.api.GetWindowLongW(handle, -16) & 0x20)
            buffer = ctypes.create_unicode_buffer(4096)
            result = ctypes.c_size_t()
            ctypes.set_last_error(0)
            ok = True if password else self.api.SendMessageTimeoutW(
                handle, 0x000D, len(buffer), ctypes.addressof(buffer), 0x0002, 200,
                ctypes.byref(result))
            rect = wintypes.RECT()
            self.api.GetWindowRect(handle, ctypes.byref(rect))
            return {"hwnd": handle, "parent_hwnd": self.api.GetParent(handle) or 0,
                    "control_id": self.api.GetDlgCtrlID(handle), "class_name": class_name,
                    "text": "<password omitted>" if password else buffer.value,
                    "text_status": "ok" if ok else "failed_or_timeout",
                    "win32_error": 0 if ok else ctypes.get_last_error(),
                    "visible": bool(self.api.IsWindowVisible(handle)),
                    "enabled": bool(self.api.IsWindowEnabled(handle)),
                    "rect": [rect.left, rect.top, rect.right, rect.bottom]}

        nodes.append(read(hwnd))

        @self.callback_type
        def visit(handle, _):
            if len(nodes) >= max_nodes or time.monotonic() >= deadline:
                errors.append("Win32 snapshot truncated by node/time limit")
                return False
            try:
                nodes.append(read(handle))
            except Exception as exc:
                errors.append(str(exc))
            return True

        # Unlike the production fetcher, include invisible tab pages too.
        self.api.EnumChildWindows(hwnd, visit, 0)
        return {"nodes": nodes, "errors": errors}


UIA_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$cfg = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__CONFIG__')) | ConvertFrom-Json
$root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]([long]$cfg.hwnd))
$walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
$script:nodes = New-Object 'System.Collections.Generic.List[object]'
$script:issues = New-Object 'System.Collections.Generic.List[string]'
$script:truncated = $false
$clock = [Diagnostics.Stopwatch]::StartNew()

function Read-Element($element, $id, $parent, $depth) {
    $c = $element.Current
    $patterns = @($element.GetSupportedPatterns() | ForEach-Object { $_.ProgrammaticName })
    $record = [ordered]@{
        id=$id; parent_id=$parent; depth=$depth; name=$c.Name;
        control_type=$c.ControlType.ProgrammaticName; class_name=$c.ClassName;
        automation_id=$c.AutomationId; framework_id=$c.FrameworkId;
        hwnd=$c.NativeWindowHandle; enabled=$c.IsEnabled; offscreen=$c.IsOffscreen;
        is_password=$c.IsPassword; patterns=$patterns;
        rect=@($c.BoundingRectangle.X, $c.BoundingRectangle.Y,
               $c.BoundingRectangle.Width, $c.BoundingRectangle.Height)
    }
    if (!$c.IsPassword) {
        if ($patterns -contains 'ValuePatternIdentifiers.Pattern') {
            try { $p=$element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern); $record.value=$p.Current.Value }
            catch { $record.value_error=$_.Exception.Message }
        }
        if ($patterns -contains 'TextPatternIdentifiers.Pattern') {
            try { $p=$element.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern); $record.document_text=$p.DocumentRange.GetText(4096) }
            catch { $record.text_error=$_.Exception.Message }
        }
    }
    if ($patterns -contains 'SelectionItemPatternIdentifiers.Pattern') {
        try { $p=$element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern); $record.selected=$p.Current.IsSelected } catch {}
    }
    if ($patterns -contains 'GridPatternIdentifiers.Pattern') {
        try {
            $p=$element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern)
            $record.grid_rows=$p.Current.RowCount; $record.grid_columns=$p.Current.ColumnCount
        } catch { $record.grid_error=$_.Exception.Message }
    }
    if ($patterns -contains 'GridItemPatternIdentifiers.Pattern') {
        try {
            $p=$element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern)
            $record.row=$p.Current.Row; $record.column=$p.Current.Column
        } catch {}
    }
    if ($patterns -contains 'TableItemPatternIdentifiers.Pattern') {
        try {
            $p=$element.GetCurrentPattern([System.Windows.Automation.TableItemPattern]::Pattern)
            $record.column_headers=@($p.Current.GetColumnHeaderItems() | ForEach-Object { $_.Current.Name })
        } catch { $record.column_headers_error=$_.Exception.Message }
    }
    return [pscustomobject]$record
}

function Walk($element, $parent, $depth) {
    if ($script:nodes.Count -ge $cfg.max_nodes -or $clock.Elapsed.TotalSeconds -ge 20) {
        $script:truncated=$true; return
    }
    $id=$script:nodes.Count
    try { $record=Read-Element $element $id $parent $depth; $script:nodes.Add($record) }
    catch { $script:issues.Add($_.Exception.Message); return }
    if ($depth -ge $cfg.max_depth) { $script:truncated=$true; return }
    try {
        $child=$walker.GetFirstChild($element)
        while ($null -ne $child) {
            Walk $child $id ($depth+1)
            if ($script:truncated -and ($script:nodes.Count -ge $cfg.max_nodes -or $clock.Elapsed.TotalSeconds -ge 20)) { break }
            $child=$walker.GetNextSibling($child)
        }
    } catch { $script:issues.Add($_.Exception.Message) }
}

Walk $root -1 0
[ordered]@{nodes=@($script:nodes.ToArray()); truncated=$script:truncated; errors=@($script:issues.ToArray())} | ConvertTo-Json -Depth 12 -Compress
"""


def uia_snapshot(hwnd, args):
    config = {"hwnd": hwnd, "max_nodes": args.max_nodes, "max_depth": args.max_depth}
    encoded_config = base64.b64encode(json.dumps(config).encode("utf-8")).decode("ascii")
    script = UIA_SCRIPT.replace("__CONFIG__", encoded_config)
    encoded_script = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    try:
        process = subprocess.run(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA",
             "-EncodedCommand", encoded_script],
            capture_output=True, timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if process.returncode:
            return {"error": process.stderr.decode("utf-8", errors="replace")}
        return json.loads(process.stdout.decode("utf-8-sig"))
    except subprocess.TimeoutExpired:
        return {"error": "UI Automation timed out after 60 seconds"}
    except Exception as exc:
        return {"error": str(exc)}


def history_scope(nodes, tab_name):
    normalize = lambda text: "".join(text.replace("&", "").split())
    targets = [n for n in nodes if normalize(n.get("name", "")) == normalize(tab_name)]
    ids = {n["id"] for n in targets}
    for node in nodes:
        if node.get("parent_id") in ids:
            ids.add(node["id"])
    return [n for n in nodes if n["id"] in ids]


def summarize(report, scope):
    types = {}
    for node in scope:
        kind = node["control_type"]
        types[kind] = types.get(kind, 0) + 1
    tables = []
    for table in (node for node in scope if "grid_rows" in node):
        ids = {table["id"]}
        for node in scope:
            if node.get("parent_id") in ids:
                ids.add(node["id"])
        children = [node for node in scope if node["id"] in ids]
        cells = [node for node in children if "row" in node]
        tables.append({"id": table["id"], "hwnd": table["hwnd"],
                       "rows": table["grid_rows"], "columns": table["grid_columns"],
                       "headers_read": [n["name"] for n in children if n["control_type"] == "ControlType.Header"],
                       "cells_read": len(cells),
                       "cells_with_value": sum("value" in n for n in cells),
                       "column_indexes_read": sorted({n["column"] for n in cells}),
                       "patterns": table["patterns"]})
    return {"status": report["status"],
            "win32_count": len(report.get("win32", {}).get("nodes", [])),
            "uia_count": len(report.get("uia", {}).get("nodes", [])),
            "history_scope_count": len(scope), "history_types": types, "tables": tables}


def extract_table_rows(nodes):
    """Map each cell to its own header; hidden column gaps must not shift values."""
    tables = []
    for table in (node for node in nodes if "grid_rows" in node):
        ids = {table["id"]}
        for node in nodes:
            if node.get("parent_id") in ids:
                ids.add(node["id"])
        children = [node for node in nodes if node["id"] in ids]
        headers = [n["name"] for n in children if n["control_type"] == "ControlType.Header"]
        rows = {}
        for cell in (node for node in children if "row" in node):
            row = rows.setdefault(cell["row"], {"row_index": cell["row"], "values": {},
                                               "cells": [], "unmapped_cells": []})
            labels = cell.get("column_headers", [])
            if not labels:
                # Supports older saved snapshots. Match the UIA cell name, never zip headers to columns.
                labels = [label for label in headers if cell.get("name", "").startswith(label + " 行 ")]
            item = {"column_index": cell["column"], "headers": labels,
                    "value": cell.get("value"), "value_available": "value" in cell}
            if "value_error" in cell:
                item["error"] = cell["value_error"]
            row["cells"].append(item)
            if len(labels) == 1 and labels[0] not in row["values"] and "value" in cell:
                row["values"][labels[0]] = cell["value"]
            else:
                row["unmapped_cells"].append(item)
        tables.append({"table_id": table["id"], "hwnd": table["hwnd"],
                       "grid_rows": table["grid_rows"], "grid_columns": table["grid_columns"],
                       "headers_read": headers, "rows": [rows[key] for key in sorted(rows)]})
    return tables


def requested_history_rows(tables):
    matches = []
    for table in tables:
        if not set(table["headers_read"]) & set(HISTORY_COLUMNS):
            continue
        rows = []
        for row in table["rows"]:
            missing = [name for name in HISTORY_COLUMNS if name not in row["values"]]
            rows.append({"row_index": row["row_index"], "complete": not missing,
                         "columns_read": len(HISTORY_COLUMNS) - len(missing),
                         "values": {name: row["values"].get(name) for name in HISTORY_COLUMNS},
                         "missing_columns": missing})
        matches.append({"table_id": table["table_id"], "hwnd": table["hwnd"], "rows": rows})
    return {"requested_columns": HISTORY_COLUMNS, "tables": matches}


def write_report(report, directory, tab_name):
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "structure.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = ["CTI screen inspection", f"status: {report['status']}", f"time: {report['time']}"]
    if report.get("error"):
        lines.append(report["error"])
    if report.get("window"):
        lines.append("window: " + json.dumps(report["window"], ensure_ascii=False))
    for node in report.get("win32", {}).get("nodes", []):
        lines.append(f"WIN32 hwnd={node['hwnd']} parent={node['parent_hwnd']} visible={node['visible']} class={node['class_name']} text={node['text']!r}")
    uia = report.get("uia", {})
    nodes = uia.get("nodes", [])
    lines.append(f"\nUIA: {len(nodes)} elements")
    for node in nodes:
        lines.append("  " * node["depth"] + json.dumps(node, ensure_ascii=False))
    scope = history_scope(nodes, tab_name)
    (directory / "history.json").write_text(json.dumps(scope, ensure_ascii=False, indent=2), encoding="utf-8")
    (directory / "summary.json").write_text(json.dumps(summarize(report, scope), ensure_ascii=False, indent=2), encoding="utf-8")
    tables = extract_table_rows(nodes)
    (directory / "table_rows.json").write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")
    requested = {"captured_at": report["time"], "snapshot_status": report["status"],
                 **requested_history_rows(tables)}
    (directory / "requested_history_rows.json").write_text(json.dumps(requested, ensure_ascii=False, indent=2), encoding="utf-8")
    lines.append(f"Named tab/page and descendants: {len(scope)} (see history.json)")
    lines.append("UIA errors: " + json.dumps({k: v for k, v in uia.items() if k in ("error", "errors", "truncated")}, ensure_ascii=False))
    lines.append("Note: UIA tab content may be a sibling of TabItem. Inspect the full tree as well.")
    (directory / "structure.txt").write_text("\n".join(lines), encoding="utf-8")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--title", default="CTIメイン", help="Window title substring")
    parser.add_argument("--hwnd", type=lambda value: int(value, 0), help="Choose a listed window handle when multiple CTI windows exist")
    parser.add_argument("--tab-name", default="架電履歴")
    parser.add_argument("--max-nodes", type=int, default=2000)
    parser.add_argument("--max-depth", type=int, default=30)
    parser.add_argument("--output", type=Path, help="Output directory (default: logs/cti_inspection/<timestamp>)")
    args = parser.parse_args()
    if sys.platform != "win32":
        parser.error("Windows is required")
    if args.max_nodes < 1 or args.max_depth < 1:
        parser.error("Node and depth limits must be positive")
    directory = args.output or ROOT / "logs" / "cti_inspection" / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    report = {"time": datetime.now().astimezone().isoformat(), "status": "starting"}
    try:
        windows = Windows()
        candidates = windows.windows(args.title)
        report["candidates"] = candidates
        selected = [w for w in candidates if args.hwnd is None or w["hwnd"] == args.hwnd]
        if len(selected) != 1:
            report["status"] = "window_not_found" if not selected else "ambiguous_window"
            report["error"] = "Open CTI, or choose one candidate with --hwnd."
        else:
            report["window"] = selected[0]
            hwnd = selected[0]["hwnd"]
            report["win32"] = windows.snapshot(hwnd, args.max_nodes)
            report["uia"] = uia_snapshot(hwnd, args)
            report["status"] = "partial" if report["uia"].get("error") or report["uia"].get("errors") or report["uia"].get("truncated") or report["win32"]["errors"] else "captured"
    except Exception as exc:
        report.update(status="error", error=str(exc))
    write_report(report, directory, args.tab_name)
    # Keep customer details in local files; console contains only counts and status.
    scope = history_scope(report.get("uia", {}).get("nodes", []), args.tab_name)
    print(json.dumps({"status": report["status"], "output": str(directory),
                      "candidate_count": len(report.get("candidates", [])),
                      "win32_count": len(report.get("win32", {}).get("nodes", [])),
                      "uia_count": len(report.get("uia", {}).get("nodes", [])),
                      "history_scope_count": len(scope)}, ensure_ascii=True))
    return 0 if report["status"] == "captured" else 1


if __name__ == "__main__":
    raise SystemExit(main())
