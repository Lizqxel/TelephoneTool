"""Windows accessibility helpers; never activate a window or switch tabs."""

import base64
import ctypes
from ctypes import wintypes
import json
import subprocess
import time


def read_msaa(handles, columns=None):
    """各ハンドルの表をMSAAで読む。

    columns を渡すと、その列名をすべて持つ表だけは、その列のセルしか読まない（ほかの表はすべて読む）。
    CTIの表はセルを1つずつプロセス越しに読むため、50行×19列で約4秒かかる。使う列だけにすると数分の1になり、
    顧客の電話番号などの列も読まずに済む。
    """
    config = base64.b64encode(json.dumps({"handles": handles, "columns": list(columns) if columns else None}).encode()).decode()
    script = base64.b64encode(MSAA_SCRIPT.replace("__CONFIG__", config).encode("utf-16-le")).decode()
    try:
        process = subprocess.run(["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", script],
                                 capture_output=True, timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("CTIの読み取りがタイムアウトしました。処理が完了してから再実行してください。") from exc
    if process.returncode:
        raise RuntimeError("CTIのアクセシビリティ情報を読み取れませんでした。")
    results = json.loads(process.stdout.decode("utf-8-sig"))
    return [results] if isinstance(results, dict) else results


def read_status_bar(handles):
    """ステータスバー（MSAAのRole 23）の項目名。UI Automationと違い、CTIの最小化中も読める。"""
    config = base64.b64encode(json.dumps({"handles": handles}).encode()).decode()
    script = base64.b64encode(STATUS_SCRIPT.replace("__CONFIG__", config).encode("utf-16-le")).decode()
    try:
        process = subprocess.run(["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", script],
                                 capture_output=True, timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.TimeoutExpired:
        return []
    if process.returncode:
        return []
    names = json.loads(process.stdout.decode("utf-8-sig") or "[]")
    return [names] if isinstance(names, str) else [n for n in names if n]


def msaa_tables(results):
    tables = []
    for result in results:
        nodes = result["nodes"]
        if not nodes:
            continue
        if result["truncated"] or result["errors"]:
            raise RuntimeError("履歴の読み取りが完了しませんでした。件数や検索範囲を確認してください。")
        raw_headers = [n["name"] for n in nodes if n["role"] == 25]
        # 見込検索のように同名カラム（スタッフ名・スタッフID）がある表は2つ目以降を「名前#2」とする。
        headers, seen = [], {}
        for name in raw_headers:
            seen[name] = seen.get(name, 0) + 1
            headers.append(name if seen[name] == 1 else f"{name}#{seen[name]}")

        def belongs(cell, header):
            return cell["name"] == header or (cell["name"] or "").startswith(header + " 行 ")

        # 列を絞って読んだ表は、行ごとに読んだ列（index は行の中での位置）だけを持つ。
        kept = result.get("columns")
        rows = []
        for row in (n for n in nodes if n["parent_id"] == nodes[0]["id"] and n["role"] == 28):
            values = {}
            cells = [n for n in nodes if n["parent_id"] == row["id"] and n["role"] == 29]
            positions = [c.get("index", i) for i, c in enumerate(cells)]
            positional = ((len(cells) == len(raw_headers) or kept is not None) and len(set(positions)) == len(positions)
                          and all(0 <= p < len(raw_headers) and belongs(c, raw_headers[p]) for c, p in zip(cells, positions)))
            for index, cell in zip(positions, cells):
                if positional:
                    label = headers[index]
                else:
                    labels = [h for h in raw_headers if belongs(cell, h)]
                    if len(labels) != 1 or raw_headers.count(labels[0]) != 1:
                        raise RuntimeError("履歴のセルを読み取れませんでした。")
                    label = labels[0]
                if cell["value"] is None or cell.get("error"):
                    raise RuntimeError("履歴のセルを読み取れませんでした。")
                if label in values:
                    raise RuntimeError("履歴のカラム名が重複しています。")
                values[label] = cell["value"]
            if kept is not None and cells and len(values) != len(kept):
                raise RuntimeError("履歴のセルを読み取れませんでした。")
            if values:
                rows.append(values)
        tables.append({"hwnd": result["hwnd"], "headers": headers, "rows": rows})
    return tables

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
            "IsIconic": ([wintypes.HWND], wintypes.BOOL),
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
if ($cfg.control_type -eq 'StatusBar') {
    $condition = New-Object System.Windows.Automation.PropertyCondition -ArgumentList @(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::StatusBar)
    $root = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($null -eq $root) { throw 'CTI status bar not found' }
}
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
    config = {"hwnd": hwnd, "max_nodes": args.max_nodes, "max_depth": args.max_depth,
              "control_type": getattr(args, "control_type", "")}
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

STATUS_SCRIPT = r'''
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName Accessibility
$source=@'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;
public static class StatusBarAcc {
    [DllImport("oleacc.dll")]
    static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out IAccessible result);
    [DllImport("oleacc.dll")]
    static extern int AccessibleChildren(IAccessible container, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] object[] children, out int obtained);
    public static List<string> Read(long hwnd) {
        var names=new List<string>();
        IAccessible acc; Guid iid=new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
        if (AccessibleObjectFromWindow(new IntPtr(hwnd), 0xfffffffc, ref iid, out acc) < 0 || acc == null) return names;
        if (Convert.ToInt32(acc.get_accRole(0)) != 23) return names;
        int count=acc.accChildCount; var children=new object[count]; int obtained;
        if (AccessibleChildren(acc, 0, count, children, out obtained) < 0) return names;
        for (int i=0; i<obtained; i++) {
            try {
                var child=children[i] as IAccessible;
                names.Add(child != null ? child.get_accName(0) : acc.get_accName(children[i]));
            } catch {}
        }
        return names;
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location)
$cfg=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__CONFIG__')) | ConvertFrom-Json
$names=@(foreach($h in $cfg.handles) { [StatusBarAcc]::Read([long]$h) })
ConvertTo-Json -InputObject $names -Compress
'''

MSAA_SCRIPT = r'''
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
    public int id, parent_id, depth, child_id, role, state, child_count, index;
    public string name, value, error;
}
public class AccResult {
    public long hwnd;
    public int hresult;
    public bool truncated;
    public List<AccNode> nodes = new List<AccNode>();
    public List<string> errors = new List<string>();
    public List<string> columns;  // 列を絞って読んだときの、読んだ列名（行の中の並び順）
}
// 見出し行のあとの行は、指定した列のセルだけを読む（見出し行に指定の列がすべてあるときだけ）。
public class ColumnFilter {
    public HashSet<string> wanted;
    public bool[] keep;
}
public static class HiddenHistoryAcc {
    [DllImport("oleacc.dll")]
    static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out IAccessible result);
    [DllImport("oleacc.dll")]
    static extern int AccessibleChildren(IAccessible container, int start, int count,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] object[] children, out int obtained);

    static void Walk(IAccessible acc, int childId, int parent, int depth, int index, AccResult result, Stopwatch clock, ColumnFilter filter) {
        if (result.nodes.Count >= 50000 || depth > 16 || clock.Elapsed.TotalSeconds > 20) {
            result.truncated=true; return;
        }
        // 列を絞っているときは、データ行の値（全セルをつないだ文字列）とセルの子の数は読まない。
        bool filtered=filter != null && filter.keep != null;
        var node=new AccNode { id=result.nodes.Count, parent_id=parent, depth=depth, child_id=childId, index=index };
        try { node.name=acc.get_accName(childId); } catch {}
        try { node.role=Convert.ToInt32(acc.get_accRole(childId)); } catch {}
        try { node.state=Convert.ToInt32(acc.get_accState(childId)); } catch {}
        if ((node.state & 0x20000000) == 0 && !(filtered && depth == 1)) {
            try { node.value=acc.get_accValue(childId); } catch (Exception ex) { node.error=ex.Message; }
        }
        if (childId == 0 && !(filtered && depth == 2)) { try { node.child_count=acc.accChildCount; } catch {} }
        result.nodes.Add(node);
        if (node.child_count <= 0) return;
        int count=Math.Min(node.child_count, 50000-result.nodes.Count);
        if (count < node.child_count) result.truncated=true;
        object[] children; int obtained; int hr;
        if (filtered && depth == 1 && count == filter.keep.Length) {
            // 読む列のセルだけを取り出す（全セルを取り出すと、それだけでプロセス越しの手間がかかる）。
            children=new object[count]; obtained=count;
            for (int k=0; k<count; k++) {
                if (!filter.keep[k]) continue;
                object[] one=new object[1]; int got;
                hr=AccessibleChildren(acc, k, 1, one, out got);
                if (hr < 0 || got != 1) { result.errors.Add("AccessibleChildren HRESULT: " + hr); return; }
                children[k]=one[0];
            }
        } else {
            children=new object[count];
            hr=AccessibleChildren(acc, 0, count, children, out obtained);
            if (hr < 0) { result.errors.Add("AccessibleChildren HRESULT: " + hr); return; }
        }
        bool skipCells=filtered && depth == 1 && obtained == filter.keep.Length;
        int first=result.nodes.Count;
        for (int i=0; i<obtained; i++) {
            if (clock.Elapsed.TotalSeconds > 20) { result.truncated=true; break; }
            if (skipCells && !filter.keep[i]) continue;
            var child=children[i] as IAccessible;
            if (child != null) Walk(child, 0, node.id, depth+1, i, result, clock, filter);
            else if (children[i] is int) Walk(acc, (int)children[i], node.id, depth+1, i, result, clock, filter);
        }
        if (filter != null && filter.keep == null && depth == 1) {
            // 見出し行（子がすべて列見出し）に指定の列がそろっていれば、以降の行はその列だけ読む。
            var names=new List<string>();
            for (int n=first; n<result.nodes.Count; n++)
                if (result.nodes[n].parent_id == node.id) {
                    if (result.nodes[n].role != 25) { names=null; break; }
                    names.Add(result.nodes[n].name ?? "");
                }
            if (names != null && names.Count == obtained && filter.wanted.IsSubsetOf(names)) {
                filter.keep=new bool[names.Count];
                result.columns=new List<string>();
                for (int n=0; n<names.Count; n++)
                    if (filter.wanted.Contains(names[n])) { filter.keep[n]=true; result.columns.Add(names[n]); }
            }
        }
    }
    public static AccResult Read(long hwnd, string[] columns) {
        var result=new AccResult { hwnd=hwnd };
        try {
            IAccessible acc; Guid iid=new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
            result.hresult=AccessibleObjectFromWindow(new IntPtr(hwnd), 0xfffffffc, ref iid, out acc);
            if (result.hresult >= 0 && acc != null) {
                // Container trees may repeat descendants. Only enumerate native tables.
                var filter=(columns == null || columns.Length == 0) ? null : new ColumnFilter { wanted=new HashSet<string>(columns) };
                if (Convert.ToInt32(acc.get_accRole(0)) == 24)
                    Walk(acc, 0, -1, 0, 0, result, Stopwatch.StartNew(), filter);
            }
        } catch (Exception ex) { result.errors.Add(ex.ToString()); }
        return result;
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location)
$cfg=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__CONFIG__')) | ConvertFrom-Json
$columns=$null
if ($cfg.columns) { $columns=[string[]]@($cfg.columns) }
$results=@(foreach($h in $cfg.handles) { [HiddenHistoryAcc]::Read([long]$h, $columns) })
$results | ConvertTo-Json -Depth 12 -Compress
'''
