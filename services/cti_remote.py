"""CTI（別プロセスのWinForms）の日時欄を、利用者が変えたときと同じ通知で書き換える。

日時欄（SysDateTimePick32）へのメッセージは値をポインタで渡すため、プロセスをまたぐと届かない。
CTIのプロセス内に一時的なメモリを確保して値を置き、DTM_SETSYSTEMTIME で表示を変えてから、
DTN_DATETIMECHANGE を届けて WinForms の DateTimePicker.Value を更新させる（CTIは32bit、こちらは64bitでも可）。
WM_NOTIFY は別プロセスへ送れない（アクセス拒否）ため、WinForms の親が子へ転送するときの
WM_REFLECT + WM_NOTIFY を日時欄へ直接送る（利用者が日時を変えたときに WinForms 内部で起きることと同じ）。
"""

import ctypes
from ctypes import wintypes
import platform
import struct

PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04
WM_NOTIFY = 0x004E
WM_REFLECT = 0x2000  # WinForms が子へ転送する通知の番号（WM_USER + 0x1C00）
DTM_GETSYSTEMTIME = 0x1001
DTM_SETSYSTEMTIME = 0x1002
GDT_VALID = 0
DTN_DATETIMECHANGE = -759
SMTO_ABORTIFHUNG = 0x0002


class RemoteMemory:
    """相手プロセス内の読み書きできる一時領域。with を抜けると解放する。"""

    def __init__(self, pid, size=64):
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        k = self.kernel
        k.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k.OpenProcess.restype = wintypes.HANDLE
        k.VirtualAllocEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD]
        k.VirtualAllocEx.restype = ctypes.c_void_p
        k.VirtualFreeEx.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_size_t, wintypes.DWORD]
        k.WriteProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
        k.ReadProcessMemory.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
        k.IsWow64Process.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.BOOL)]
        k.CloseHandle.argtypes = [wintypes.HANDLE]
        self.size = size
        self.handle = k.OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_QUERY_LIMITED_INFORMATION,
                                    False, pid)
        if not self.handle:
            raise RuntimeError(f"CTIのプロセスを開けませんでした（エラー{ctypes.get_last_error()}）。")
        self.address = k.VirtualAllocEx(self.handle, None, size, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        if not self.address:
            error = ctypes.get_last_error()
            k.CloseHandle(self.handle)
            raise RuntimeError(f"CTIのプロセスに一時領域を確保できませんでした（エラー{error}）。")
        wow = wintypes.BOOL()
        k.IsWow64Process(self.handle, ctypes.byref(wow))
        # 32bitのCTI（WOW64）ならポインタは4バイト。64bit Windows上の64bitプロセスなら8バイト。
        self.pointer_size = 4 if wow.value or not platform.machine().endswith("64") else 8

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.kernel.VirtualFreeEx(self.handle, self.address, 0, MEM_RELEASE)
        self.kernel.CloseHandle(self.handle)

    def write(self, data):
        if len(data) > self.size:
            raise ValueError("一時領域より大きいデータは書けません。")
        written = ctypes.c_size_t()
        buffer = ctypes.create_string_buffer(data, len(data))
        if not self.kernel.WriteProcessMemory(self.handle, self.address, buffer, len(data), ctypes.byref(written)) or written.value != len(data):
            raise RuntimeError("CTIのプロセスへ値を書き込めませんでした。")

    def read(self, size):
        buffer = ctypes.create_string_buffer(size)
        done = ctypes.c_size_t()
        if not self.kernel.ReadProcessMemory(self.handle, self.address, buffer, size, ctypes.byref(done)) or done.value != size:
            raise RuntimeError("CTIのプロセスから値を読み取れませんでした。")
        return buffer.raw


def system_time(value):
    """datetime を SYSTEMTIME（16バイト）にする。曜日は日曜=0。"""
    return struct.pack("<8H", value.year, value.month, (value.weekday() + 1) % 7, value.day,
                       value.hour, value.minute, value.second, value.microsecond // 1000)


def date_time_change(pointer_size, picker, value):
    """NMDATETIMECHANGE（NMHDR＋dwFlags＋SYSTEMTIME）を相手プロセスのポインタ幅で作る。"""
    header = (struct.pack("<IIi", picker["hwnd"] & 0xffffffff, picker["control_id"] & 0xffffffff, DTN_DATETIMECHANGE)
              if pointer_size == 4 else
              struct.pack("<QQi4x", picker["hwnd"], picker["control_id"] & 0xffffffffffffffff, DTN_DATETIMECHANGE))
    body = struct.pack("<I", GDT_VALID) + system_time(value)
    return header + body + (b"\0" * 4 if pointer_size == 8 else b"")


def set_date_time(api, pid, picker, value, timeout_ms=5000):
    """日時欄の値を変え、WinForms に変更を知らせる。どちらかが応答しなければ RuntimeError。"""

    def send(hwnd, message, wparam, lparam):
        result = ctypes.c_size_t()
        if not api.SendMessageTimeoutW(hwnd, message, wparam, lparam, SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(result)):
            raise RuntimeError("CTIの検索日時の変更に応答がありません。")
        return result.value

    with RemoteMemory(pid) as memory:
        memory.write(system_time(value))
        if not send(picker["hwnd"], DTM_SETSYSTEMTIME, GDT_VALID, memory.address):
            raise RuntimeError("CTIの検索日時を変更できませんでした。")
        memory.write(date_time_change(memory.pointer_size, picker, value))
        send(picker["hwnd"], WM_REFLECT + WM_NOTIFY, picker["control_id"] & 0xffffffff, memory.address)
