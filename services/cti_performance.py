"""Read CTI identity and import privacy-minimised CTI history CSV files."""

from __future__ import annotations

import csv
import hashlib
import io
import logging
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import win32gui
except Exception:  # pragma: no cover - non-Windows test environments
    win32gui = None

try:
    from pywinauto import Desktop
except Exception:  # pragma: no cover - optional fallback during source tests
    Desktop = None


IDENTITY_PATTERN = re.compile(
    r"(?P<name>[\u3000\u3040-\u30ff\u3400-\u9fffA-Za-z0-9 _]+?)さん\s*[:：]\s*(?P<staff_id>[A-Za-z]{1,8}\d{4,})"
)


def parse_identity_text(text: str) -> Optional[Tuple[str, str]]:
    match = IDENTITY_PATTERN.search((text or "").strip())
    if not match:
        return None
    name = match.group("name").replace("管理メニュー", "")
    name = re.sub(r"[\s\u3000]+", " ", name).strip()
    return match.group("staff_id").strip(), name


class CTIIdentityReader:
    def __init__(self) -> None:
        self.window_handle = None

    def _find_window(self) -> Optional[int]:
        if win32gui is None:
            return None
        found = []

        def callback(hwnd, _):
            try:
                if win32gui.IsWindowVisible(hwnd) and "CTIメイン" in win32gui.GetWindowText(hwnd):
                    found.append(hwnd)
                    return False
            except Exception:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        self.window_handle = found[0] if found else None
        return self.window_handle

    @staticmethod
    def _control_text(hwnd: int) -> str:
        try:
            return win32gui.GetWindowText(hwnd) or ""
        except Exception:
            pass
        return ""

    def read_identity(self) -> Optional[Tuple[str, str]]:
        if win32gui is None:
            return None
        hwnd = self.window_handle
        if not hwnd or not win32gui.IsWindow(hwnd):
            hwnd = self._find_window()
        if not hwnd:
            return None
        candidates: List[str] = []
        try:
            title = win32gui.GetWindowText(hwnd)
            if title:
                candidates.append(title)

            def callback(child, _):
                text = self._control_text(child)
                if text:
                    candidates.append(text)
                return True

            win32gui.EnumChildWindows(hwnd, callback, None)
        except Exception as exc:
            logging.debug("CTI担当者情報の取得に失敗: %s", exc)
            self.window_handle = None
            return None
        for text in candidates:
            parsed = parse_identity_text(text)
            if parsed:
                return parsed
        # WinFormsのStatusStripでは氏名とIDが別ラベルになる場合がある。
        parsed = parse_identity_text(" ".join(candidates))
        if parsed:
            return parsed
        return self._read_identity_uia()

    @staticmethod
    def _read_identity_uia() -> Optional[Tuple[str, str]]:
        """Read ToolStripStatusLabel items that do not own a native HWND."""
        if Desktop is None:
            return None
        try:
            window = Desktop(backend="uia").window(title_re=r"^CTI.*")
            if not window.exists(timeout=1):
                return None
            for control in window.descendants():
                name = str(getattr(control.element_info, "name", "") or "")
                if not re.search(r"[A-Za-z]{1,8}\d{4,}", name):
                    continue
                parsed = parse_identity_text(name)
                if parsed:
                    return parsed
        except Exception as exc:
            logging.debug("UI AutomationでのCTI担当者取得に失敗: %s", exc)
        return None

    @staticmethod
    def read_summary() -> Optional[Dict[str, int]]:
        """Read CTI's authoritative counters when the calling tab is visible."""
        if Desktop is None:
            return None
        labels = {
            "総コール数": "calls",
            "有効コール": "dialogues",
            "新規見込": "tosses",
        }
        try:
            window = Desktop(backend="uia").window(title_re=r"^CTI.*")
            if not window.exists(timeout=1):
                return None
            controls = window.descendants()
            text_controls = [
                control for control in controls
                if getattr(control.element_info, "control_type", "") == "Text"
            ]
            result: Dict[str, int] = {}
            for control in text_controls:
                label = str(getattr(control.element_info, "name", "") or "").strip()
                metric = labels.get(label)
                if not metric:
                    continue
                label_rect = control.rectangle()
                candidates = []
                for value_control in text_controls:
                    if value_control is control:
                        continue
                    value_text = str(getattr(value_control.element_info, "name", "") or "").strip()
                    match = re.match(r"^(\d+)(?:\(|$)", value_text)
                    if not match:
                        continue
                    value_rect = value_control.rectangle()
                    horizontal_overlap = min(label_rect.right, value_rect.right) - max(label_rect.left, value_rect.left)
                    vertical_gap = value_rect.top - label_rect.bottom
                    if horizontal_overlap > 0 and -2 <= vertical_gap <= 12:
                        candidates.append((abs(vertical_gap), -horizontal_overlap, int(match.group(1))))
                if candidates:
                    candidates.sort()
                    result[metric] = candidates[0][2]
            return result if len(result) == 3 else None
        except Exception as exc:
            logging.debug("UI AutomationでのCTI集計取得に失敗: %s", exc)
            return None


def _decode_csv(path: str) -> str:
    raw = Path(path).read_bytes()
    for encoding in ("utf-8-sig", "cp932", "shift_jis", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _parse_datetime(value: str) -> Optional[datetime]:
    value = (value or "").strip()
    for fmt in (
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M", "%Y年%m月%d日 %H:%M:%S", "%Y年%m月%d日 %H:%M",
    ):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    return None


def parse_duration_seconds(value: Any) -> int:
    text = str(value or "").strip()
    if not text or text in ("-", "－"):
        return 0
    if re.fullmatch(r"\d+", text):
        return int(text)
    parts = text.split(":")
    try:
        nums = [int(part) for part in parts]
    except ValueError:
        return 0
    if len(nums) == 2:
        return nums[0] * 60 + nums[1]
    if len(nums) == 3:
        return nums[0] * 3600 + nums[1] * 60 + nums[2]
    return 0


def _truthy_metric(value: str) -> bool:
    text = str(value or "").strip()
    return text not in ("", "0", "0.0", "-", "－", "なし", "False", "false")


class CTIHistoryCSVImporter:
    """Parse the CTI history export without retaining customer PII."""

    DATETIME_FIELDS = ("対応日時", "架電日時", "日時")
    STAFF_FIELDS = ("対応スタッフ", "スタッフ", "担当者")
    STAFF_ID_FIELDS = ("スタッフID", "担当者ID", "ログインID")
    DURATION_FIELDS = ("IP電話通話時間", "通話時間")
    RESULT_FIELDS = ("架電結果詳細", "架電結果", "結果")
    DIALOGUE_FIELDS = ("有効コール", "対話")
    TOSS_FIELDS = ("新規見込", "見込")
    UNIQUE_FIELDS = ("管理番号", "対応番号", "電話番号")

    @staticmethod
    def _value(row: Dict[str, str], names: tuple) -> str:
        normalized = {str(key).strip().lstrip("\ufeff"): value for key, value in row.items() if key is not None}
        for name in names:
            if name in normalized:
                return str(normalized[name] or "").strip()
        return ""

    def parse(
        self,
        path: str,
        fallback_staff_id: str = "",
        fallback_staff_name: str = "",
    ) -> List[Dict[str, Any]]:
        text = _decode_csv(path)
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(io.StringIO(text), dialect=dialect)
        records: List[Dict[str, Any]] = []
        for index, row in enumerate(reader, start=2):
            called_at = _parse_datetime(self._value(row, self.DATETIME_FIELDS))
            if called_at is None:
                continue
            staff_name = self._value(row, self.STAFF_FIELDS) or fallback_staff_name or "担当者不明"
            staff_id = self._value(row, self.STAFF_ID_FIELDS) or fallback_staff_id
            if not staff_id:
                staff_id = "name:" + hashlib.sha256(staff_name.encode("utf-8")).hexdigest()[:16]
            talk_seconds = parse_duration_seconds(self._value(row, self.DURATION_FIELDS))
            result_parts = [self._value(row, (name,)) for name in self.RESULT_FIELDS]
            result = " / ".join(dict.fromkeys(part for part in result_parts if part))
            explicit_dialogue = self._value(row, self.DIALOGUE_FIELDS)
            is_dialogue = _truthy_metric(explicit_dialogue) if explicit_dialogue else talk_seconds > 0
            toss_value = self._value(row, self.TOSS_FIELDS)
            is_toss = _truthy_metric(toss_value) or "新規見込" in result
            unique = self._value(row, self.UNIQUE_FIELDS)
            digest_source = "|".join(
                [called_at.isoformat(), staff_id, unique, result, str(talk_seconds), str(index)]
            )
            event_id = "csv-" + hashlib.sha256(digest_source.encode("utf-8")).hexdigest()
            records.append(
                {
                    "event_id": event_id,
                    "business_date": called_at.date().isoformat(),
                    "hour": called_at.hour,
                    "staff_id": staff_id,
                    "staff_name": staff_name,
                    "called_at": called_at.replace(
                        tzinfo=timezone(timedelta(hours=9))
                    ).isoformat(timespec="seconds"),
                    "result": result,
                    "is_dialogue": is_dialogue,
                    "is_toss": is_toss,
                    "talk_seconds": talk_seconds,
                    "source": "cti_history",
                }
            )
        return records

