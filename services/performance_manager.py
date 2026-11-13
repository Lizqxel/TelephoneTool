"""Runtime coordinator for CTI collection, local persistence, and sync."""

from __future__ import annotations

import logging
import threading
import uuid
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Optional, Tuple

from PySide6.QtCore import QObject, QTimer, Signal

from services.cti_performance import CTIHistoryCSVImporter, CTIIdentityReader
from services.performance_store import PerformanceStore
from services.performance_sync import PerformanceApiClient


class PerformanceManager(QObject):
    updated = Signal()
    status_changed = Signal(str)
    identity_changed = Signal(str, str)
    identity_resolved = Signal(object)
    sync_completed = Signal(object, str)

    def __init__(self, settings: Dict[str, Any], cti_monitor=None, parent=None, db_path: str = None) -> None:
        super().__init__(parent)
        self.settings = settings or {}
        self.config = dict(self.settings.get("performanceTracking") or {})
        self.enabled = bool(self.config.get("enabled", True))
        self.store = PerformanceStore(db_path)
        self.client_id = self.config.get("clientId") or self.store.get_or_create_client_id()
        self.cti_monitor = cti_monitor
        self.identity_reader = CTIIdentityReader()
        self.csv_importer = CTIHistoryCSVImporter()
        self.staff_id = str(self.config.get("fallbackStaffId") or "")
        self.staff_name = str(self.config.get("fallbackStaffName") or "")
        self._state_lock = threading.RLock()
        self._current_call: Optional[Dict[str, Any]] = None
        self._last_status = "不明"
        self._identity_tick = 0
        self._identity_read_running = False
        self._sync_running = False
        self._remote_cache: Dict[str, Dict[str, Any]] = {}
        self._last_sync_status = "ローカル保存中"

        endpoint = str(self.config.get("endpoint") or "")
        token = str(self.config.get("sharedToken") or "")
        self.api = PerformanceApiClient(endpoint, token)

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(250)
        self.poll_timer.timeout.connect(self._poll_cti)

        self.sync_timer = QTimer(self)
        self.sync_timer.setInterval(max(10, int(self.config.get("syncIntervalSeconds", 30))) * 1000)
        self.sync_timer.timeout.connect(self.sync_now)
        self.sync_completed.connect(self._on_sync_completed)
        self.identity_resolved.connect(self._apply_identity)

        if self.cti_monitor is not None:
            self.cti_monitor.on_action_button_callback = self.on_cti_action

    def start(self) -> None:
        if not self.enabled:
            self.status_changed.emit("実績機能は設定で無効です")
            return
        self._read_identity()
        self.poll_timer.start()
        self.sync_timer.start()
        QTimer.singleShot(1200, self.sync_now)

    def stop(self) -> None:
        self.poll_timer.stop()
        self.sync_timer.stop()
        with self._state_lock:
            self._finalize_current("アプリ終了")

    def update_settings(self, settings: Dict[str, Any]) -> None:
        self.settings = settings or {}
        self.config = dict(self.settings.get("performanceTracking") or {})
        self.enabled = bool(self.config.get("enabled", True))
        self.api = PerformanceApiClient(
            str(self.config.get("endpoint") or ""),
            str(self.config.get("sharedToken") or ""),
        )
        self.sync_timer.setInterval(max(10, int(self.config.get("syncIntervalSeconds", 30))) * 1000)
        if self.enabled and not self.poll_timer.isActive():
            self.start()
        elif not self.enabled:
            self.stop()

    def set_cti_monitor(self, monitor) -> None:
        self.cti_monitor = monitor
        if monitor is not None:
            monitor.on_action_button_callback = self.on_cti_action

    @staticmethod
    def _status_text(value: Any) -> str:
        return getattr(value, "value", str(value or "不明"))

    def _poll_cti(self) -> None:
        if not self.enabled:
            return
        self._identity_tick += 1
        if self._identity_tick >= 120:
            self._identity_tick = 0
            self._read_identity()
        monitor = self.cti_monitor
        if monitor is None:
            return
        status = self._status_text(getattr(monitor, "current_status", "不明"))
        if status == self._last_status:
            self._finalize_expired_pending()
            return
        previous = self._last_status
        self._last_status = status
        self._handle_status_change(previous, status)

    def _read_identity(self) -> None:
        if self._identity_read_running:
            return
        self._identity_read_running = True

        def worker() -> None:
            try:
                self.identity_resolved.emit(
                    {
                        "identity": self.identity_reader.read_identity(),
                        "summary": self.identity_reader.read_summary(),
                    }
                )
            except Exception:
                logging.exception("CTI担当者取得中にエラー")
                self.identity_resolved.emit(None)

        threading.Thread(target=worker, name="CTIIdentityReaderThread", daemon=True).start()

    def _apply_identity(self, payload) -> None:
        self._identity_read_running = False
        if isinstance(payload, dict):
            identity = payload.get("identity")
            summary = payload.get("summary")
        else:
            identity = payload
            summary = None
        if identity:
            staff_id, staff_name = identity
            if (staff_id, staff_name) != (self.staff_id, self.staff_name):
                old_staff_id, _ = self._identity()
                if old_staff_id.startswith("device:") and old_staff_id != staff_id:
                    today = datetime.now(timezone(timedelta(hours=9))).date().isoformat()
                    self.store.migrate_staff_identity(
                        old_staff_id, staff_id, staff_name, business_date=today
                    )
                self.staff_id, self.staff_name = staff_id, staff_name
                self.identity_changed.emit(staff_id, staff_name)
                logging.info("CTI担当者を検出しました: %s (%s)", staff_name, staff_id)
        if summary and self.staff_id:
            now = datetime.now(timezone(timedelta(hours=9)))
            self.store.set_daily_totals(
                now.date().isoformat(), self.staff_id, self.staff_name,
                summary.get("calls", 0), summary.get("dialogues", 0),
                summary.get("tosses", 0), now.hour,
            )
            self.updated.emit()

    def _identity(self) -> Tuple[str, str]:
        staff_id = self.staff_id or ("device:" + self.client_id[:16])
        staff_name = self.staff_name or "担当者取得待ち"
        return staff_id, staff_name

    def _handle_status_change(self, previous: str, current: str) -> None:
        now = datetime.now(timezone(timedelta(hours=9)))
        with self._state_lock:
            if current == "発信中":
                self._finalize_current("次の架電開始")
                staff_id, staff_name = self._identity()
                self._current_call = {
                    "event_id": "live-" + str(uuid.uuid4()),
                    "started_at": now,
                    "connected_at": None,
                    "ended_at": None,
                    "staff_id": staff_id,
                    "staff_name": staff_name,
                    "result": "",
                    "is_toss": False,
                    "talk_seconds": 0,
                }
                self.store.add_delta(
                    now.date().isoformat(), now.hour, staff_id, staff_name, calls=1, source="live"
                )
                self.updated.emit()
            elif current == "通話中":
                if self._current_call is None:
                    staff_id, staff_name = self._identity()
                    self._current_call = {
                        "event_id": "live-" + str(uuid.uuid4()),
                        "started_at": now,
                        "connected_at": None,
                        "ended_at": None,
                        "staff_id": staff_id,
                        "staff_name": staff_name,
                        "result": "",
                        "is_toss": False,
                        "talk_seconds": 0,
                    }
                    self.store.add_delta(
                        now.date().isoformat(), now.hour, staff_id, staff_name, calls=1, source="live"
                    )
                if self._current_call.get("connected_at") is None:
                    self._current_call["connected_at"] = now
                    start = self._current_call["started_at"]
                    self.store.add_delta(
                        start.date().isoformat(), start.hour,
                        self._current_call["staff_id"], self._current_call["staff_name"],
                        dialogues=1, source="live",
                    )
                    self.updated.emit()
            elif current == "待ち受け中" and self._current_call is not None:
                if self._current_call.get("ended_at") is None:
                    self._current_call["ended_at"] = now
                    connected = self._current_call.get("connected_at")
                    if connected:
                        seconds = max(0, int((now - connected).total_seconds()))
                        self._current_call["talk_seconds"] = seconds
                        start = self._current_call["started_at"]
                        self.store.add_delta(
                            start.date().isoformat(), start.hour,
                            self._current_call["staff_id"], self._current_call["staff_name"],
                            talk_seconds=seconds, source="live",
                        )
                        self.updated.emit()

    def on_cti_action(self, button_name: str) -> None:
        """Called by CTIStatusMonitor from its monitoring thread."""
        with self._state_lock:
            if self._current_call is None:
                return
            self._current_call["result"] = str(button_name or "")
            if button_name == "見込" and not self._current_call.get("is_toss"):
                self._current_call["is_toss"] = True
                start = self._current_call["started_at"]
                self.store.add_delta(
                    start.date().isoformat(), start.hour,
                    self._current_call["staff_id"], self._current_call["staff_name"],
                    tosses=1, source="live",
                )
            self._finalize_current(str(button_name or "結果確定"))
        self.updated.emit()

    def _finalize_expired_pending(self) -> None:
        with self._state_lock:
            if not self._current_call or not self._current_call.get("ended_at"):
                return
            age = (datetime.now(timezone(timedelta(hours=9))) - self._current_call["ended_at"]).total_seconds()
            if age >= 90:
                self._finalize_current("結果未取得")

    def _finalize_current(self, fallback_result: str = "") -> None:
        call = self._current_call
        if call is None:
            return
        started = call["started_at"]
        result = call.get("result") or fallback_result
        self.store.add_call_detail(
            {
                "event_id": call["event_id"],
                "business_date": started.date().isoformat(),
                "hour": started.hour,
                "staff_id": call["staff_id"],
                "staff_name": call["staff_name"],
                "called_at": started.isoformat(timespec="seconds"),
                "result": result,
                "is_dialogue": call.get("connected_at") is not None,
                "is_toss": bool(call.get("is_toss")),
                "talk_seconds": int(call.get("talk_seconds", 0)),
                "source": "live",
            }
        )
        self._current_call = None

    def import_history_csv(self, path: str) -> int:
        self._read_identity()
        with self._state_lock:
            self._finalize_current("CSV取込")
        records = self.csv_importer.parse(path, self.staff_id, self.staff_name)
        count = self.store.replace_history(records)
        with self._state_lock:
            for business_date in {record["business_date"] for record in records}:
                self._remote_cache.pop(business_date, None)
        self.updated.emit()
        self.sync_now()
        return count

    def dashboard(self, business_date: str) -> Dict[str, Any]:
        with self._state_lock:
            cached = self._remote_cache.get(business_date)
        if cached is not None:
            return cached
        return self.store.dashboard(business_date)

    def sync_now(self, business_date: str = None) -> None:
        if not self.enabled or self._sync_running:
            return
        target_date = business_date or datetime.now(timezone(timedelta(hours=9))).date().isoformat()
        if not self.api.configured:
            self._last_sync_status = "ローカル保存中（同期先未設定）"
            self.status_changed.emit(self._last_sync_status)
            self.updated.emit()
            return
        self._sync_running = True
        self.status_changed.emit("同期中…")

        def worker() -> None:
            try:
                payload = self.store.sync_payload(target_date)
                self.api.sync(self.client_id, target_date, payload)
                dashboard = self.api.fetch(target_date)
                self.sync_completed.emit(dashboard, "")
            except Exception as exc:
                logging.warning("実績同期エラー: %s", exc)
                self.sync_completed.emit({}, str(exc))

        threading.Thread(target=worker, name="PerformanceSyncThread", daemon=True).start()

    def _on_sync_completed(self, dashboard: Dict[str, Any], error: str) -> None:
        self._sync_running = False
        if error:
            self._last_sync_status = "同期失敗・ローカル保存済み"
            self.status_changed.emit(self._last_sync_status)
        else:
            business_date = str(
                dashboard.get("date")
                or datetime.now(timezone(timedelta(hours=9))).date().isoformat()
            )
            with self._state_lock:
                self._remote_cache[business_date] = dashboard
            now_jst = datetime.now(timezone(timedelta(hours=9)))
            self._last_sync_status = "同期済み " + now_jst.strftime("%H:%M:%S")
            self.store.set_metadata("last_sync_at", now_jst.isoformat(timespec="seconds"))
            self.status_changed.emit(self._last_sync_status)
        self.updated.emit()

