"""HTTP client for the independent performance Apps Script endpoint."""

from __future__ import annotations

from typing import Any, Dict

import requests


class PerformanceSyncError(RuntimeError):
    pass


class PerformanceApiClient:
    def __init__(self, endpoint: str, token: str, timeout: int = 20) -> None:
        self.endpoint = (endpoint or "").strip()
        self.token = token or ""
        self.timeout = max(5, int(timeout))

    @property
    def configured(self) -> bool:
        return self.endpoint.startswith(("https://", "http://127.0.0.1", "http://localhost")) and bool(self.token)

    def sync(self, client_id: str, business_date: str, payload: Dict[str, Any]) -> Dict[str, Any]:
        if not self.configured:
            raise PerformanceSyncError("実績同期先が未設定です")
        body = {
            "action": "sync",
            "token": self.token,
            "clientId": client_id,
            "businessDate": business_date,
            "stats": payload.get("stats", []),
            "details": payload.get("details", []),
        }
        response = requests.post(self.endpoint, json=body, timeout=self.timeout)
        response.raise_for_status()
        data = response.json()
        if not data.get("ok"):
            raise PerformanceSyncError(str(data.get("error") or "実績同期に失敗しました"))
        return data

    def fetch(self, business_date: str) -> Dict[str, Any]:
        if not self.configured:
            raise PerformanceSyncError("実績同期先が未設定です")
        response = requests.post(
            self.endpoint,
            json={"action": "query", "token": self.token, "date": business_date},
            timeout=self.timeout,
        )
        response.raise_for_status()
        data = response.json()
        if not data.get("ok"):
            raise PerformanceSyncError(str(data.get("error") or "実績取得に失敗しました"))
        return data.get("dashboard") or {"date": business_date, "staff": [], "totals": {}}

