"""郵便番号から住所候補を取得するための小さなサービス。"""

from __future__ import annotations

from typing import Callable

import requests


ZIPCLOUD_URL = "https://zipcloud.ibsnet.co.jp/api/search"


def lookup_addresses(
    postal_code: str,
    request_get: Callable = requests.get,
) -> list[str]:
    """7桁の郵便番号に対応する住所候補を返す。

    zipcloud は町域ごとに複数の候補を返すため、重複を除いて表示用の
    完全な住所に整形する。通信エラーは呼び出し元が画面に表示できるよう
    ``RuntimeError`` に統一する。
    """
    digits = "".join(character for character in (postal_code or "") if character.isdigit())
    if len(digits) != 7:
        return []

    try:
        response = request_get(ZIPCLOUD_URL, params={"zipcode": digits}, timeout=8)
        response.raise_for_status()
        payload = response.json()
    except requests.RequestException as exc:
        raise RuntimeError("郵便番号から住所を検索できませんでした") from exc
    except ValueError as exc:
        raise RuntimeError("郵便番号検索の応答を読み取れませんでした") from exc

    if payload.get("status") != 200:
        raise RuntimeError(payload.get("message") or "郵便番号から住所を検索できませんでした")

    candidates = []
    for result in payload.get("results") or []:
        address = "".join(
            str(result.get(key) or "") for key in ("address1", "address2", "address3")
        ).strip()
        if address and address not in candidates:
            candidates.append(address)
    return candidates
