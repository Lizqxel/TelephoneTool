import pytest

from services.postal_code_lookup import lookup_addresses


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return self.payload


def test_lookup_addresses_formats_and_deduplicates_candidates():
    seen = {}

    def request_get(url, params, timeout):
        seen.update(url=url, params=params, timeout=timeout)
        return FakeResponse(
            {
                "status": 200,
                "results": [
                    {"address1": "東京都", "address2": "千代田区", "address3": "千代田"},
                    {"address1": "東京都", "address2": "千代田区", "address3": "千代田"},
                    {"address1": "東京都", "address2": "千代田区", "address3": "丸の内"},
                ],
            }
        )

    assert lookup_addresses("100-0001", request_get) == [
        "東京都千代田区千代田",
        "東京都千代田区丸の内",
    ]
    assert seen["params"] == {"zipcode": "1000001"}


def test_lookup_addresses_rejects_incomplete_postal_code():
    assert lookup_addresses("100-00") == []


def test_lookup_addresses_exposes_api_error():
    with pytest.raises(RuntimeError, match="該当する住所"):
        lookup_addresses(
            "1000001",
            lambda *args, **kwargs: FakeResponse({"status": 400, "message": "該当する住所がありません"}),
        )
