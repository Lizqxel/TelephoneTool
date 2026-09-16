import pytest

from services.jcom_simulation_service import extract_pricing_from_text


def test_zero_yen_is_not_missing_and_details_are_extracted():
    text = """J:COM NET 光(N)
光(N) 1Gコース
24カ月契約
加入翌月の月額
0円（税込）
3カ月間
料金内訳
基本料金 5,258円（税込）
WEB限定スタート割(NET) 3カ月間 -5,258円（税込）
NET加入特典 キャッシュバック10,000円分 別途エントリーが必要です
"""
    data = extract_pricing_from_text(text)
    assert data["next_month_price_yen"] == 0
    assert data["base_price_yen"] == 5258
    assert data["discount_period_text"] == "3カ月間"
    assert data["contract_period"] == "24カ月契約"
    assert data["course"] == "光(N) 1Gコース"
    assert data["discounts"][0].amount_yen == -5258
    assert "キャッシュバック" in data["benefits_text"]


def test_no_discount_and_unknown_lines_are_not_fabricated():
    data = extract_pricing_from_text(
        "光(N) 1Gコース\n加入翌月の月額 5,808円（税込）\n基本料金 5,808円（税込）\n未知の内訳 777円"
    )
    assert data["next_month_price_yen"] == 5808
    assert data["base_price_yen"] == 5808
    assert data["discounts"] == []


@pytest.mark.parametrize("text", ["", "光(N) 1Gコース ----円"])
def test_loading_or_missing_amount_fails(text):
    with pytest.raises(ValueError):
        extract_pricing_from_text(text)


def test_duplicate_discount_line_is_not_double_counted():
    line = "WEB限定スタート割 6カ月間 -4,142円（税込）"
    data = extract_pricing_from_text(
        f"光(N) 1Gコース\n加入翌月 1,666円\n基本料金 5,808円\n{line}\n{line}"
    )
    assert len(data["discounts"]) == 1


def test_multiline_discount_keeps_sign_and_period():
    data = extract_pricing_from_text(
        "光(N) 1Gコース\n加入翌月 0円\n基本料金 5,258円\n"
        "WEB限定スタート割(NET)\n3カ月間\n-5,258円（税込）"
    )
    assert data["discounts"][0].amount_yen == -5258
    assert data["discounts"][0].period_text == "3カ月間"


def test_amount_missing_is_structured_extraction_failure():
    with pytest.raises(ValueError, match="構造化"):
        extract_pricing_from_text("光(N) 1Gコース\n料金内訳\n金額は取得できません")


def test_current_site_multiline_result_format():
    data = extract_pricing_from_text(
        "シミュレーション結果\n利用サービス\nネット\n"
        "月額利用料金（加入翌月金額）\n5,258円(税込)\n0\n3カ月\n円(税込)/月\n"
        "料金内訳\nネット\n0\n円(税込)\nJ:COM NET 光(N)\n"
        "[24カ月契約]\n[ネット]1Gコース\n[テレビ]なし\n"
        "5,258円(税込)\nWEB限定スタート割(NET)\n[3カ月間]\n-5,258円(税込)\n"
        "その他特典\nキャッシュバック(NET加入特典)\nキャッシュバック\n10,000円分"
    )
    assert data["next_month_price_yen"] == 0
    assert data["base_price_yen"] == 5258
    assert data["contract_period"] == "24カ月契約"
    assert data["discount_period_text"] == "3カ月間"
    assert data["discounts"][0].amount_yen == -5258


def test_au_hikari_1g_result_is_extracted_with_actual_course_name():
    data = extract_pricing_from_text(
        "シミュレーション結果\n利用サービス\nネット\n"
        "月額利用料金（加入翌月金額）\n5,610円(税込)\n3,333\n6カ月\n円(税込)/月\n"
        "料金内訳\n2年契約について\nネット\n光 1Gコース on auひかり\n"
        "5,610円(税込)\nWEB限定スタート割(NET)\n[6カ月間]\n-2,277円(税込)\n"
        "その他特典\nキャッシュバック\n20,000円分"
    )
    assert data["course"] == "光 1Gコース on auひかり"
    assert data["line_type"] == "光 on auひかり"
    assert data["next_month_price_yen"] == 3333
    assert data["base_price_yen"] == 5610
    assert data["contract_period"] == "2年契約"
    assert data["discount_period_text"] == "6カ月間"
    assert data["discounts"][0].amount_yen == -2277
