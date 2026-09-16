from services.jcom_simulation_models import AddressCandidate
from utils.jcom_address_matcher import (
    MISSING_ADDRESS,
    NEXT_WITH_CURRENT,
    choose_address_candidate,
    classify_special_candidate,
    confirmed_address_matches,
    normalize_address_for_match,
    split_address_components,
)


ADDRESS = "北海道函館市富岡町３丁目２７番地１３号コーポＳＡＨ２０６号"


def candidates(*texts):
    return [AddressCandidate(str(index), text) for index, text in enumerate(texts)]


def test_example_address_is_selected_by_stage():
    decision = choose_address_candidate(ADDRESS, "", candidates("２丁目", "３丁目", "４丁目"))
    assert decision.candidate_id == "1"
    decision = choose_address_candidate(ADDRESS, "３丁目", candidates("２７", "１２７"))
    assert decision.candidate_id == "0"


def test_existing_area_search_order_is_used_for_address_components():
    assert split_address_components(ADDRESS) == (
        "富岡町",
        "3丁目",
        "27",
        "13",
        "コーポSAH",
        "206号",
    )


def test_component_hints_select_the_current_stage_without_prompt():
    hints = split_address_components(ADDRESS)
    decision = choose_address_candidate(
        ADDRESS, "", candidates("２丁目", "３丁目", "１３"), hints
    )
    assert decision.candidate_id == "1"
    decision = choose_address_candidate(
        ADDRESS, "３丁目", candidates("２７", "１２７", "２０６号"), hints
    )
    assert decision.candidate_id == "0"


def test_current_site_area_count_and_accordion_labels_match_components():
    address = "埼玉県所沢市林１丁目３５−１２"
    hints = split_address_components(address)
    decision = choose_address_candidate(
        address,
        "",
        candidates(
            "所沢市林１丁目の物件(4件)",
            "所沢市林２丁目の物件(2件)",
        ),
        hints,
    )
    assert decision.candidate_id == "0"

    selected = "所沢市林１丁目の物件(4件)"
    decision = choose_address_candidate(
        address, selected, candidates("３５番地", "３５０番地"), hints
    )
    assert decision.candidate_id == "0"

    selected += "|３５番地"
    decision = choose_address_candidate(
        address, selected, candidates("テスト１２号棟", "テスト１３号棟"), hints
    )
    assert decision.candidate_id == "0"
    decision = choose_address_candidate(ADDRESS, "３丁目２７", candidates("３", "１３"))
    assert decision.candidate_id == "1"
    decision = choose_address_candidate(ADDRESS, "３丁目|２７|１３", candidates("コーポＳＡＨ", "コーポSAB"))
    assert decision.candidate_id == "0"
    decision = choose_address_candidate(ADDRESS, "３丁目|２７|１３|コーポＳＡＨ", candidates("２０６号", "２０６０号"))
    assert decision.candidate_id == "0"


def test_full_width_latin_and_spaces_match_but_long_mark_is_preserved():
    assert normalize_address_for_match(" コーポＳＡＨ ") == "コーポsah"
    assert normalize_address_for_match("コーポ") != normalize_address_for_match("コポ")


def test_numeric_components_do_not_confuse_3_and_13():
    decision = choose_address_candidate("函館市3番13号", "函館市3番", candidates("3", "13"))
    assert decision.candidate_id == "1"


def test_kanji_place_name_is_not_converted():
    assert "四日市" in normalize_address_for_match("三重県四日市市西新地")


def test_ambiguous_and_no_candidate_require_user():
    assert choose_address_candidate("甲乙", "", candidates("甲", "甲")).requires_user
    assert choose_address_candidate("甲乙", "", candidates("丙", "丁")).requires_user


def test_next_with_current_is_never_auto_selected():
    special = "【表示中の住所】で次へ"
    assert classify_special_candidate(special) == NEXT_WITH_CURRENT
    decision = choose_address_candidate(ADDRESS, "３丁目", candidates(special))
    assert decision.requires_user


def test_missing_property_link_is_not_candidate():
    decision = choose_address_candidate(ADDRESS, "", candidates("住所・物件がない方はこちら"))
    assert decision.requires_user
    current_site_text = "住所および物件が見つからないお客さまへ"
    assert classify_special_candidate(current_site_text) == MISSING_ADDRESS
    assert classify_special_candidate("該当する住所がない方はこちら") == MISSING_ADDRESS


def test_designated_city_ward_is_not_left_in_town_component():
    assert split_address_components("千葉県千葉市若葉区中田町１１７１−３") == (
        "中田町",
        "1171",
        "3",
    )


def test_confirmed_address_accepts_site_added_zero_and_building_name():
    original = "東京都八王子市寺田町４３２−１０６"
    confirmed = "八王子市寺田町432番地0号グリーンヒル寺田１０６号棟"
    hints = split_address_components(original)
    assert confirmed_address_matches(original, confirmed, hints)
    assert not confirmed_address_matches(
        original,
        "八王子市寺田町432番地0号グリーンヒル寺田１０７号棟",
        hints,
    )


def test_chome_less_and_hyphen_address():
    decision = choose_address_candidate("東京都新宿区西新宿2-8-1", "東京都新宿区西新宿", candidates("2", "20"))
    assert decision.candidate_id == "0"
