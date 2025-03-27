"""
住所正規化のテストモジュール

このモジュールは、住所の正規化と分割機能をテストします。
"""

import pytest
from services.area_search import normalize_string, split_address

def test_normalize_string():
    """文字列正規化のテスト"""
    # 全角数字を半角に変換
    assert normalize_string("１２３４５") == "12345"
    
    # 漢数字を半角数字に変換
    assert normalize_string("一丁目二番地三号") == "1丁目2番地3号"
    
    # 全角ハイフンを半角に変換
    assert normalize_string("１－２－３") == "1-2-3"
    
    # スペースの正規化
    assert normalize_string("東京都　港区") == "東京都港区"
    
    # 「大字」「字」の削除
    assert normalize_string("大字山寺字子安町") == "山寺子安町"

def test_split_address():
    """住所分割のテスト"""
    # 基本的な住所
    address = "東京都港区芝公園４丁目２－８"
    result = split_address(address)
    assert result["prefecture"] == "東京都"
    assert result["city"] == "港区"
    assert result["town"] == "芝公園"
    assert result["block"] == "4"
    assert result["number"] == "2-8"
    assert result["building_id"] == ""

    # 大字・字を含む住所
    address = "山形県東村山郡山辺町大字山寺字子安町４３２１"
    result = split_address(address)
    assert result["prefecture"] == "山形県"
    assert result["city"] == "東村山郡山辺町"
    assert result["town"] == "山寺子安町"
    assert result["block"] == ""
    assert result["number"] == "4321"
    assert result["building_id"] == ""

    # 甲・乙を含む住所
    address = "山梨県甲府市下飯田１丁目１番１号甲２０５号室"
    result = split_address(address)
    assert result["prefecture"] == "山梨県"
    assert result["city"] == "甲府市"
    assert result["town"] == "下飯田"
    assert result["block"] == "1"
    assert result["number"] == "1-1"
    assert result["building_id"] == "甲205" 