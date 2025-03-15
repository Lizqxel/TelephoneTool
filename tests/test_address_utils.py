"""
住所正規化・分割のテストユーティリティ

このモジュールは、住所の正規化と分割のロジックをテストするための
機能を提供します。
"""

import re
from typing import Tuple, Dict, Optional
import jaconv
import pytest

def convert_kansuji_to_number(text: str) -> str:
    """
    漢数字を半角数字に変換する
    
    Args:
        text (str): 変換する文字列
        
    Returns:
        str: 変換後の文字列
    """
    kansuji_map = {
        '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '壱': '1', '弐': '2', '参': '3', '肆': '4', '伍': '5',
        '陸': '6', '漆': '7', '捌': '8', '玖': '9', '拾': '10'
    }
    
    pattern = '|'.join(map(re.escape, kansuji_map.keys()))
    return re.sub(pattern, lambda x: kansuji_map[x.group()], text)

def normalize_address_v2(address: str, preserve_format: bool = False) -> str:
    """
    住所文字列を正規化する

    Args:
        address (str): 正規化する住所文字列
        preserve_format (bool): 元の形式を保持するかどうか

    Returns:
        str: 正規化された住所文字列
    """
    if not address:
        return ""

    # 漢数字を含むかどうかをチェック
    has_kansuji = any(c in address for c in '一二三四五六七八九十壱弐参肆伍陸漆捌玖拾')

    # 全角スペースと半角スペースを削除
    address = re.sub(r'[\s　]+', '', address)

    # 漢数字を半角数字に変換
    kansuji = {
        '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '壱': '1', '弐': '2', '参': '3', '肆': '4', '伍': '5',
        '陸': '6', '漆': '7', '捌': '8', '玖': '9', '拾': '10'
    }
    for k, v in kansuji.items():
        address = address.replace(k, v)

    # 全角数字を半角数字に変換
    zen = "０１２３４５６７８９"
    han = "0123456789"
    trans_table = str.maketrans(zen, han)
    address = address.translate(trans_table)

    # 全角ハイフンを半角ハイフンに変換
    address = re.sub(r'[－−‐⁃‑‒–—﹘―⎯⏤ーｰ─━]', '-', address)

    # 丁目を含む場合は丁目以降の表記を保持
    if "丁目" in address:
        # 丁目の後の番地と号を処理
        match = re.search(r'(\d+)丁目(\d+)番地?(\d+)?号?', address)
        if match:
            chome, ban, go = match.groups()
            if go:
                if preserve_format or has_kansuji:
                    return re.sub(r'(\d+)丁目(\d+)番地?(\d+)号?', f'{chome}丁目{ban}番{go}号', address)
                else:
                    return re.sub(r'(\d+)丁目(\d+)番地?(\d+)号?', f'{chome}丁目{ban}-{go}', address)
            else:
                return re.sub(r'(\d+)丁目(\d+)番地?', f'{chome}丁目{ban}番', address)
        
        # 番地がない場合は丁目までを保持
        match = re.search(r'(\d+)丁目', address)
        if match:
            return address

    # その他の場合はハイフンに統一
    address = re.sub(r'(\d+)(番地?|号|の)(\d+)', r'\1-\3', address)
    address = re.sub(r'(\d+)(番地?|号)(?![0-9])', r'\1', address)
    address = re.sub(r'の(?=\d)', '-', address)

    return address

def split_address_v2(address: str) -> Tuple[str, Optional[str], Optional[str]]:
    """
    住所を分割する関数（バージョン2）
    入力形式：[漢字による住所][数字]-[数字](-[数字])
    例：三重県伊勢市船江4丁目19-10

    Args:
        address (str): 分割する住所

    Returns:
        tuple: (ベース部分, 番地1, 番地2)
        例：('三重県伊勢市船江4丁目', '19', '10')
    """
    if not address:
        return ("", None, None)

    # 基本パターン：[漢字と数字の住所]-[数字]-[数字]
    pattern = r'^(.+?)(\d+)-(\d+)(?:-(\d+))?$'
    match = re.search(pattern, address)
    
    if match:
        base = match.group(1)
        num1 = match.group(2)
        num2 = match.group(3)
        num3 = match.group(4)  # オプショナル
        
        # 3つの数字がある場合（例：1-3-4）
        if num3:
            return (base, num2, num3)
        # 2つの数字がある場合（例：19-10）
        else:
            return (base, num1, num2)
    
    return (address, None, None)

@pytest.mark.parametrize("test_input,expected", [
    # 基本的なケース
    ("東京都新宿区西新宿2-8-1", "東京都新宿区西新宿2-8-1"),
    ("東京都新宿区西新宿２－８－１", "東京都新宿区西新宿2-8-1"),
    
    # スペース処理
    ("東京都 新宿区　西新宿 2-8-1", "東京都新宿区西新宿2-8-1"),
    ("東京都　新宿区　西新宿　２－８－１", "東京都新宿区西新宿2-8-1"),
    
    # 漢数字
    ("東京都新宿区西新宿二丁目八番一号", "東京都新宿区西新宿2丁目8番1号"),
    ("東京都新宿区西新宿弐丁目捌番壱号", "東京都新宿区西新宿2丁目8番1号"),
    
    # 表記ゆれ
    ("東京都新宿区西新宿2丁目8番地1号", "東京都新宿区西新宿2丁目8-1"),
    ("東京都新宿区西新宿2番地8号", "東京都新宿区西新宿2-8"),
    ("東京都新宿区西新宿2の8の1", "東京都新宿区西新宿2-8-1"),
    
    # エッジケース
    ("", ""),
    ("東京都新宿区西新宿", "東京都新宿区西新宿"),
    ("東京都新宿区西新宿2", "東京都新宿区西新宿2"),
])
def test_normalize_address(test_input, expected):
    """住所正規化のテスト"""
    assert normalize_address_v2(test_input, preserve_format=False) == expected

@pytest.mark.parametrize("test_input,expected", [
    # 基本パターン（数字2つ）
    ("三重県伊勢市船江4丁目19-10", ("三重県伊勢市船江4丁目", "19", "10")),
    ("三重県伊勢市大世古4丁目1-36", ("三重県伊勢市大世古4丁目", "1", "36")),
    
    # 基本パターン（数字3つ）
    ("兵庫県川西市萩原台西1-3-4", ("兵庫県川西市萩原台西", "3", "4")),
    
    # 数字なしのケース
    ("三重県伊勢市", ("三重県伊勢市", None, None)),
    
    # 空文字列
    ("", ("", None, None)),
])
def test_split_address(test_input, expected):
    """住所分割のテスト"""
    assert split_address_v2(test_input) == expected

if __name__ == "__main__":
    pytest.main([__file__, "-v"]) 