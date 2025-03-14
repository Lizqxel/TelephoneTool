"""
住所分割テスト

このスクリプトは、住所分割の正規表現をテストするためのものです。
"""

import re
import sys

def split_address(address):
    """
    住所を分割する関数
    
    Args:
        address (str): 分割する住所
        
    Returns:
        dict: 分割結果を含む辞書（基本住所、番地、号）
    """
    sys.stdout.write(f"処理開始: {address}\n")
    sys.stdout.flush()
    
    # 住所を分割（都道府県、市区町村、町名、番地・号）
    address = address.replace('　', ' ')  # 全角スペースを半角に統一
    
    # 方角を表す文字
    directions = ['東', '西', '南', '北']
    
    # 数字を含む部分を検出する正規表現
    number_pattern = r'([0-9０-９]+)'
    
    # 丁目・条・番地・号などの単位
    units = ['丁目', '条', '番地', '番', '号']
    
    # 住所から数字部分を抽出
    number_matches = list(re.finditer(number_pattern, address))
    sys.stdout.write(f"数字マッチ数: {len(number_matches)}\n")
    sys.stdout.flush()
    
    if not number_matches:
        # 数字が見つからない場合は分割しない
        sys.stdout.write("数字が見つかりません\n")
        sys.stdout.flush()
        return {"base_address": address, "street_number": None, "building_number": None}
    
    # 丁目・条を含む基本住所を特定
    base_end_pos = 0
    for i, match in enumerate(number_matches):
        num = match.group(1)
        start_pos = match.start()
        end_pos = match.end()
        sys.stdout.write(f"数字{i+1}: {num}, 位置: {start_pos}-{end_pos}\n")
        sys.stdout.flush()
        
        # この数字の後に単位（丁目・条など）があるか確認
        for unit in ['丁目', '条']:
            unit_pos = address.find(unit, end_pos)
            if unit_pos == end_pos:  # 数字の直後に単位がある
                # この単位を含む部分までを基本住所に含める
                base_end_pos = end_pos + len(unit)
                sys.stdout.write(f"単位「{unit}」を検出: 基本住所終了位置 = {base_end_pos}\n")
                sys.stdout.flush()
                break
    
    # 基本住所と残りの部分を分離
    if base_end_pos > 0:
        base_address = address[:base_end_pos]
        remaining = address[base_end_pos:]
        sys.stdout.write(f"基本住所（単位あり）: {base_address}\n")
        sys.stdout.write(f"残りの部分: {remaining}\n")
        sys.stdout.flush()
    else:
        # 丁目・条がない場合は最初の数字の前までを基本住所とする
        base_address = address[:number_matches[0].start()]
        remaining = address[number_matches[0].start():]
        sys.stdout.write(f"基本住所（単位なし）: {base_address}\n")
        sys.stdout.write(f"残りの部分: {remaining}\n")
        sys.stdout.flush()
    
    # 番地と号を抽出
    street_number = None
    building_number = None
    
    # 残りの部分から番地・号を抽出
    if remaining:
        # ハイフンで区切られている場合
        if '-' in remaining or 'ー' in remaining or '－' in remaining:
            # ハイフンを統一
            normalized = remaining.replace('－', '-').replace('ー', '-')
            parts = normalized.split('-')
            if len(parts) >= 1:
                street_number = parts[0].strip()
                if len(parts) >= 2:
                    building_number = parts[1].strip()
            print(f"ハイフン区切り: 番地={street_number}, 号={building_number}")
        else:
            # 番地・号などの単位で区切られている場合
            remaining_numbers = list(re.finditer(number_pattern, remaining))
            if remaining_numbers:
                # 最初の数字を番地とする
                street_number = remaining_numbers[0].group(1)
                
                # 2つ目の数字があれば号とする
                if len(remaining_numbers) >= 2:
                    building_number = remaining_numbers[1].group(1)
                print(f"単位区切り: 番地={street_number}, 号={building_number}")
    
    result = {
        "base_address": base_address.strip(),
        "street_number": street_number,
        "building_number": building_number
    }
    print(f"最終結果: {result}")
    return result

def test_address_split(address):
    """
    住所を分割してテストする関数
    
    Args:
        address (str): テスト対象の住所
    """
    print(f"\nテスト対象住所: {address}")
    
    # 住所分割
    result = split_address(address)
    
    print(f"分割結果:")
    print(f"  基本住所: {result['base_address']}")
    print(f"  番地: {result['street_number']}")
    print(f"  号: {result['building_number']}")
    print("-------------------")

# テスト実行
if __name__ == "__main__":
    print("テスト開始")
    test_addresses = [
        "兵庫県川辺郡猪名川町白金2丁目3-31",
        "東京都新宿区西新宿1-1-1",
        "大阪府大阪市中央区大手前2丁目",
        "北海道札幌市中央区北1条西2丁目3番地4号",
        "愛知県名古屋市中区三の丸3-1-2",
        "京都府京都市上京区烏丸通今出川上る相国寺門前町647",
        "福岡県福岡市博多区博多駅前2丁目1番1号",
        "宮城県仙台市青葉区国分町3-7-1"
    ]
    
    for addr in test_addresses:
        test_address_split(addr)
    
    print("テスト終了") 