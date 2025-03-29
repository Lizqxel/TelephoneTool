"""
文字列処理ユーティリティ

このモジュールは、住所や文字列の正規化などの
ユーティリティ機能を提供します。
"""

import re


def normalize_string(text):
    """
    文字列を正規化する
    
    Args:
        text (str): 正規化する文字列
        
    Returns:
        str: 正規化された文字列
    """
    if not text:
        return text
        
    # 全角数字を半角に変換
    normalized = text
    zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
    normalized = normalized.translate(zen_to_han)
    
    # 漢数字を半角数字に変換
    kanji_numbers = {
        '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10'
    }
    for kanji, number in kanji_numbers.items():
        normalized = normalized.replace(kanji, number)
    
    # 全角ハイフンを半角に変換
    normalized = normalized.replace('−', '-').replace('ー', '-').replace('－', '-')
    
    # すべてのスペース（全角・半角）を削除
    normalized = normalized.replace('　', '').replace(' ', '')
    
    # 「大字」「字」を削除
    normalized = normalized.replace('大字', '').replace('字', '')
    
    # 余分な空白を削除
    normalized = normalized.strip()
    
    return normalized


def normalize_address(address):
    """
    住所文字列を正規化する
    
    Args:
        address (str): 正規化する住所文字列
        
    Returns:
        str: 正規化された住所文字列
    """
    # 空白文字の正規化
    address = address.replace('　', ' ').strip()
    
    # ハイフンの正規化
    address = address.replace('−', '-').replace('ー', '-').replace('－', '-')
    
    # 数字の正規化（全角→半角）
    zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
    address = address.translate(zen_to_han)
    
    return address


def split_address(address):
    """
    住所を分割して各要素を抽出する

    Args:
        address (str): 分割する住所
        
    Returns:
        dict: 分割された住所の要素
            - prefecture: 都道府県
            - city: 市区町村
            - town: 町名（大字・字を除く）
            - block: 丁目
            - number: 番地
            - building_id: 建物ID
    """
    try:
        # 都道府県を抽出
        prefecture_match = re.match(r'^(.+?[都道府県])', address)
        prefecture = prefecture_match.group(1) if prefecture_match else None
        block = None
        number_part = None
        town = ""
        
        if prefecture:
            remaining_address = address[len(prefecture):].strip()
            # 市区町村を抽出（郡がある場合も考慮）
            city_match = re.match(r'^(.+?郡.+?[町村]|.+?[市区町村])', remaining_address)
            city = city_match.group(1) if city_match else None
            
            if city:
                # 残りの住所から基本住所と番地を分離
                remaining = remaining_address[len(city):].strip()
                
                # 特殊な表記（大字、字、甲乙丙丁）を含む部分を抽出
                special_location_match = re.search(r'(大字.+?字.*?|大字.*?|字.*?)([甲乙丙丁])?(\d+)', remaining)
                
                if special_location_match:
                    # 番地のみを抽出（甲乙丙丁は除外）
                    number_part = special_location_match.group(3)
                    # 基本住所は市区町村までとする
                    town = ""
                else:
                    # 丁目を含む場合は、丁目の後ろの番地を抽出
                    chome_match = re.search(r'(\d+)丁目', remaining)
                    if chome_match:
                        # 丁目より後ろの部分から番地を探す
                        after_chome = remaining[remaining.find('丁目') + 2:].strip()
                        # ハイフンを含む番地のパターンを優先的に検索
                        number_match = re.search(r'(\d+(?:[-－]\d+)?)', after_chome)
                        if number_match:
                            number_part = number_match.group(1)
                            town = remaining[:remaining.find('丁目') - len(chome_match.group(1))].strip()
                        else:
                            number_part = None
                            town = remaining[:remaining.find('丁目')].strip()
                        block = chome_match.group(1)
                    else:
                        # 通常の番地パターンを検索
                        number_match = re.search(r'(\d+(?:[-－]\d+)?)', remaining)
                        if number_match:
                            number_part = number_match.group(1)
                            town = remaining[:number_match.start()].strip()
                        else:
                            number_part = None
                            town = remaining
                
                return {
                    'prefecture': prefecture,
                    'city': city,
                    'town': town if town else "",
                    'block': block,
                    'number': number_part,
                    'building_id': None
                }
            
            return {
                'prefecture': prefecture,
                'city': remaining_address if prefecture else None,
                'town': "",
                'block': None,
                'number': None,
                'building_id': None
            }
        
        return {
            'prefecture': None,
            'city': None,
            'town': "",
            'block': None,
            'number': None,
            'building_id': None
        }
        
    except Exception as e:
        import logging
        logging.error(f"住所分割中にエラー: {str(e)}")
        return None 


def calculate_similarity(str1, str2):
    """
    2つの文字列の類似度をレーベンシュタイン距離を使用して計算する
    
    Args:
        str1 (str): 比較する文字列1
        str2 (str): 比較する文字列2
        
    Returns:
        float: 類似度（0.0～1.0）
    """
    if not str1 or not str2:
        return 0.0
    
    # 文字列を正規化（全角→半角、大文字→小文字）
    str1 = str1.lower()
    str2 = str2.lower()
    
    # レーベンシュタイン距離を計算
    len1 = len(str1)
    len2 = len(str2)
    
    # 行列を初期化
    matrix = [[0 for _ in range(len2 + 1)] for _ in range(len1 + 1)]
    
    # 行列の初期値を設定
    for i in range(len1 + 1):
        matrix[i][0] = i
    for j in range(len2 + 1):
        matrix[0][j] = j
    
    # レーベンシュタイン距離を計算
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            cost = 0 if str1[i-1] == str2[j-1] else 1
            matrix[i][j] = min(
                matrix[i-1][j] + 1,      # 削除
                matrix[i][j-1] + 1,      # 挿入
                matrix[i-1][j-1] + cost  # 置換
            )
    
    # 類似度を計算（0.0～1.0）
    max_len = max(len1, len2)
    if max_len == 0:
        return 1.0
    
    distance = matrix[len1][len2]
    similarity = 1.0 - (distance / max_len)
    
    return similarity 