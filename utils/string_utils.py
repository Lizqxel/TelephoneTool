"""
文字列処理ユーティリティ

このモジュールは、文字列の正規化や類似度計算などの
ユーティリティ関数を提供します。
"""

import re


def normalize_string(text):
    """
    文字列を正規化する関数 (小文字に変換、句読点を削除、全角数字を半角に変換)
    
    Args:
        text (str): 正規化する文字列
        
    Returns:
        str: 正規化された文字列
    """
    if not text:
        return ""
        
    # 全角数字を半角に変換
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    
    # 小文字に変換
    text = text.lower()
    
    # 句読点、記号を削除
    text = re.sub(r'[^\w\s]', '', text)
    
    # 空白を削除
    text = text.replace(" ", "")
    
    return text


def calculate_similarity(str1, str2):
    """
    2つの文字列の類似度を計算する関数
    
    Args:
        str1 (str): 比較する文字列1
        str2 (str): 比較する文字列2
        
    Returns:
        float: 類似度（0〜100の範囲）
    """
    # 両方の文字列を正規化
    s1 = normalize_string(str1)
    s2 = normalize_string(str2)
    
    if not s1 or not s2:
        return 0
    
    # 部分文字列の一致を確認（住所の一部が含まれているかどうか）
    if s1 in s2 or s2 in s1:
        # 一方がもう一方に含まれている場合は高いスコアを返す
        min_len = min(len(s1), len(s2))
        max_len = max(len(s1), len(s2))
        return 80 + 20 * (min_len / max_len)  # 80〜100の範囲
    
    # 共通する文字の数をカウント
    common_chars = sum(min(s1.count(c), s2.count(c)) for c in set(s1 + s2))
    
    # 類似度を計算（0〜100の範囲）
    total_chars = len(s1) + len(s2)
    if total_chars == 0:
        return 0
    
    # 共通文字の割合を計算
    similarity = (2 * common_chars / total_chars) * 100
    
    # 都道府県名が一致する場合はボーナスを追加
    prefectures = ["北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", 
                  "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", 
                  "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", 
                  "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", 
                  "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", 
                  "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", 
                  "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]
    
    for pref in prefectures:
        if pref in str1 and pref in str2:
            similarity += 20  # 都道府県が一致する場合はボーナス
            break
    
    # 数字が一致する場合もボーナスを追加（郵便番号や番地の一致）
    digits1 = ''.join(c for c in str1 if c.isdigit())
    digits2 = ''.join(c for c in str2 if c.isdigit())
    
    if digits1 and digits2:
        digit_similarity = sum(1 for a, b in zip(digits1, digits2) if a == b) / max(len(digits1), len(digits2))
        similarity += digit_similarity * 10  # 数字の一致度に応じてボーナス
    
    return min(similarity, 100)  # 最大100に制限 