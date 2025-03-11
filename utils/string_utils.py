"""
文字列ユーティリティモジュール

このモジュールは、文字列の正規化や類似度計算などの
ユーティリティ関数を提供します。
"""

import re


def normalize_string(text):
    """
    文字列を正規化する関数
    
    Args:
        text (str): 正規化する文字列
    
    Returns:
        str: 正規化された文字列
    """
    if not text:
        return ""
    
    # 全角数字を半角に変換
    text = text.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    
    # 全角ハイフンを半角に変換
    text = text.replace("−", "-").replace("ー", "-").replace("―", "-")
    
    # 全角スペースを半角に変換
    text = text.replace("　", " ")
    
    # 余分な空白とハイフンを削除
    text = re.sub(r'[\s-]+', '', text)
    
    return text.lower()


def calculate_similarity(str1, str2):
    """
    2つの文字列の類似度を計算する関数
    
    Args:
        str1 (str): 比較する文字列1
        str2 (str): 比較する文字列2
    
    Returns:
        float: 類似度（0.0 ~ 1.0）
    """
    if not str1 or not str2:
        return 0.0
    
    # 正規化された文字列が完全一致する場合
    if str1 == str2:
        return 1.0
    
    # str1がstr2に含まれる、またはその逆の場合
    if str1 in str2 or str2 in str1:
        return 0.8
    
    # 数字部分のみを抽出して比較
    numbers1 = re.findall(r'\d+', str1)
    numbers2 = re.findall(r'\d+', str2)
    
    if numbers1 and numbers2:
        # 数字が完全一致
        if numbers1 == numbers2:
            return 0.9
        # 最初の数字が一致
        if numbers1[0] == numbers2[0]:
            return 0.7
    
    return 0.0


def validate_name(text):
    """
    名前が有効かどうかを検証する関数
    
    Args:
        text (str): 検証する名前
        
    Returns:
        bool: 名前が有効な場合はTrue、無効な場合はFalse
    """
    if not text:
        return True  # 空文字列は許可
        
    # 数字を含む場合は無効
    if re.search(r'\d', text):
        return False
        
    return True


def validate_furigana(text):
    """
    フリガナが有効かどうかを検証する関数
    
    Args:
        text (str): 検証するフリガナ
        
    Returns:
        bool: フリガナが有効な場合はTrue、無効な場合はFalse
    """
    if not text:
        return True  # 空文字列は許可
        
    # 数字を含む場合は無効
    if re.search(r'\d', text):
        return False
        
    # カタカナ、ひらがな、スペース以外の文字を含む場合は無効
    if re.search(r'[^\u3040-\u309F\u30A0-\u30FF\s]', text):
        return False
        
    return True 