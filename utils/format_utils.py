"""
フォーマット処理ユーティリティ

このモジュールは、電話番号や郵便番号などの
フォーマット処理に関するユーティリティ関数を提供します。
"""

import re


def format_phone_number(phone_number):
    """
    電話番号を自動フォーマットする関数（ハイフンあり）
    
    Args:
        phone_number (str): フォーマットする電話番号
        
    Returns:
        str: フォーマットされた電話番号
    """
    # 数字以外の文字を削除
    digits_only = re.sub(r'\D', '', phone_number)
    
    # 桁数に応じてフォーマット
    if len(digits_only) == 11:  # 携帯電話（090-xxxx-xxxx）
        return f"{digits_only[0:3]}-{digits_only[3:7]}-{digits_only[7:11]}"
    elif len(digits_only) == 10:  # 固定電話（03-xxxx-xxxx）
        if digits_only.startswith('0'):
            if digits_only.startswith('03') or digits_only.startswith('06'):  # 東京/大阪
                return f"{digits_only[0:2]}-{digits_only[2:6]}-{digits_only[6:10]}"
            else:  # その他の市外局番
                return f"{digits_only[0:3]}-{digits_only[3:6]}-{digits_only[6:10]}"
    
    # フォーマットできない場合は元の値を返す
    return phone_number


def format_phone_number_without_hyphen(phone_number):
    """
    電話番号を自動フォーマットする関数（ハイフンなし）
    
    Args:
        phone_number (str): フォーマットする電話番号
        
    Returns:
        str: フォーマットされた電話番号（ハイフンなし）
    """
    # 数字以外の文字を削除
    return re.sub(r'\D', '', phone_number)


def format_postal_code(postal_code):
    """
    郵便番号を自動フォーマットする関数
    
    Args:
        postal_code (str): フォーマットする郵便番号
        
    Returns:
        str: フォーマットされた郵便番号
    """
    # 数字以外の文字を削除
    digits_only = re.sub(r'\D', '', postal_code)
    
    # 7桁の場合はハイフンを挿入
    if len(digits_only) == 7:
        return f"{digits_only[0:3]}-{digits_only[3:7]}"
    
    # フォーマットできない場合は元の値を返す
    return postal_code


def convert_to_half_width(text):
    """
    全角文字を半角に変換する関数
    
    Args:
        text (str): 変換する文字列
        
    Returns:
        str: 半角に変換された文字列
    """
    if not text:
        return ""
        
    # 全角数字を半角に変換
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    
    # 全角英字を半角に変換
    text = text.translate(str.maketrans('ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',
                                      'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    
    # 全角カッコを半角に変換
    text = text.translate(str.maketrans('（）［］｛｝「」『』【】〔〕', '()[]{}""\'\'<>[]'))
    
    # 全角スペースを半角に変換
    text = text.replace('　', ' ')
    
    return text 