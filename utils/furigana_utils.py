"""
よみたんAPIを使用してフリガナ変換を行うユーティリティモジュール。

このモジュールは、漢字テキストをカタカナに変換する機能を提供します。
"""

import logging
import requests
from urllib.parse import quote, unquote

def convert_to_furigana(text: str) -> str:
    """
    漢字テキストをカタカナに変換する

    Args:
        text (str): 変換する文字列

    Returns:
        str: カタカナ変換結果。エラー時はNone

    Raises:
        Exception: API呼び出しに失敗した場合
    """
    logging.info(f"変換対象テキスト: {text}")
    
    try:
        # APIのエンドポイントとパラメータ
        url = "https://yomitan.harmonicom.jp/api/v2/yomi"
        params = {
            "ic": "UTF8",
            "oc": "UTF8",
            "kana": "k",
            "text": text,
            "num": 3
        }
        
        # APIリクエスト
        response = requests.get(url, params=params)
        response.raise_for_status()
        
        # レスポンスの解析
        data = response.json()
        logging.info(f"APIレスポンス: {data}")
        
        # 変換結果の取得（最初の候補を使用）
        if "yomi" in data and data["yomi"]:
            result = data["yomi"][0]
            logging.info(f"変換結果: {result}")
            return result
        
        return None
        
    except Exception as e:
        logging.error(f"フリガナ変換エラー: {str(e)}")
        raise Exception(f"フリガナ変換に失敗しました: {str(e)}") 