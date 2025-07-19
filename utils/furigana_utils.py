"""
よみたんAPIとpykakasiを使用してフリガナ変換を行うユーティリティモジュール。

このモジュールは、漢字テキストをカタカナに変換する機能を提供します。
よみたんAPIが利用できない場合は、pykakasiを使用してフォールバックします。
"""

import logging
import requests
from urllib.parse import quote, unquote

def convert_to_furigana_with_pykakasi(text: str) -> str:
    """
    pykakasiを使用して漢字テキストをカタカナに変換する

    Args:
        text (str): 変換する文字列

    Returns:
        str: カタカナ変換結果。エラー時はNone
    """
    try:
        import pykakasi
        
        logging.info(f"pykakasiを使用してフリガナ変換を開始: {text}")
        
        # pykakasiの初期化
        kks = pykakasi.kakasi()
        
        # テキストを変換
        result = kks.convert(text)
        
        # ひらがなをカタカナに変換して結合
        furigana = "".join([r['hira'] for r in result])
        
        # ひらがなをカタカナに変換
        import jaconv
        katakana = jaconv.hira2kata(furigana)
        
        logging.info(f"pykakasi変換結果: {katakana}")
        return katakana
        
    except ImportError as e:
        logging.error(f"pykakasiのインポートエラー: {str(e)}")
        return None
    except Exception as e:
        logging.error(f"pykakasiフリガナ変換エラー: {str(e)}")
        return None

def convert_to_furigana(text: str) -> str:
    """
    漢字テキストをカタカナに変換する（よみたんAPI優先、失敗時はpykakasi）

    Args:
        text (str): 変換する文字列

    Returns:
        str: カタカナ変換結果。エラー時はNone
    """
    logging.info(f"変換対象テキスト: {text}")
    
    # まずよみたんAPIを試行
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
        logging.info(f"よみたんAPIレスポンス: {data}")
        
        # 変換結果の取得（最初の候補を使用）
        if "yomi" in data and data["yomi"]:
            result = data["yomi"][0]
            logging.info(f"よみたんAPI変換結果: {result}")
            return result
        
        logging.warning("よみたんAPIから結果が返ってきませんでした。pykakasiにフォールバックします。")
        
    except Exception as e:
        logging.warning(f"よみたんAPI呼び出しエラー: {str(e)}。pykakasiにフォールバックします。")
    
    # よみたんAPIが失敗した場合、pykakasiを使用
    logging.info("pykakasiを使用してフリガナ変換を試行します")
    pykakasi_result = convert_to_furigana_with_pykakasi(text)
    
    if pykakasi_result:
        logging.info(f"pykakasi変換成功: {pykakasi_result}")
        return pykakasi_result
    else:
        logging.error("よみたんAPIとpykakasiの両方でフリガナ変換に失敗しました")
        return None 