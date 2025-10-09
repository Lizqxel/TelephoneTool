"""
フリガナ変換ユーティリティ（pykakasi専用版）。

このモジュールは、漢字テキストをカタカナに変換する機能を提供します。
外部API（よみたん）は使用せず、`pykakasi` のみで変換します。

主な仕様:
- 入力となる任意の文字列を `pykakasi` でかな変換し、`jaconv` でカタカナへ統一
- エラー時は詳細ログを出力し、`None` を返却

制限事項:
- 人名・地名など固有名詞の長音・拗音の揺れは発生しうる
- `pykakasi` の辞書精度に依存
"""

import logging
# import requests  # よみたんAPI無効化に伴い未使用
# from urllib.parse import quote, unquote  # 未使用

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
    漢字テキストをカタカナに変換する（pykakasiのみ使用）

    Args:
        text (str): 変換する文字列

    Returns:
        str: カタカナ変換結果。エラー時はNone
    """
    logging.info(f"変換対象テキスト: {text}")
    
    # よみたんAPI部分（無効化）
    """
    try:
        url = "https://yomitan.harmonicom.jp/api/v2/yomi"
        params = {"ic": "UTF8", "oc": "UTF8", "kana": "k", "text": text, "num": 3}
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        if "yomi" in data and data["yomi"]:
            return data["yomi"][0]
    except Exception:
        pass
    """

    # pykakasi のみを使用
    logging.info("pykakasiのみを使用してフリガナ変換を実行します")
    pykakasi_result = convert_to_furigana_with_pykakasi(text)
    
    if pykakasi_result:
        logging.info(f"pykakasi変換成功: {pykakasi_result}")
        return pykakasi_result
    else:
        logging.error("pykakasiによるフリガナ変換に失敗しました: function=convert_to_furigana, text=%s", text)
        return None 