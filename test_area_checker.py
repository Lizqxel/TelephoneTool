"""
NTT西日本フレッツ光の提供エリア確認モジュールのテスト

このスクリプトは、area_checker.pyモジュールをテストするためのものです。
指定された住所でエリア確認を実行し、結果を表示します。
"""

import logging
import sys
import traceback
from area_checker import AreaChecker

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)  # 標準出力にログを表示
    ]
)

def test_area_checker():
    """エリア確認機能のテスト"""
    # テスト対象の住所
    prefecture = "山口県"
    city = "山口市"
    town = "阿知須"
    block = "3499"
    
    print(f"住所 {prefecture}{city}{town}{block} のエリア確認を開始します")
    
    try:
        # エリア確認クラスのインスタンス化
        area_checker = AreaChecker()
        
        # エリア確認実行
        result = area_checker.check_area(prefecture, city, town, block)
        
        # 結果を表示
        print("\n===== エリア確認結果 =====")
        print(f"ステータス: {result.get('status', 'Unknown')}")
        print(f"メッセージ: {result.get('message', 'No message')}")
        print("=========================\n")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        print("\n===== エラーの詳細 =====")
        traceback.print_exc()
        print("=========================\n")

if __name__ == "__main__":
    test_area_checker() 