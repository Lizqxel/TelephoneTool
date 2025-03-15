"""
住所正規化・分割のテスト実行スクリプト
"""

from test_address_utils import test_address_normalization, test_address_splitting

def run_all_tests():
    """
    全てのテストを実行し、結果を表示する
    """
    # 住所正規化のテスト
    print("=== 住所正規化のテスト ===")
    norm_results = test_address_normalization()
    total_norm = len(norm_results)
    success_norm = sum(1 for r in norm_results if r["success"])
    
    for result in norm_results:
        status = "✓" if result["success"] else "✗"
        print(f"{status} 入力: {result['input']}")
        if not result["success"]:
            print(f"  期待値: {result['expected']}")
            print(f"  実際値: {result['actual']}")
    
    print(f"\n正規化テスト結果: {success_norm}/{total_norm} 成功")
    
    # 住所分割のテスト
    print("\n=== 住所分割のテスト ===")
    split_results = test_address_splitting()
    total_split = len(split_results)
    success_split = sum(1 for r in split_results if r["success"])
    
    for result in split_results:
        status = "✓" if result["success"] else "✗"
        print(f"{status} 入力: {result['input']}")
        if not result["success"]:
            print(f"  期待値: {result['expected']}")
            print(f"  実際値: {result['actual']}")
    
    print(f"\n分割テスト結果: {success_split}/{total_split} 成功")
    
    # 総合結果
    total = total_norm + total_split
    success = success_norm + success_split
    print(f"\n=== 総合結果 ===")
    print(f"全体: {success}/{total} 成功 ({success/total*100:.1f}%)")

if __name__ == "__main__":
    run_all_tests() 