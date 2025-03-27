"""
住所分割機能のテスト
"""

from services.area_search import split_address

def test_address(address):
    """住所分割のテスト関数"""
    print(f"\nテスト対象住所: {address}")
    print(f"処理開始: {address}")
    
    # 住所を分割
    result = split_address(address)
    print("分割結果:")
    print(f"  都道府県: {result['prefecture']}")
    print(f"  市区町村: {result['city']}")
    print(f"  町名: {result['town']}")
    if result['block']:
        print(f"  丁目: {result['block']}")
    print(f"  番地: {result['number']}")
    if result['building_id']:
        print(f"  建物ID: {result['building_id']}")
    print("-" * 20)

def main():
    """テストを実行する関数"""
    print("\nテスト開始\n")
    
    # テストケース
    test_addresses = [
        "兵庫県川辺郡猪名川町白金2丁目3-31",
        "東京都新宿区西新宿1-1-1",
        "大阪府大阪市中央区大手前2丁目",
        "北海道札幌市中央区北1条西2丁目3番地4号",
        "愛知県名古屋市中区三の丸3-1-2",
        "京都府京都市上京区烏丸通今出川上る相国寺門前町647",
        "福岡県福岡市博多区博多駅前2丁目1番1号",
        "宮城県仙台市青葉区国分町3-7-1",
        "奈良県奈良市富雄北1-7-15",
        "東京都千代田区丸の内1丁目",
        "大阪府大阪市北区梅田1丁目2番3号",
        "京都府京都市中京区寺町通御池上る上本能寺前町488",
        "北海道札幌市中央区北1条西2-3",
        "沖縄県那覇市おもろまち1-2-3",
        "奈良県吉野郡下市町大字新住甲123",
        "奈良県吉野郡下市町字新住乙456",
        "奈良県吉野郡下市町大字新住字新田甲789",
        "奈良県吉野郡下市町字新住字新田乙101",
        "奈良県吉野郡下市町大字新住甲123番地",
        "奈良県吉野郡下市町字新住乙456番地",
        "三重県伊勢市宇治浦田3丁目14-36"  # 新しいテストケースを追加
    ]
    
    for address in test_addresses:
        test_address(address)
    
    print("テスト終了")

if __name__ == "__main__":
    main() 