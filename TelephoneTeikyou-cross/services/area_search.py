"""
NTT西日本の提供エリア検索サービス

このモジュールは、NTT西日本の提供エリア検索を
自動化するための機能を提供します。

主な機能：
- 郵便番号による住所検索
- 住所の自動選択
- 番地・号の入力
- 建物情報の選択
- 提供エリア判定

制限事項：
- キャッシュ機能は使用しません
- エラー発生時は詳細なログを出力します
"""

import logging
import time
import re
import os
import json
from datetime import datetime  # datetime のインポートを追加
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from PIL import Image  # PILライブラリを追加

from services.web_driver import create_driver, load_browser_settings
from utils.string_utils import normalize_string, calculate_similarity
from utils.address_utils import split_address, normalize_address

# グローバル変数でブラウザドライバーを保持
global_driver = None

def is_east_japan(address):
    """
    住所が東日本かどうかを判定する
    
    Args:
        address (str): 判定する住所
        
    Returns:
        bool: 東日本ならTrue、西日本ならFalse
    """
    # 東日本の都道府県リスト
    east_japan_prefectures = [
        "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県"
    ]
    
    # 住所から都道府県を抽出
    prefecture_pattern = r'^(東京都|北海道|(?:京都|大阪)府|.+?県)'
    prefecture_match = re.match(prefecture_pattern, address)
    if not prefecture_match:
        return False
    
    prefecture = prefecture_match.group(1)
    return prefecture in east_japan_prefectures

def normalize_address(address):
    """
    住所文字列を正規化する関数
    
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
                        # ハイフンが2つある場合は、最初の数字を丁目として扱う
                        double_hyphen_match = re.search(r'(\d+)-(\d+)-(\d+)', remaining)
                        if double_hyphen_match:
                            block = double_hyphen_match.group(1)
                            number_part = f"{double_hyphen_match.group(2)}-{double_hyphen_match.group(3)}"
                            town = remaining[:double_hyphen_match.start()].strip()
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
                    'town': town if town else "",  # Noneの代わりに空文字列を返す
                    'block': block,
                    'number': number_part,
                    'building_id': None
                }
            
            return {
                'prefecture': prefecture,
                'city': remaining_address if prefecture else None,
                'town': "",  # Noneの代わりに空文字列を返す
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
        logging.error(f"住所分割中にエラー: {str(e)}")
        return None

def normalize_string(text):
    """
    住所文字列を正規化する
    
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

def extract_base_address(address):
    """
    住所から基本部分（丁目まで）を抽出する
    
    Args:
        address (str): 住所文字列
        
    Returns:
        str: 基本部分の住所
    """
    # 丁目を含む場合は丁目まで抽出
    chome_match = re.search(r'^(.+?[0-9]+丁目)', address)
    if chome_match:
        return chome_match.group(1)
    
    # 数字を含む場合は最初の数字まで抽出
    number_match = re.search(r'^(.+?[0-9]+)', address)
    if number_match:
        return number_match.group(1)
    
    return address

def calculate_address_similarity(input_address, candidate_address):
    """
    住所の類似度を計算する
    
    Args:
        input_address (str): 入力された住所
        candidate_address (str): 候補の住所
        
    Returns:
        float: 類似度（0.0 ~ 1.0）
    """
    # 住所を正規化
    input_normalized = normalize_string(input_address)
    candidate_normalized = normalize_string(candidate_address)
    
    # 都道府県、市区町村、町名に分割
    input_parts = input_normalized.split('市')
    candidate_parts = candidate_normalized.split('市')
    
    if len(input_parts) != 2 or len(candidate_parts) != 2:
        # 市で分割できない場合は単純な類似度を返す
        return calculate_similarity(input_normalized, candidate_normalized)
    
    # 都道府県＋市の部分
    input_city = input_parts[0] + '市'
    candidate_city = candidate_parts[0] + '市'
    
    # 町名部分
    input_town = input_parts[1]
    candidate_town = candidate_parts[1]
    
    # 都道府県＋市の一致度（重み: 0.6）
    city_similarity = 0.6 if input_city == candidate_city else 0.0
    
    # 町名の類似度（重み: 0.4）
    town_similarity = calculate_similarity(input_town, candidate_town) * 0.4
    
    return city_similarity + town_similarity

def is_address_match(input_address, candidate_address):
    """
    入力された住所と候補の住所が一致するかを判定する
    
    Args:
        input_address (str): 入力された住所
        candidate_address (str): 候補の住所
        
    Returns:
        tuple: (bool, float) 一致判定と類似度
    """
    # 両方の住所を正規化
    normalized_input = normalize_string(input_address)
    normalized_candidate = normalize_string(candidate_address)
    
    logging.info(f"住所比較 - 入力: {normalized_input} vs 候補: {normalized_candidate}")
    
    # 完全一致の場合
    if normalized_input == normalized_candidate:
        logging.info("完全一致しました")
        return True, 1.0
    
    # 類似度を計算
    similarity = calculate_address_similarity(normalized_input, normalized_candidate)
    logging.info(f"類似度: {similarity}")
    
    # 類似度が0.8以上なら一致とみなす
    if similarity >= 0.8:
        logging.info(f"十分な類似度（{similarity}）で一致と判定")
        return True, similarity
    
    # 基本的な住所部分の比較
    input_parts = set(normalized_input.split())
    candidate_parts = set(normalized_candidate.split())
    
    # 共通部分の割合を計算
    common_parts = input_parts & candidate_parts
    if len(input_parts) > 0:
        part_similarity = len(common_parts) / len(input_parts)
        logging.info(f"共通部分の割合: {part_similarity}")
        
        if part_similarity >= 0.8:
            logging.info(f"共通部分の割合（{part_similarity}）で一致と判定")
            return True, part_similarity
    
    logging.info(f"不一致と判定（類似度: {similarity}）")
    return False, similarity

def find_best_address_match(input_address, candidates):
    """
    最も一致度の高い住所を見つける
    
    Args:
        input_address (str): 入力された住所
        candidates (list): 候補の住所リスト
        
    Returns:
        tuple: (最も一致する候補, 類似度)
    """
    best_candidate = None
    best_similarity = -1
    
    for candidate in candidates:
        try:
            candidate_text = candidate.text.strip().split('\n')[0]
            _, similarity = is_address_match(input_address, candidate_text)
            
            logging.info(f"候補 '{candidate_text}' の類似度: {similarity}")
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_candidate = candidate
        except Exception as e:
            logging.warning(f"候補の処理中にエラー: {str(e)}")
            continue
    
    if best_candidate and best_similarity >= 0.5:  # 最低限の類似度しきい値
        logging.info(f"最適な候補が見つかりました（類似度: {best_similarity}）: {best_candidate.text.strip()}")
        return best_candidate, best_similarity
    
    return None, best_similarity

def handle_building_selection(driver, progress_callback=None):
    """
    建物選択モーダルの検出とハンドリング
    モーダルが表示されない場合は正常に処理を続行
    """
    try:
        # 建物選択モーダルが表示されているか確認（短い待機時間で）
        modal = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.ID, "buildingNameSelectModal"))
        )
        
        if not modal.is_displayed():
            logging.info("建物選択モーダルは表示されていません - 処理を続行します")
            return None

        logging.info("建物選択モーダルが表示されました（集合住宅判定）")
        if progress_callback:
            progress_callback("集合住宅と判定しました。スクリーンショットを保存します。")
        # スクリーンショットを保存
        screenshot_path = "apartment_detected.png"
        take_full_page_screenshot(driver, screenshot_path)
        logging.info(f"集合住宅判定時のスクリーンショットを保存しました: {screenshot_path}")
        # 判定結果を返す
        return {
            "status": "apartment",
            "message": "集合住宅（アパート・マンション等）",
            "details": {
                "判定結果": "集合住宅",
                "提供エリア": "集合住宅（アパート・マンション等）",
                "備考": "該当住所は集合住宅（アパート・マンション等）です。"
            },
            "screenshot": screenshot_path
        }
            
    except TimeoutException:
        logging.info("建物選択モーダルは表示されていません - 処理を続行します")
        return None
    except Exception as e:
        logging.error(f"建物選択モーダルの処理中にエラー: {str(e)}")
        take_full_page_screenshot(driver, "debug_building_modal_error.png")
        raise

def create_driver(headless=False):
    """
    Chromeドライバーを作成する関数
    
    Args:
        headless (bool): ヘッドレスモードで起動するかどうか
        
    Returns:
        WebDriver: 作成されたドライバーインスタンス
    """
    try:
        options = webdriver.ChromeOptions()
        
        if headless:
            options.add_argument('--headless=new')
            # ヘッドレスモード用の追加設定
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-software-rasterizer')
            
        # 共通の最適化設定
        options.add_argument('--window-size=800,600')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-notifications')
        options.add_argument('--disable-popup-blocking')
        
        # メモリ使用量の最適化
        options.add_argument('--disable-application-cache')
        options.add_argument('--aggressive-cache-discard')
        options.add_argument('--disable-default-apps')
        
        driver = webdriver.Chrome(options=options)
        logging.info(f"Chromeドライバーを作成しました（ヘッドレスモード: {headless}）")
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗: {str(e)}")
        raise

def take_full_page_screenshot(driver, save_path):
    """
    ページ全体のスクリーンショットを取得する（スクロール部分も含む）
    現在のウィンドウサイズを維持したまま撮影します

    Args:
        driver (webdriver): Seleniumのwebdriverインスタンス
        save_path (str): スクリーンショットの保存パス

    Returns:
        str: 保存されたスクリーンショットの絶対パス
    """
    # 元のスクロール位置を保存
    original_scroll = driver.execute_script("return window.pageYOffset;")

    try:
        # ページの先頭にスクロール
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.5)
        
        # ページ全体のサイズを取得
        total_height = driver.execute_script("""
            return Math.max(
                document.documentElement.scrollHeight,
                document.body.scrollHeight,
                document.documentElement.offsetHeight,
                document.body.offsetHeight,
                document.documentElement.clientHeight,
                document.body.clientHeight
            );
        """)
        
        # 現在のビューポートサイズを取得
        viewport_width = driver.execute_script("return window.innerWidth;")
        viewport_height = driver.execute_script("return window.innerHeight;")
        
        logging.info(f"ページ高さ: {total_height}, ビューポートサイズ: 幅={viewport_width}, 高さ={viewport_height}")
        
        # ページ全体が1画面に収まる場合は単純にスクリーンショットを撮影
        if total_height <= viewport_height:
            logging.info("ページ全体が1画面に収まります - 単純スクリーンショット")
            driver.save_screenshot(save_path)
            return os.path.abspath(save_path)
        
        # スクリーンショットを保存するリスト
        screenshots = []
        
        # スクロールの開始位置
        current_position = 0
        overlap = 100  # 画像の重複部分を多めにして継ぎ目を確実に対処
        
        screenshot_count = 0
        while current_position < total_height:
            # 指定位置までスクロール
            driver.execute_script(f"window.scrollTo(0, {current_position});")
            time.sleep(1.0)  # スクロール後の描画を十分に待機
            
            # 一時的なスクリーンショットファイル名
            temp_screenshot = f"temp_screenshot_{screenshot_count}.png"
            
            # スクリーンショットを撮影
            driver.save_screenshot(temp_screenshot)
            screenshots.append({
                'path': temp_screenshot,
                'position': current_position
            })
            
            logging.info(f"スクリーンショット {screenshot_count + 1} を撮影: 位置={current_position}")
            
            # 次のスクロール位置（重複を考慮）
            current_position += viewport_height - overlap
            screenshot_count += 1
            
            # 無限ループ防止
            if screenshot_count > 25:
                logging.warning("スクリーンショット撮影回数が上限に達しました")
                break
        
        # 画像を読み込み
        images = []
        for screenshot in screenshots:
            try:
                img = Image.open(screenshot['path'])
                images.append({
                    'image': img,
                    'position': screenshot['position']
                })
            except Exception as e:
                logging.error(f"画像の読み込みに失敗: {screenshot['path']}, エラー: {str(e)}")
        
        if not images:
            logging.error("有効な画像がありません")
            # フォールバック：通常のスクリーンショット
            driver.save_screenshot(save_path)
            return os.path.abspath(save_path)
        
        # 合成後の画像サイズを計算
        max_width = max(img['image'].width for img in images)
        
        # 実際の合成高さを計算（重複を考慮）
        combined_height = images[0]['image'].height
        for i in range(1, len(images)):
            combined_height += images[i]['image'].height - overlap
        
        # 実際のページ高さを超えないように制限
        combined_height = min(combined_height, total_height)
        
        logging.info(f"合成画像サイズ: 幅={max_width}, 高さ={combined_height}")
        
        # 新しい画像を作成
        combined_image = Image.new('RGB', (max_width, combined_height), 'white')
        
        # 画像を縦に結合
        y_offset = 0
        for i, img_data in enumerate(images):
            img = img_data['image']
            
            # 最初の画像以外は重複部分をカット
            if i > 0:
                img = img.crop((0, overlap, img.width, img.height))
            
            # 最後の部分で画像がはみ出る場合の処理
            if y_offset + img.height > combined_height:
                crop_height = combined_height - y_offset
                if crop_height > 0:
                    img = img.crop((0, 0, img.width, crop_height))
                else:
                    break
            
            combined_image.paste(img, (0, y_offset))
            y_offset += img.height
            
            logging.info(f"画像 {i + 1} を合成: y_offset={y_offset}")
        
        # 最終画像を保存
        combined_image.save(save_path, 'PNG', quality=95)
        combined_image.close()
        
        # 画像オブジェクトを閉じる
        for img_data in images:
            img_data['image'].close()
        
        # 一時ファイルを削除
        for screenshot in screenshots:
            try:
                os.remove(screenshot['path'])
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗: {str(e)}")
        
        logging.info(f"ページ全体のスクリーンショットを保存しました: {save_path}")
        return os.path.abspath(save_path)
        
    except Exception as e:
        logging.error(f"スクリーンショット撮影中にエラー: {str(e)}")
        # エラー時のフォールバック：通常のスクリーンショット
        try:
            driver.save_screenshot(save_path)
            return os.path.abspath(save_path)
        except Exception as fallback_error:
            logging.error(f"フォールバックスクリーンショットも失敗: {str(fallback_error)}")
            raise
        
    finally:
        # スクロール位置を元に戻す
        try:
            driver.execute_script(f"window.scrollTo(0, {original_scroll});")
            time.sleep(0.5)
        except Exception as e:
            logging.warning(f"スクロール位置の復元に失敗: {str(e)}")

def search_service_area(postal_code, address, progress_callback=None):
    """
    提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        progress_callback (callable): 進捗状況を通知するコールバック関数
        
    Returns:
        dict: 検索結果を含む辞書
    """
    # 東日本か西日本かを判定
    if is_east_japan(address):
        logging.info("東日本の提供エリア検索を実行します")
        # 東日本の検索機能を動的にインポート
        from services.area_search_east import search_service_area as search_service_area_east
        return search_service_area_east(postal_code, address, progress_callback)
    else:
        logging.info("西日本の住所が入力されました")
        return {
            "status": "error",
            "message": "申し訳ありませんが、このツールは東日本の住所のみ対応しています。",
            "details": {
                "判定結果": "NG",
                "提供エリア": "未対応エリア",
                "備考": "東日本の住所（北海道、青森県、岩手県、宮城県、秋田県、山形県、福島県、茨城県、栃木県、群馬県、埼玉県、千葉県、東京都、神奈川県、新潟県、富山県、石川県、福井県、山梨県、長野県）のみ対応しています。"
            }
        }

# 西日本の検索機能は無効化
def search_service_area_west(postal_code, address, progress_callback=None):
    """
    この関数は無効化されています。
    東日本の住所のみ対応しています。
    """
    return {
        "status": "error",
        "message": "申し訳ありませんが、このツールは東日本の住所のみ対応しています。",
        "details": {
            "判定結果": "NG",
            "提供エリア": "未対応エリア",
            "備考": "東日本の住所のみ対応しています。"
        }
    }