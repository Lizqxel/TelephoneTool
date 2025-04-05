"""
提供エリア検索サービス

このモジュールは、NTT西日本の提供エリア検索を
自動化するための機能を提供します。
"""

import logging
import time
import re
import os
import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from PIL import Image  # PILライブラリを追加

from services.web_driver import create_driver
from utils.string_utils import normalize_string, calculate_similarity
from services.area_search_east import search_service_area as search_service_area_east

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
        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
        "静岡県", "愛知県", "三重県"
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

def handle_building_selection(driver):
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
            return
            
        logging.info("建物選択モーダルが表示されました")
        
        # 「該当する建物名がない」リンクを探して選択
        try:
            no_building_link = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "li.not_adress a"))
            )
            logging.info("「該当する建物名がない」リンクを検出しました")
            
            # クリックを試行
            try:
                no_building_link.click()
                logging.info("通常のクリックで「該当する建物名がない」を選択しました")
            except Exception as click_error:
                logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                try:
                    driver.execute_script("arguments[0].click();", no_building_link)
                    logging.info("JavaScriptでクリックしました")
                except Exception as js_error:
                    logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                    ActionChains(driver).move_to_element(no_building_link).click().perform()
                    logging.info("ActionChainsでクリックしました")
            
            # クリック後の待機
            time.sleep(2)
            
            # モーダルが閉じられるのを待機
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "buildingNameSelectModal"))
            )
            logging.info("建物選択モーダルが閉じられました")
            
        except Exception as e:
            logging.error(f"「該当する建物名がない」の選択に失敗: {str(e)}")
            driver.save_screenshot("debug_no_building_error.png")
            raise
            
    except TimeoutException:
        logging.info("建物選択モーダルは表示されていません - 処理を続行します")
    except Exception as e:
        logging.error(f"建物選択モーダルの処理中にエラー: {str(e)}")
        driver.save_screenshot("debug_building_modal_error.png")
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

    Args:
        driver (webdriver): Seleniumのwebdriverインスタンス
        save_path (str): スクリーンショットの保存パス

    Returns:
        str: 保存されたスクリーンショットの絶対パス
    """
    # 元のウィンドウサイズと位置、スクロール位置を保存
    original_size = driver.get_window_size()
    original_position = driver.get_window_position()
    original_scroll = driver.execute_script("return window.pageYOffset;")

    try:
        # ページ全体のサイズを取得
        total_width = driver.execute_script("return Math.max(document.documentElement.scrollWidth, document.body.scrollWidth);")
        total_height = driver.execute_script("return Math.max(document.documentElement.scrollHeight, document.body.scrollHeight);")
        
        # ビューポートの高さを取得
        viewport_width = driver.execute_script("return window.innerWidth;")
        viewport_height = driver.execute_script("return window.innerHeight;")
        
        # スクリーンショットを保存するリスト
        screenshots = []
        
        # スクロールの開始位置
        current_position = 0
        
        # ウィンドウサイズを設定（幅はビューポートに合わせる）
        driver.set_window_size(viewport_width, viewport_height)
        
        while current_position < total_height:
            # 指定位置までスクロール
            driver.execute_script(f"window.scrollTo(0, {current_position});")
            time.sleep(0.5)  # スクロール後の描画を待機
            
            # 一時的なスクリーンショットファイル名
            temp_screenshot = f"temp_screenshot_{current_position}.png"
            
            # スクリーンショットを撮影
            driver.save_screenshot(temp_screenshot)
            screenshots.append(temp_screenshot)
            
            # 次のスクロール位置（ビューポートの高さ分）
            current_position += viewport_height
        
        # スクリーンショットを合成
        images = [Image.open(s) for s in screenshots]
        
        # 合成後の画像サイズを計算
        max_width = max(img.width for img in images)
        total_height = min(sum(img.height for img in images), total_height)  # ページの実際の高さを超えないように
        
        # 新しい画像を作成
        combined_image = Image.new('RGB', (max_width, total_height))
        
        # 画像を縦に結合
        y_offset = 0
        for img in images:
            # 最後の画像の場合、はみ出る部分をトリミング
            if y_offset + img.height > total_height:
                crop_height = total_height - y_offset
                img = img.crop((0, 0, img.width, crop_height))
            
            combined_image.paste(img, (0, y_offset))
            y_offset += img.height
            
            # 一時ファイルを削除
            img.close()
        
        # 最終画像を保存
        combined_image.save(save_path)
        combined_image.close()
        
        # 一時ファイルを削除
        for screenshot in screenshots:
            try:
                os.remove(screenshot)
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗: {str(e)}")
        
        return os.path.abspath(save_path)
        
    finally:
        # ウィンドウサイズと位置、スクロール位置を元に戻す
        driver.set_window_size(original_size['width'], original_size['height'])
        driver.set_window_position(original_position['x'], original_position['y'])
        driver.execute_script(f"window.scrollTo(0, {original_scroll});")

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
        return search_service_area_east(postal_code, address, progress_callback)
    else:
        logging.info("西日本の提供エリア検索を実行します")
        return search_service_area_west(postal_code, address, progress_callback)

def search_service_area_west(postal_code, address, progress_callback=None):
    """
    NTT西日本の提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        progress_callback (callable): 進捗状況を通知するコールバック関数
        
    Returns:
        dict: 検索結果を含む辞書
    """
    global global_driver
    
    # デバッグログ：入力値の確認
    logging.info(f"=== 検索開始 ===")
    logging.info(f"入力郵便番号（変換前）: {postal_code}")
    logging.info(f"入力住所（変換前）: {address}")
    
    # 郵便番号と住所の正規化
    try:
        postal_code = normalize_address(postal_code)
        address = normalize_address(address)
        logging.info(f"正規化後郵便番号: {postal_code}")
        logging.info(f"正規化後住所: {address}")
    except Exception as e:
        logging.error(f"正規化処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"住所の正規化に失敗しました: {str(e)}"}

    # 住所を分割
    if progress_callback:
        progress_callback("住所情報を解析中...")
    
    address_parts = split_address(address)
    if not address_parts:
        return {"status": "error", "message": "住所の分割に失敗しました。"}
        
    # 基本住所を構築（番地と号を除く）
    base_address = f"{address_parts['prefecture']}{address_parts['city']}{address_parts['town']}"
    if address_parts['block']:
        base_address += f"{address_parts['block']}丁目"
    
    logging.info(f"住所分割結果 - 基本住所: {base_address}, 番地: {address_parts['number']}")
    
    # 番地と号を分離
    street_number = None
    building_number = None
    if address_parts['number']:
        parts = address_parts['number'].split('-')
        street_number = parts[0]
        building_number = parts[1] if len(parts) > 1 else None
    
    # 郵便番号のフォーマットチェック
    postal_code_clean = postal_code.replace("-", "")
    if len(postal_code_clean) != 7 or not postal_code_clean.isdigit():
        return {"status": "error", "message": "郵便番号は7桁の数字で入力してください。"}
    
    # ブラウザ設定を読み込む
    browser_settings = {}
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                browser_settings = settings.get("browser_settings", {})
                logging.info(f"ブラウザ設定を読み込みました: {browser_settings}")
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
        browser_settings = {
            "headless": False,
            "show_popup": True,
            "auto_close": True,
            "page_load_timeout": 60,
            "script_timeout": 60
        }
    
    # ヘッドレスモードの設定を取得
    headless_mode = browser_settings.get("headless", False)
    # ポップアップ表示設定を取得
    show_popup = browser_settings.get("show_popup", True)
    # ブラウザ自動終了設定を取得
    auto_close = browser_settings.get("auto_close", True)
    # タイムアウト設定を取得
    page_load_timeout = browser_settings.get("page_load_timeout", 60)
    script_timeout = browser_settings.get("script_timeout", 60)
    
    logging.info(f"ブラウザ設定 - ヘッドレス: {headless_mode}, ポップアップ表示: {show_popup}, 自動終了: {auto_close}")
    
    driver = None
    try:
        # ドライバーを作成してサイトを開く
        driver = create_driver(headless=headless_mode)
        
        # グローバル変数に保存
        global_driver = driver
        
        # タイムアウト設定を適用
        driver.set_page_load_timeout(page_load_timeout)
        driver.set_script_timeout(script_timeout)
        
        driver.implicitly_wait(0)  # 暗黙の待機を無効化
        
        # サイトにアクセス
        driver.get("https://flets-w.com/cart/")
        logging.info("NTT西日本のサイトにアクセスしています...")
        
        # 2. 郵便番号を入力
        try:
            # 郵便番号入力フィールドが操作可能になるまで待機
            zip_field = WebDriverWait(driver, 10).until(
                lambda d: d.find_element(By.XPATH, "//*[@id='id_tak_tx_ybk_yb']") if d.execute_script("return document.readyState") == "complete" else None
            )
            
            if not zip_field:
                raise TimeoutException("郵便番号入力フィールドが見つかりませんでした")
            
            # フィールドが表示され、操作可能になるまで短い間隔で確認
            for _ in range(10):  # 最大1秒間試行
                try:
                    zip_field.clear()
                    zip_field.send_keys(postal_code_clean)
                    logging.info(f"郵便番号 {postal_code_clean} を入力しました")
                    break
                except:
                    time.sleep(0.1)
                    continue
            
        except Exception as e:
            logging.error(f"郵便番号入力フィールドの操作に失敗: {str(e)}")
            raise
        
        # 3. 検索ボタンを押す
        try:
            search_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_bt_ybk_jks']"))
            )
            search_button.click()
            logging.info("検索ボタンをクリックしました")
        except Exception as e:
            logging.error(f"検索ボタンの操作に失敗: {str(e)}")
            raise
        
        # 住所候補が表示されるのを待つ（最大10秒）
        try:
            if progress_callback:
                progress_callback("基本住所の候補を検索中...")
            
            # 住所選択モーダルが表示されるまで待機
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "addressSelectModal"))
            )
            logging.info("住所選択モーダルが表示されました")
            
            # モーダル内のリストが表示されるまで待機
            WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#addressSelectModal ul li a"))
            )
            
            # 候補リストを取得（最適化された方法）
            candidates = driver.find_elements(By.CSS_SELECTOR, "#addressSelectModal ul li a")
            logging.info(f"{len(candidates)} 件の候補が見つかりました")
            
            # 候補の内容をログ出力
            for i, candidate in enumerate(candidates[:5]):  # 最初の5件のみログ出力
                logging.info(f"候補 {i+1}: {candidate.text.strip()}")
            
            # 有効な候補（テキストが空でない）をフィルタリング
            valid_candidates = [c for c in candidates if c.text.strip()]
            
            if not valid_candidates:
                raise NoSuchElementException("有効な住所候補が見つかりませんでした")
            
            logging.info(f"有効な候補数: {len(valid_candidates)}")
            
            # 住所候補が多い場合は、検索フィールドで絞り込み
            try:
                search_field = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='絞り込みワードを入力']"))
                )
                logging.info("住所検索フィールドが見つかりました")
                
                # 検索用の住所フォーマットを作成
                search_address = base_address.replace("県", "県 ").replace("市", "市 ").replace("町", "町 ").replace("区", "区 ").strip()
                logging.info(f"検索用にフォーマットされた住所: {search_address}")
                
                # 検索フィールドをクリアして入力
                search_field.clear()
                search_field.send_keys(search_address)
                logging.info(f"検索フィールドに '{search_address}' を入力しました")
                
                # 入力後の表示更新を待機（短縮）
                WebDriverWait(driver, 3).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//*[@id='addressSelectModal']//div[contains(@class, 'clickable')]"))
                )
                
                # 絞り込み後の候補を取得
                filtered_candidates = driver.find_elements(By.XPATH, "//*[@id='addressSelectModal']//div[contains(@class, 'clickable')]")
                if filtered_candidates:
                    valid_candidates = filtered_candidates
                    logging.info(f"絞り込み後の候補数: {len(valid_candidates)}")
                    
                    # 絞り込み後の候補をログ出力（最初の5件のみ）
                    for i, candidate in enumerate(valid_candidates[:5]):
                        logging.info(f"絞り込み後の候補 {i+1}: '{candidate.text.strip()}'")
                else:
                    logging.warning("絞り込み後の候補が見つかりませんでした")
                    logging.info("元の候補リストを使用します")
            except Exception as e:
                logging.warning(f"住所検索フィールドの操作に失敗: {str(e)}")
            
            # 住所を選択
            best_candidate, similarity = find_best_address_match(base_address, valid_candidates)
            
            if best_candidate:
                selected_address = best_candidate.text.strip().split('\n')[0]
                logging.info(f"選択された住所: '{selected_address}' (類似度: {similarity})")
                
                # 選択された住所でクリックを実行
                try:
                    WebDriverWait(driver, 3).until(EC.element_to_be_clickable(best_candidate))
                    best_candidate.click()
                    logging.info("住所を選択しました")
                except Exception as click_error:
                    logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                    try:
                        driver.execute_script("arguments[0].click();", best_candidate)
                        logging.info("JavaScriptでクリックしました")
                    except Exception as js_error:
                        logging.warning(f"JavaScriptのクリックに失敗: {str(js_error)}")
                        ActionChains(driver).move_to_element(best_candidate).click().perform()
                        logging.info("ActionChainsでクリックしました")
            else:
                logging.error(f"適切な住所候補が見つかりませんでした。入力住所: {base_address}")
                raise ValueError("適切な住所候補が見つかりませんでした")
            
            # 住所選択後の読み込みを待つ
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "addressSelectModal"))
            )
            logging.info("住所選択モーダルが閉じられました")
            
            # 5. 番地入力画面が表示された場合は、番地を入力
            try:
                if progress_callback:
                    progress_callback("番地を入力中...")
                
                # 番地入力ダイアログが表示されるまで待機
                banchi_dialog = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID01"))
                )
                logging.info("番地入力ダイアログが表示されました")
                
                # 番地がない場合は「該当する住所がない」を選択
                if not street_number:
                    logging.info("番地が指定されていないため、「該当する住所がない」を選択します")
                    try:
                        # 「該当する住所がない」のリンクを探す
                        no_address_link = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//dialog[@id='DIALOG_ID01']//a[contains(text(), '該当する住所がない')]"))
                        )
                        
                        # スクロールしてリンクを表示
                        driver.execute_script("arguments[0].scrollIntoView(true);", no_address_link)
                        time.sleep(1)
                        
                        # クリックを試行（複数の方法）
                        try:
                            no_address_link.click()
                            logging.info("通常のクリックで「該当する住所がない」を選択しました")
                        except Exception as click_error:
                            logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                            try:
                                driver.execute_script("arguments[0].click();", no_address_link)
                                logging.info("JavaScriptでクリックしました")
                            except Exception as js_error:
                                logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                ActionChains(driver).move_to_element(no_address_link).click().perform()
                                logging.info("ActionChainsでクリックしました")
                        
                        # クリック後の待機
                        time.sleep(2)
                        
                    except Exception as e:
                        logging.error(f"「該当する住所がない」の選択に失敗: {str(e)}")
                        driver.save_screenshot("debug_no_address_error.png")
                        raise
                else:
                    # 番地を入力
                    input_street_number = street_number
                    logging.info(f"入力予定の番地: {input_street_number}")
                    
                    # 全角数字に変換
                    zen_numbers = "０１２３４５６７８９"
                    han_numbers = "0123456789"
                    trans_table = str.maketrans(han_numbers, zen_numbers)
                    zen_street_number = input_street_number.translate(trans_table)
                    logging.info(f"全角変換後の番地: {zen_street_number}")
                    
                    # 番地ボタンを探す
                    try:
                        # 全ての番地ボタンを取得
                        all_buttons = driver.find_elements(By.XPATH, "//dialog[@id='DIALOG_ID01']//a")
                        logging.info(f"全ての番地ボタン数: {len(all_buttons)}")
                        
                        # 「該当する住所がない」ボタンと目的の番地ボタンを探す
                        banchi_button = None
                        no_address_button = None
                        
                        for button in all_buttons:
                            try:
                                button_text = button.text.strip()
                                logging.info(f"番地ボタンのテキスト: {button_text}")
                                
                                if button_text == "該当する住所がない":
                                    no_address_button = button
                                    logging.info("「該当する住所がない」ボタンが見つかりました")
                                # 全角・半角どちらでも一致するか確認
                                elif button_text == input_street_number or button_text == zen_street_number:
                                    banchi_button = button
                                    logging.info(f"番地ボタンが見つかりました: {button_text}")
                            except Exception as e:
                                logging.warning(f"ボタンテキストの取得中にエラー: {str(e)}")
                                continue
                        
                        # 目的の番地が見つからない場合は「該当する住所がない」を選択
                        target_button = banchi_button if banchi_button else no_address_button
                        
                        if target_button:
                            try:
                                # スクロールしてボタンを表示
                                driver.execute_script("arguments[0].scrollIntoView(true);", target_button)
                                time.sleep(1)
                                
                                # クリックを試行（複数の方法）
                                try:
                                    target_button.click()
                                    logging.info("通常のクリックで選択しました")
                                except Exception as click_error:
                                    logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                    try:
                                        driver.execute_script("arguments[0].click();", target_button)
                                        logging.info("JavaScriptでクリックしました")
                                    except Exception as js_error:
                                        logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                        ActionChains(driver).move_to_element(target_button).click().perform()
                                        logging.info("ActionChainsでクリックしました")
                            
                            except Exception as e:
                                logging.error(f"ボタンのクリックに失敗: {str(e)}")
                                raise
                        else:
                            logging.error("適切なボタンが見つかりませんでした")
                            driver.save_screenshot("debug_banchi_not_found.png")
                            raise ValueError("適切なボタンが見つかりませんでした")
                            
                    except Exception as e:
                        logging.error(f"番地選択処理中にエラー: {str(e)}")
                        driver.save_screenshot("debug_banchi_error.png")
                        logging.info("エラー発生時のスクリーンショットを保存しました")
                        raise
                    
                # 番地入力後の読み込みを待つ
                try:
                    WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.ID, "DIALOG_ID01"))
                    )
                    logging.info("番地入力ダイアログが閉じられました")
                except TimeoutException:
                    logging.warning("番地入力ダイアログが閉じられるのを待機中にタイムアウト")
                    # ダイアログが閉じられない場合でも処理を続行
                
            except TimeoutException:
                logging.info("番地入力画面はスキップされました")
            
            # 6. 号入力画面が表示された場合は、号を入力
            try:
                if progress_callback:
                    progress_callback("号を入力中...")
                
                # 号入力ダイアログが表示されるまで待機
                gou_dialog = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID02"))
                )
                logging.info("号入力ダイアログが表示されました")
                
                # 号を入力
                input_building_number = building_number if building_number else "1"
                logging.info(f"入力予定の号: {input_building_number}")
                
                # 全角数字に変換
                zen_numbers = "０１２３４５６７８９"
                han_numbers = "0123456789"
                trans_table = str.maketrans(han_numbers, zen_numbers)
                zen_building_number = input_building_number.translate(trans_table)
                logging.info(f"全角変換後の号: {zen_building_number}")
                
                # 号ボタンを探す
                try:
                    # 全ての号ボタンを取得
                    all_buttons = driver.find_elements(By.XPATH, "//dialog[@id='DIALOG_ID02']//a")
                    logging.info(f"全ての号ボタン数: {len(all_buttons)}")
                    
                    # 「該当する住所がない」ボタンと目的の号ボタンを探す
                    gou_button = None
                    no_address_button = None
                    
                    for button in all_buttons:
                        try:
                            button_text = button.text.strip()
                            logging.info(f"号ボタンのテキスト: {button_text}")
                            
                            if button_text == "該当する住所がない":
                                no_address_button = button
                                logging.info("「該当する住所がない」ボタンが見つかりました")
                            # 全角・半角どちらでも一致するか確認
                            elif button_text == input_building_number or button_text == zen_building_number:
                                gou_button = button
                                logging.info(f"号ボタンが見つかりました: {button_text}")
                        except Exception as e:
                            logging.warning(f"ボタンテキストの取得中にエラー: {str(e)}")
                            continue
                    
                    # 目的の号が見つからない場合は「該当する住所がない」を選択
                    target_button = gou_button if gou_button else no_address_button
                    
                    if target_button:
                        try:
                            # スクロールしてボタンを表示
                            driver.execute_script("arguments[0].scrollIntoView(true);", target_button)
                            time.sleep(1)
                            
                            # クリックを試行（複数の方法）
                            try:
                                target_button.click()
                                logging.info("通常のクリックで選択しました")
                            except Exception as click_error:
                                logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                try:
                                    driver.execute_script("arguments[0].click();", target_button)
                                    logging.info("JavaScriptでクリックしました")
                                except Exception as js_error:
                                    logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                    ActionChains(driver).move_to_element(target_button).click().perform()
                                    logging.info("ActionChainsでクリックしました")
                        
                            # クリック後の待機
                            time.sleep(2)
                            
                        except Exception as e:
                            logging.error(f"ボタンのクリックに失敗: {str(e)}")
                            raise
                    else:
                        logging.error("適切なボタンが見つかりませんでした")
                        driver.save_screenshot("debug_gou_not_found.png")
                        raise ValueError("適切なボタンが見つかりませんでした")
                        
                except Exception as e:
                    logging.error(f"号選択処理中にエラー: {str(e)}")
                    driver.save_screenshot("debug_gou_error.png")
                    logging.info("エラー発生時のスクリーンショットを保存しました")
                    raise
                
                # 号入力後の読み込みを待つ
                try:
                    WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.ID, "DIALOG_ID02"))
                    )
                    logging.info("号入力ダイアログが閉じられました")
                except TimeoutException:
                    logging.warning("号入力ダイアログが閉じられるのを待機中にタイムアウト")
                    # ダイアログが閉じられない場合でも処理を続行
                
            except TimeoutException:
                logging.info("号入力画面はスキップされました")
            
            # 建物選択モーダルの処理
            handle_building_selection(driver)
            
            # 7. 結果の判定
            try:
                if progress_callback:
                    progress_callback("検索結果を確認中...")
                
                # 検索結果確認ボタンをクリック
                logging.info("検索結果確認ボタンの検出を開始します")
                
                # 指定されたIDを持つボタンを待機して検出
                final_search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "id_tak_bt_nx"))
                )
                
                # ボタンが見つかった場合の情報をログ出力
                button_html = final_search_button.get_attribute('outerHTML')
                logging.info(f"検索結果確認ボタンが見つかりました: {button_html}")
                
                # スクロールしてボタンを表示
                driver.execute_script("arguments[0].scrollIntoView(true);", final_search_button)
                time.sleep(1)
                
                # ボタンをクリック
                final_search_button.click()
                logging.info("検索結果確認ボタンをクリックしました")
                
                # クリック後の画面遷移を待機
                time.sleep(2)
                
                # 提供可否の画像を確認
                try:
                    # 提供可能、調査中、提供不可の画像パターンを定義
                    image_patterns = {
                        "available": {
                            "urls": [
                                "//img[@src='https://flets-w.ntt-west.co.jp/resources/form_element/img/img_available_03.png']",
                                "//img[contains(@src, 'img_available_03.png')]",
                                "//img[contains(@src, 'available')]",
                                "//img[@alt='提供可能']"
                            ],
                            "status": "available",
                            "message": "提供可能",
                            "details": {
                                "判定結果": "OK",
                                "提供エリア": "提供可能エリアです",
                                "備考": "フレッツ光のサービスがご利用いただけます"
                            }
                        },
                        "investigation": {
                            "urls": [
                                "//img[@src='https://flets-w.ntt-west.co.jp/resources/form_element/img/img_investigation_03_1.png']",
                                "//img[contains(@src, 'img_investigation_03')]",
                                "//img[contains(@src, 'investigation')]"
                            ],
                            "status": "failure",
                            "message": "判定失敗",
                            "details": {
                                "判定結果": "判定失敗",
                                "提供エリア": "調査が必要なエリアです",
                                "備考": "住所を特定できないため、担当者がお調べします"
                            }
                        },
                        "not_provided": {
                            "urls": [
                                "//img[@src='https://flets-w.ntt-west.co.jp/resources/form_element/img/img_not_provided.png']",
                                "//img[contains(@src, 'img_not_provided')]",
                                "//img[contains(@src, 'not_provided')]",
                                "//img[contains(@alt, '提供不可')]",
                                "//img[contains(@alt, '未提供')]"
                            ],
                            "status": "unavailable",
                            "message": "未提供",
                            "details": {
                                "判定結果": "NG",
                                "提供エリア": "提供対象外エリアです",
                                "備考": "申し訳ございませんが、このエリアではサービスを提供しておりません"
                            }
                        }
                    }
                    
                    found_image = None
                    found_pattern = None
                    
                    # 各パターンを順番に確認
                    for pattern_name, pattern_info in image_patterns.items():
                        for url_pattern in pattern_info["urls"]:
                            try:
                                image = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, url_pattern))
                                )
                                if image and image.is_displayed():
                                    logging.info(f"{pattern_name}の画像が見つかりました: {url_pattern}")
                                    found_image = image
                                    found_pattern = pattern_info
                                    break
                            except Exception as e:
                                logging.warning(f"パターン {url_pattern} での検索に失敗: {str(e)}")
                                continue
                        if found_image:
                            break
                    
                    if found_image and found_pattern:
                        # 画像確認時のスクリーンショットを保存
                        screenshot_path = f"debug_{found_pattern['status']}_confirmation.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        result_message = f"{found_pattern['message']}状態が確認されました"
                        logging.info(f"{result_message} - スクリーンショットを保存しました")
                        
                        # 結果を返す（JavaScriptによる画面書き換えは行わない）
                        result = found_pattern.copy()
                        result["screenshot"] = screenshot_path
                        result["show_popup"] = show_popup  # ポップアップ表示設定を追加
                        logging.info(f"検索結果を返します: status={result['status']}, message={result['message']}")
                        if progress_callback:
                            progress_callback(f"{result['message']}が確認されました")
                        return result
                    else:
                        # 提供不可時のスクリーンショットを保存
                        screenshot_path = "debug_unavailable_confirmation.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        logging.info("判定失敗と判定されました（画像非表示） - スクリーンショットを保存しました")
                        if progress_callback:
                            progress_callback("判定できませんでした")
                        return {
                            "status": "failure",
                            "message": "判定失敗",
                            "details": {
                                "判定結果": "判定失敗",
                                "提供エリア": "判定できませんでした",
                                "備考": "提供可否を判定できませんでした"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup  # ポップアップ表示設定を追加
                        }
                        logging.info(f"検索結果を返します: status={result['status']}, message={result['message']}")
                        return result
                        
                except TimeoutException:
                    # タイムアウト時のスクリーンショットを保存
                    screenshot_path = "debug_timeout_confirmation.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    logging.info("提供可能画像が見つかりませんでした - スクリーンショットを保存しました")
                    if progress_callback:
                        progress_callback("タイムアウトが発生しました")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": "提供可否の確認中にタイムアウトが発生しました"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup  # ポップアップ表示設定を追加
                    }
                except Exception as e:
                    # エラー時のスクリーンショットを保存
                    screenshot_path = "debug_error_confirmation.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    logging.error(f"提供判定の確認中にエラー: {str(e)}")
                    if progress_callback:
                        progress_callback("エラーが発生しました")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": f"エラーが発生しました: {str(e)}"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup  # ポップアップ表示設定を追加
                    }
            
            except Exception as e:
                logging.error(f"結果の判定中にエラー: {str(e)}")
                screenshot_path = "debug_result_error.png"
                take_full_page_screenshot(driver, screenshot_path)
                if progress_callback:
                    progress_callback("エラーが発生しました")
                return {
                    "status": "failure", 
                    "message": "判定失敗",
                    "details": {
                        "判定結果": "判定失敗",
                        "提供エリア": "判定できませんでした",
                        "備考": f"結果の判定中にエラーが発生しました: {str(e)}"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup  # ポップアップ表示設定を追加
                }
                
        except TimeoutException as e:
            logging.error(f"住所候補の表示待ちでタイムアウトしました: {str(e)}")
            screenshot_path = "debug_address_timeout.png"
            take_full_page_screenshot(driver, screenshot_path)
            if progress_callback:
                progress_callback("タイムアウトが発生しました")
            return {
                "status": "failure",
                "message": "判定失敗",
                "details": {
                    "判定結果": "判定失敗",
                    "提供エリア": "判定できませんでした",
                    "備考": "住所候補が見つかりませんでした"
                },
                "screenshot": screenshot_path,
                "show_popup": show_popup  # ポップアップ表示設定を追加
            }
        except Exception as e:
            logging.error(f"住所選択処理中にエラーが発生しました: {str(e)}")
            screenshot_path = "debug_address_error.png"
            take_full_page_screenshot(driver, screenshot_path)
            if progress_callback:
                progress_callback("エラーが発生しました")
            return {
                "status": "failure",
                "message": "判定失敗",
                "details": {
                    "判定結果": "判定失敗",
                    "提供エリア": "判定できませんでした",
                    "備考": f"住所選択に失敗しました: {str(e)}"
                },
                "screenshot": screenshot_path,
                "show_popup": show_popup  # ポップアップ表示設定を追加
            }
    
    except Exception as e:
        logging.error(f"自動化に失敗しました: {str(e)}")
        screenshot_path = "debug_general_error.png"
        if driver:
            take_full_page_screenshot(driver, screenshot_path)
        if progress_callback:
            progress_callback("エラーが発生しました")
        return {
            "status": "failure",
            "message": "判定失敗",
            "details": {
                "判定結果": "判定失敗",
                "提供エリア": "判定できませんでした",
                "備考": f"エラーが発生しました: {str(e)}"
            },
            "screenshot": screenshot_path,
            "show_popup": show_popup  # ポップアップ表示設定を追加
        }
    
    finally:
        # どのような場合でもブラウザは閉じない
        if driver:
            logging.info("ブラウザウィンドウを維持します - 手動で閉じてください")            # driver.quit() を呼び出さない
            # グローバル変数にドライバーを保持
            global_driver = driver
