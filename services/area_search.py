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

from services.web_driver import create_driver
from utils.string_utils import normalize_string, calculate_similarity

# グローバル変数でブラウザドライバーを保持
global_driver = None

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
            
        # ウィンドウサイズを800x600に設定
        options.add_argument('--window-size=800,600')
        
        driver = webdriver.Chrome(options=options)
        logging.info("Chromeドライバーを作成しました")
        
        return driver
    except Exception as e:
        logging.error(f"ドライバーの作成に失敗: {str(e)}")
        raise

def search_service_area(postal_code, address):
    """
    NTT西日本の提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        
    Returns:
        dict: 検索結果を含む辞書
    """
    global global_driver
    
    logging.info(f"郵便番号 {postal_code}、住所 {address} の処理を開始します")
    
    # 住所を分割
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
    except Exception as e:
        logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")
    
    # ヘッドレスモードの設定を取得
    headless_mode = browser_settings.get("headless", False)
    # ポップアップ表示設定を取得
    show_popup = browser_settings.get("show_popup", True)
    # ブラウザ自動終了設定を取得
    auto_close = browser_settings.get("auto_close", True)
    
    # ポップアップを表示する場合は強制的にヘッドレスモードを無効化
    if show_popup:
        headless_mode = False
        logging.info("ポップアップ表示が有効なため、ヘッドレスモードを無効化します")
    
    driver = None
    try:
        # 1. ドライバーを作成してサイトを開く
        # headless=Falseを強制して常にブラウザウィンドウを表示する
        if show_popup:
            driver = create_driver(headless=False)
            logging.info("表示モードでブラウザを強制的に起動します")
        else:
            driver = create_driver(headless=headless_mode)
            
        # グローバル変数に保存
        global_driver = driver
        
        driver.implicitly_wait(0)  # 暗黙の待機を無効化
        
        # 非ヘッドレスモードの場合、ウィンドウが確実に表示されるよう少し待機
        if not headless_mode:
            time.sleep(2)  # ウィンドウが初期化されるまで2秒待機
            logging.info("表示モードでブラウザを起動しました - ウィンドウが表示されていることを確認してください")
        
        driver.get("https://flets-w.com/cart/")
        logging.info("NTT西日本のサイトにアクセスしています...")
        
        # ページが完全に読み込まれるまで待機（10秒まで）
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        logging.info("ページが読み込まれました")
        
        # 2. 郵便番号を入力
        try:
            zip_field = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_tx_ybk_yb']"))
            )
            zip_field.clear()
            zip_field.send_keys(postal_code_clean)
            logging.info(f"郵便番号 {postal_code_clean} を入力しました")
        except Exception as e:
            logging.error(f"郵便番号入力フィールドが見つかりませんでした: {str(e)}")
            logging.info(f"ページのHTML: {driver.page_source[:500]}...")
            raise
        
        # 3. 検索ボタンを押す
        try:
            search_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_bt_ybk_jks']"))
            )
            search_button.click()
            logging.info("検索ボタンをクリックしました")
        except Exception as e:
            logging.error(f"検索ボタンが見つかりませんでした: {str(e)}")
            raise
        
        # 住所候補が表示されるのを待つ（最大10秒）
        try:
            # 住所選択モーダルが表示されるまで待機
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "addressSelectModal"))
            )
            logging.info("住所選択モーダルが表示されました")
            
            # 少し待機してモーダルが完全に表示されるのを待つ
            time.sleep(1)
            
            # スクリーンショットを撮影してデバッグ
            driver.save_screenshot("debug_screenshot.png")
            logging.info("デバッグ用スクリーンショットを保存しました")
            
            # 候補リストを取得（複数の方法を試す）
            candidates = []
            
            # 方法1: aタグで候補を取得
            try:
                candidates = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#addressSelectModal ul li a"))
                )
                logging.info(f"方法1: {len(candidates)} 件の候補が見つかりました")
                
                # 候補の内容をログ出力
                for i, candidate in enumerate(candidates):
                    logging.info(f"候補 {i+1}: {candidate.text.strip()}")
            except Exception as e:
                logging.warning(f"方法1での候補取得に失敗: {str(e)}")
            
            # 方法2: liタグで候補を取得
            if not candidates:
                try:
                    candidates = WebDriverWait(driver, 5).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#addressSelectModal ul li"))
                    )
                    logging.info(f"方法2: {len(candidates)} 件の候補が見つかりました")
                    
                    # 候補の内容をログ出力
                    for i, candidate in enumerate(candidates):
                        logging.info(f"候補 {i+1}: {candidate.text.strip()}")
                except Exception as e:
                    logging.warning(f"方法2での候補取得に失敗: {str(e)}")
            
            # 方法3: JavaScriptで取得
            if not candidates:
                try:
                    js_candidates = driver.execute_script("""
                        return Array.from(document.querySelectorAll('#addressSelectModal ul li a')).filter(el => {
                            const text = el.textContent.trim();
                            return text && text.length > 0;
                        });
                    """)
                    if js_candidates:
                        candidates = js_candidates
                        logging.info(f"方法3: {len(candidates)} 件の候補が見つかりました")
                        
                        # 候補の内容をログ出力
                        for i, candidate in enumerate(candidates):
                            logging.info(f"候補 {i+1}: {candidate.get_attribute('textContent').strip()}")
                except Exception as e:
                    logging.warning(f"方法3での候補取得に失敗: {str(e)}")
            
            # デバッグ情報の出力
            logging.info("=== モーダルの構造 ===")
            modal_html = driver.find_element(By.ID, "addressSelectModal").get_attribute('outerHTML')
            logging.info(f"モーダルのHTML: {modal_html[:500]}...")
            
            # スクリーンショットを保存（デバッグ用）
            driver.save_screenshot("debug_modal.png")
            logging.info("モーダルのスクリーンショットを保存しました")
            
            # 有効な候補（テキストが空でない）をフィルタリング
            valid_candidates = [c for c in candidates if c.text.strip()]
            
            if not valid_candidates:
                raise NoSuchElementException("有効な住所候補が見つかりませんでした")
            
            logging.info(f"有効な候補数: {len(valid_candidates)}")
            for i, candidate in enumerate(valid_candidates):
                logging.info(f"有効な候補 {i+1}: {candidate.text.strip()}")
            
            # 住所候補が多い場合（スクロール可能な場合）は、検索フィールドで絞り込み
            try:
                search_field = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='絞り込みワードを入力']"))
                )
                logging.info("住所検索フィールドが見つかりました")
                
                # 検索用の住所フォーマットを作成（都道府県、市区町村、町名の間に半角スペースを挿入）
                search_address = base_address.replace("県", "県 ").replace("市", "市 ").replace("町", "町 ").replace("区", "区 ").strip()
                logging.info(f"検索用にフォーマットされた住所: {search_address}")
                
                # 検索フィールドをクリアして入力
                search_field.clear()
                search_field.send_keys(search_address)
                logging.info(f"検索フィールドに '{search_address}' を入力しました")
                
                # 入力後の表示更新を待機（少し長めに）
                time.sleep(2)
                
                # 絞り込み後の候補を取得
                filtered_candidates = driver.find_elements(By.XPATH, "//*[@id='addressSelectModal']//div[contains(@class, 'clickable')]")
                if filtered_candidates:
                    valid_candidates = filtered_candidates
                    logging.info(f"絞り込み後の候補数: {len(valid_candidates)}")
                    
                    # 絞り込み後の候補をログ出力
                    for i, candidate in enumerate(valid_candidates[:5]):
                        logging.info(f"絞り込み後の候補 {i+1}: '{candidate.text.strip()}'")
                else:
                    logging.warning("絞り込み後の候補が見つかりませんでした")
                    # 絞り込みに失敗した場合、元の候補リストを使用
                    logging.info("元の候補リストを使用します")
            except Exception as e:
                logging.warning(f"住所検索フィールドの操作に失敗: {str(e)}")
            
            # 4. 住所を選択
            best_candidate, similarity = find_best_address_match(base_address, valid_candidates)

            if best_candidate:
                selected_address = best_candidate.text.strip().split('\n')[0]
                logging.info(f"選択された住所: '{selected_address}' (類似度: {similarity})")
                
                # 選択された住所でクリックを実行
                try:
                    WebDriverWait(driver, 15).until(EC.element_to_be_clickable(best_candidate))
                    
                    # クリックを試行（複数の方法）
                    try:
                        best_candidate.click()
                        logging.info("通常のクリックで住所を選択しました")
                    except Exception as click_error:
                        logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                        try:
                            driver.execute_script("arguments[0].click();", best_candidate)
                            logging.info("JavaScriptでクリックしました")
                        except Exception as js_error:
                            logging.warning(f"JavaScriptのクリックに失敗: {str(js_error)}")
                            ActionChains(driver).move_to_element(best_candidate).click().perform()
                            logging.info("ActionChainsでクリックしました")
                except Exception as e:
                    logging.error(f"住所選択のクリックに失敗: {str(e)}")
                    raise
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
                        driver.save_screenshot(screenshot_path)
                        result_message = f"{found_pattern['message']}状態が確認されました"
                        logging.info(f"{result_message} - スクリーンショットを保存しました")
                        
                        # 結果を返す（JavaScriptによる画面書き換えは行わない）
                        result = found_pattern.copy()
                        result["screenshot"] = os.path.abspath(screenshot_path)
                        result["show_popup"] = show_popup  # ポップアップ表示設定を追加
                        return result
                    else:
                        # 提供不可時のスクリーンショットを保存
                        screenshot_path = "debug_unavailable_confirmation.png"
                        driver.save_screenshot(screenshot_path)
                        logging.info("判定失敗と判定されました（画像非表示） - スクリーンショットを保存しました")
                        return {
                            "status": "failure",
                            "message": "判定失敗",
                            "details": {
                                "判定結果": "判定失敗",
                                "提供エリア": "判定できませんでした",
                                "備考": "提供可否を判定できませんでした"
                            },
                            "screenshot": os.path.abspath(screenshot_path),
                            "show_popup": show_popup  # ポップアップ表示設定を追加
                        }
                        
                except TimeoutException:
                    # タイムアウト時のスクリーンショットを保存
                    screenshot_path = "debug_timeout_confirmation.png"
                    driver.save_screenshot(screenshot_path)
                    logging.info("提供可能画像が見つかりませんでした - スクリーンショットを保存しました")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": "提供可否の確認中にタイムアウトが発生しました"
                        },
                        "screenshot": os.path.abspath(screenshot_path),
                        "show_popup": show_popup  # ポップアップ表示設定を追加
                    }
                except Exception as e:
                    # エラー時のスクリーンショットを保存
                    screenshot_path = "debug_error_confirmation.png"
                    driver.save_screenshot(screenshot_path)
                    logging.error(f"提供判定の確認中にエラー: {str(e)}")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": f"エラーが発生しました: {str(e)}"
                        },
                        "screenshot": os.path.abspath(screenshot_path),
                        "show_popup": show_popup  # ポップアップ表示設定を追加
                    }
            
            except Exception as e:
                logging.error(f"結果の判定中にエラー: {str(e)}")
                screenshot_path = "debug_result_error.png"
                driver.save_screenshot(screenshot_path)
                return {
                    "status": "failure", 
                    "message": "判定失敗",
                    "details": {
                        "判定結果": "判定失敗",
                        "提供エリア": "判定できませんでした",
                        "備考": f"結果の判定中にエラーが発生しました: {str(e)}"
                    },
                    "screenshot": os.path.abspath(screenshot_path),
                    "show_popup": show_popup  # ポップアップ表示設定を追加
                }
                
        except TimeoutException as e:
            logging.error(f"住所候補の表示待ちでタイムアウトしました: {str(e)}")
            screenshot_path = "debug_address_timeout.png"
            driver.save_screenshot(screenshot_path)
            return {
                "status": "failure",
                "message": "判定失敗",
                "details": {
                    "判定結果": "判定失敗",
                    "提供エリア": "判定できませんでした",
                    "備考": "住所候補が見つかりませんでした"
                },
                "screenshot": os.path.abspath(screenshot_path),
                "show_popup": show_popup  # ポップアップ表示設定を追加
            }
        except Exception as e:
            logging.error(f"住所選択処理中にエラーが発生しました: {str(e)}")
            screenshot_path = "debug_address_error.png"
            driver.save_screenshot(screenshot_path)
            return {
                "status": "failure",
                "message": "判定失敗",
                "details": {
                    "判定結果": "判定失敗",
                    "提供エリア": "判定できませんでした",
                    "備考": f"住所選択に失敗しました: {str(e)}"
                },
                "screenshot": os.path.abspath(screenshot_path),
                "show_popup": show_popup  # ポップアップ表示設定を追加
            }
    
    except Exception as e:
        logging.error(f"自動化に失敗しました: {str(e)}")
        screenshot_path = "debug_general_error.png"
        if driver:
            driver.save_screenshot(screenshot_path)
        return {
            "status": "failure",
            "message": "判定失敗",
            "details": {
                "判定結果": "判定失敗",
                "提供エリア": "判定できませんでした",
                "備考": f"エラーが発生しました: {str(e)}"
            },
            "screenshot": os.path.abspath(screenshot_path),
            "show_popup": show_popup  # ポップアップ表示設定を追加
        }
    
    finally:
        # どのような場合でもブラウザは閉じない
        if driver:
            logging.info("ブラウザウィンドウを維持します - 手動で閉じてください")
            # driver.quit() を呼び出さない
            # グローバル変数にドライバーを保持
            global_driver = driver