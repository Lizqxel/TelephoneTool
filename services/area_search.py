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
    住所を分割する関数
    入力形式：[漢字による住所][数字]-[数字](-[数字])
    例：奈良県奈良市五条西2丁目2-1

    Args:
        address (str): 分割する住所文字列
        
    Returns:
        tuple: (基本住所, 番地, 号)
        例：('奈良県奈良市五条西2丁目', '2', '1')
    """
    if not address:
        return ("", None, None)

    # 住所を正規化
    address = normalize_address(address)
    
    # 丁目を含む場合の処理
    chome_match = re.search(r'^(.+?[0-9]+丁目)([0-9]+)(?:-([0-9]+))?', address)
    if chome_match:
        base = chome_match.group(1)
        num1 = chome_match.group(2)
        num2 = chome_match.group(3)
        return (base, num1, num2)

    # 基本パターン：[漢字と数字の住所]-[数字]-[数字]
    pattern = r'^(.+?)([0-9]+)-([0-9]+)(?:-([0-9]+))?$'
    match = re.search(pattern, address)
    
    if match:
        base = match.group(1)
        num1 = match.group(2)
        num2 = match.group(3)
        num3 = match.group(4)  # オプショナル
        
        # 3つの数字がある場合（例：1-3-4）
        if num3:
            return (base, num2, num3)
        # 2つの数字がある場合（例：19-10）
        else:
            return (base, num1, num2)
    
    # 単純な番地のパターン
    simple_match = re.search(r'^(.+?)([0-9]+)(?:番地?)?$', address)
    if simple_match:
        return (simple_match.group(1), simple_match.group(2), None)
    
    return (address, None, None)

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
    
    # 全角ハイフンを半角に変換
    normalized = normalized.replace('−', '-').replace('ー', '-').replace('－', '-')
    
    # すべてのスペース（全角・半角）を一旦半角スペースに統一
    normalized = normalized.replace('　', ' ')
    
    # 数字の前後のスペースを削除
    normalized = re.sub(r'\s+(\d+)', r'\1', normalized)  # 数字の前のスペースを削除
    normalized = re.sub(r'(\d+)\s+', r'\1', normalized)  # 数字の後のスペースを削除
    
    # 都道府県、市区町村の区切りを統一（スペースを削除）
    normalized = normalized.replace('県 ', '県').replace('市 ', '市').replace('区 ', '区').replace('町 ', '町')
    
    # 「大字」を削除
    normalized = normalized.replace('大字', '')
    
    # 連続するスペースを1つに統一
    normalized = ' '.join(normalized.split())
    
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

def is_address_match(input_address, candidate_address):
    """
    入力された住所と候補の住所が一致するかを判定する
    
    Args:
        input_address (str): 入力された住所（例：奈良県奈良市五条西2丁目）
        candidate_address (str): 候補の住所
        
    Returns:
        bool: 住所が一致する場合はTrue、それ以外はFalse
    """
    # 両方の住所を正規化
    normalized_input = normalize_string(input_address)
    normalized_candidate = normalize_string(candidate_address)
    
    logging.info(f"住所比較 - 入力: {normalized_input} vs 候補: {normalized_candidate}")
    
    # 完全一致の場合
    if normalized_input == normalized_candidate:
        logging.info("完全一致しました")
        return True
    
    # 丁目を含む場合の処理
    input_match = re.match(r'^(.+?)(\d+)丁目', normalized_input)
    candidate_match = re.match(r'^(.+?)(\d+)丁目', normalized_candidate)
    
    if input_match and candidate_match:
        input_base = input_match.group(1)  # 丁目の前までの部分
        input_chome = input_match.group(2)  # 丁目の数字
        candidate_base = candidate_match.group(1)
        candidate_chome = candidate_match.group(2)
        
        logging.info(f"丁目比較 - 入力: {input_base}{input_chome}丁目 vs 候補: {candidate_base}{candidate_chome}丁目")
        
        # 基本部分と丁目の数字が完全一致
        if input_base == candidate_base and input_chome == candidate_chome:
            logging.info("丁目まで完全一致しました")
            return True
    
    # 基本的な住所部分の比較（地名の追加部分を無視）
    normalized_input_parts = normalized_input.split()
    normalized_candidate_parts = normalized_candidate.split()
    
    # 入力住所の各部分が候補住所に含まれているかチェック
    if len(normalized_input_parts) <= len(normalized_candidate_parts):
        all_parts_match = all(
            any(input_part == candidate_part for candidate_part in normalized_candidate_parts)
            for input_part in normalized_input_parts
        )
        if all_parts_match:
            logging.info("基本住所部分が一致しました")
            return True
    
    logging.info("マッチしませんでした")
    return False

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
    base_address, street_number, building_number = split_address(address)
    logging.info(f"住所分割結果 - 基本住所: {base_address}, 番地: {street_number}, 号: {building_number}")
    
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
            
            # 4. 住所を選択（完全一致のみ）
            best_candidate = None
            exact_match = None
            
            normalized_input_address = normalize_string(base_address)
            logging.info(f"正規化された入力住所: {normalized_input_address}")
            
            # 各候補を個別に処理
            for candidate in valid_candidates:
                try:
                    # 候補のテキストを取得（改行で分割された場合は最初の行のみ使用）
                    candidate_text = candidate.text.strip().split('\n')[0]
                    normalized_candidate = normalize_string(candidate_text)
                    logging.info(f"候補住所の比較: {normalized_candidate}")
                    
                    # 完全一致を確認
                    if normalized_input_address == normalized_candidate:
                        exact_match = candidate
                        logging.info(f"完全一致する住所が見つかりました: {candidate_text}")
                        break
                except Exception as e:
                    logging.warning(f"候補の処理中にエラー: {str(e)}")
                    continue
            
            # 完全一致する候補のみを選択
            if exact_match:
                best_candidate = exact_match
                logging.info("完全一致する住所を選択します")
            else:
                logging.error(f"完全一致する住所が見つかりませんでした。入力住所: {base_address}")
                raise ValueError("完全一致する住所が見つかりませんでした")
            
            # 選択された住所の確認
            selected_address = best_candidate.text.strip().split('\n')[0]
            logging.info(f"最終的に選択された住所: '{selected_address}'")
            
            # 選択された住所が期待する住所と完全一致することを確認
            if normalize_string(selected_address) != normalized_input_address:
                logging.error(f"選択された住所が期待する住所と一致しません。期待: {base_address}, 実際: {selected_address}")
                raise ValueError("正しい住所を選択できませんでした")
            
            # クリック可能になるまで待機
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable(best_candidate))
            
            # 複数の方法でクリックを試みる
            try:
                # 方法1: 通常のクリック
                best_candidate.click()
                logging.info("通常のクリックで住所を選択しました")
            except Exception as e:
                logging.warning(f"通常のクリックに失敗: {str(e)}")
                try:
                    # 方法2: JavaScriptでクリック
                    driver.execute_script("arguments[0].click();", best_candidate)
                    logging.info("JavaScriptのクリックで住所を選択しました")
                except Exception as e2:
                    logging.warning(f"JavaScriptのクリックにも失敗: {str(e2)}")
                    try:
                        # 方法3: ActionChainsでクリック
                        ActionChains(driver).move_to_element(best_candidate).click().perform()
                        logging.info("ActionChainsのクリックで住所を選択しました")
                    except Exception as e3:
                        logging.error(f"すべてのクリック方法が失敗: {str(e3)}")
                        raise
            
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
                        
                        banchi_button = None
                        for button in all_buttons:
                            try:
                                button_text = button.text.strip()
                                logging.info(f"番地ボタンのテキスト: {button_text}")
                                
                                # 全角・半角どちらでも一致するか確認
                                if button_text == input_street_number or button_text == zen_street_number:
                                    banchi_button = button
                                    logging.info(f"番地ボタンが見つかりました: {button_text}")
                                    break
                            except Exception as e:
                                logging.warning(f"ボタンテキストの取得中にエラー: {str(e)}")
                                continue
                        
                        if banchi_button:
                            # ボタンが見つかった場合、クリックを試みる
                            try:
                                # スクロールしてボタンを表示
                                driver.execute_script("arguments[0].scrollIntoView(true);", banchi_button)
                                time.sleep(1)
                                
                                # クリックを試行（複数の方法）
                                try:
                                    banchi_button.click()
                                    logging.info("通常のクリックで番地を選択しました")
                                except Exception as click_error:
                                    logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                    try:
                                        driver.execute_script("arguments[0].click();", banchi_button)
                                        logging.info("JavaScriptでクリックしました")
                                    except Exception as js_error:
                                        logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                        ActionChains(driver).move_to_element(banchi_button).click().perform()
                                        logging.info("ActionChainsでクリックしました")
                            
                            except Exception as e:
                                logging.error(f"番地ボタンのクリックに失敗: {str(e)}")
                                raise
                        else:
                            logging.error("番地ボタンが見つかりませんでした")
                            driver.save_screenshot("debug_banchi_not_found.png")
                            raise ValueError("番地ボタンが見つかりませんでした")
                            
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
                    
                    gou_button = None
                    for button in all_buttons:
                        try:
                            button_text = button.text.strip()
                            logging.info(f"号ボタンのテキスト: {button_text}")
                            
                            # 全角・半角どちらでも一致するか確認
                            if button_text == input_building_number or button_text == zen_building_number:
                                gou_button = button
                                logging.info(f"号ボタンが見つかりました: {button_text}")
                                break
                        except Exception as e:
                            logging.warning(f"ボタンテキストの取得中にエラー: {str(e)}")
                            continue
                    
                    if gou_button:
                        # ボタンが見つかった場合、クリックを試みる
                        try:
                            # スクロールしてボタンを表示
                            driver.execute_script("arguments[0].scrollIntoView(true);", gou_button)
                            time.sleep(1)
                            
                            # クリックを試行（複数の方法）
                            try:
                                gou_button.click()
                                logging.info("通常のクリックで号を選択しました")
                            except Exception as click_error:
                                logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                try:
                                    driver.execute_script("arguments[0].click();", gou_button)
                                    logging.info("JavaScriptでクリックしました")
                                except Exception as js_error:
                                    logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                    ActionChains(driver).move_to_element(gou_button).click().perform()
                                    logging.info("ActionChainsでクリックしました")
                        
                            # クリック後の待機
                            time.sleep(2)
                            
                        except Exception as e:
                            logging.error(f"号ボタンのクリックに失敗: {str(e)}")
                            raise
                    else:
                        # ボタンが見つからない場合は検索フィールドを使用
                        try:
                            gou_field = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, "//*[@id='DIALOG_ID02']//input"))
                            )
                            
                            # 号を入力
                            gou_field.clear()
                            gou_field.send_keys(input_building_number)
                            logging.info(f"号「{input_building_number}」を入力しました")
                            time.sleep(1)
                            
                            # Enterキーを送信
                            gou_field.send_keys(Keys.RETURN)
                            logging.info("Enterキーを送信しました")
                            
                            # 入力後の表示更新を待機
                            time.sleep(2)
                            
                            # 入力した号に一致するボタンを探す
                            matching_buttons = driver.find_elements(By.XPATH, f"//dialog[@id='DIALOG_ID02']//a[contains(text(), '{input_building_number}')]")
                            if matching_buttons:
                                # 最初に見つかった一致するボタンをクリック
                                matching_button = matching_buttons[0]
                                driver.execute_script("arguments[0].scrollIntoView(true);", matching_button)
                                time.sleep(1)
                                matching_button.click()
                                logging.info(f"検索結果から号「{input_building_number}」を選択しました")
                            else:
                                logging.warning("検索後も一致する号ボタンが見つかりませんでした")
                                # 最も近い号を選択
                                all_buttons = driver.find_elements(By.XPATH, "//dialog[@id='DIALOG_ID02']//a")
                                closest_button = None
                                min_diff = float('inf')
                                target_num = int(input_building_number)
                                
                                for button in all_buttons:
                                    try:
                                        button_text = button.text.strip()
                                        if button_text.isdigit():
                                            button_num = int(button_text)
                                            diff = abs(button_num - target_num)
                                            if diff < min_diff:
                                                min_diff = diff
                                                closest_button = button
                                    except ValueError:
                                        continue
                                
                                if closest_button:
                                    driver.execute_script("arguments[0].scrollIntoView(true);", closest_button)
                                    time.sleep(1)
                                    closest_button.click()
                                    logging.info(f"最も近い号「{closest_button.text.strip()}」を選択しました")
                                else:
                                    raise ValueError("適切な号が見つかりませんでした")
                            
                        except TimeoutException:
                            logging.warning("検索フィールドが見つかりませんでした")
                            raise
                        
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