"""
NTT東日本の提供エリア検索サービス

このモジュールは、NTT東日本の提供エリア検索を
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
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver

from services.web_driver import create_driver, load_browser_settings
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
    住所文字列を都道府県、市区町村、町名、番地に分割する
    
    Args:
        address (str): 分割する住所文字列
        
    Returns:
        dict: 分割された住所情報
    """
    try:
        # 住所を正規化
        address = normalize_address(address)
        
        # 都道府県を抽出
        prefecture_pattern = r'^(東京都|北海道|(?:京都|大阪)府|.+?県)'
        prefecture_match = re.match(prefecture_pattern, address)
        if not prefecture_match:
            raise ValueError("都道府県が見つかりません")
        prefecture = prefecture_match.group(1)
        
        # 都道府県を除去
        remaining = address[len(prefecture):]
        
        # 市区町村を抽出
        city_pattern = r'^(.+?[市区町村])'
        city_match = re.match(city_pattern, remaining)
        if not city_match:
            raise ValueError("市区町村が見つかりません")
        city = city_match.group(1)
        
        # 市区町村を除去
        remaining = remaining[len(city):]
        
        # 町名を抽出（丁目まで含む）
        town_pattern = r'^(.+?(?:丁目)?)'
        town_match = re.match(town_pattern, remaining)
        if not town_match:
            raise ValueError("町名が見つかりません")
        town = town_match.group(1)
        
        # 町名を除去
        remaining = remaining[len(town):]
        
        # 番地と号を抽出
        number_pattern = r'^(\d+(?:-\d+)*)'
        number_match = re.match(number_pattern, remaining)
        number = number_match.group(1) if number_match else ""
        
        # 丁目を抽出
        block_pattern = r'(\d+)丁目'
        block_match = re.search(block_pattern, town)
        block = block_match.group(1) if block_match else ""
        
        # 丁目を除去
        if block:
            town = re.sub(r'\d+丁目', '', town)
        
        return {
            'prefecture': prefecture,
            'city': city,
            'town': town,
            'block': block,
            'number': number
        }
    except Exception as e:
        logging.error(f"住所の分割中にエラー: {str(e)}")
        return None

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
    similarity = calculate_similarity(normalized_input, normalized_candidate)
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
            candidate_text = candidate.text.strip()
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
    
def search_service_area(postal_code, address, progress_callback=None):
    """
    NTT東日本の提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        progress_callback (callable): 進捗状況を通知するコールバック関数
        
    Returns:
        dict: 検索結果を含む辞書
    """
    global global_driver
    
    logging.info(f"郵便番号 {postal_code}、住所 {address} の処理を開始します")
    
    # 住所を分割
    if progress_callback:
        progress_callback("住所情報を解析中...")
    
    address_parts = split_address(address)
    if not address_parts:
        return {"status": "error", "message": "住所の分割に失敗しました。"}
    
    # 郵便番号のフォーマットチェック
    postal_code_clean = postal_code.replace("-", "")
    if len(postal_code_clean) != 7 or not postal_code_clean.isdigit():
        return {"status": "error", "message": "郵便番号は7桁の数字で入力してください。"}
    
    # 郵便番号を前半3桁と後半4桁に分割
    postal_code_first = postal_code_clean[:3]
    postal_code_second = postal_code_clean[3:]
    
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
        # ブラウザを起動
        if progress_callback:
            progress_callback("ブラウザを起動中...")
        
        driver = create_driver(
            headless=headless_mode
        )
        
        # グローバル変数に保存
        global_driver = driver
         # タイムアウト設定を適用
        driver.set_page_load_timeout(page_load_timeout)
        driver.set_script_timeout(script_timeout)
        
        driver.implicitly_wait(0)  # 暗黙の待機を無効化
        
        # サイトにアクセス
        if progress_callback:
            progress_callback("サイトにアクセス中...")
        
        driver.get("https://flets.com/app_new/cao/")
        logging.info("サイトにアクセスしました")
        
        # ページのHTMLをログに出力
        logging.info(f"初期ページのHTML:\n{driver.page_source}")
        
        # 郵便番号入力ページが表示されるのを待つ
        # 郵便番号入力フィールドが表示されるまで待機
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "id_address_search_zip1"))
        )
        logging.info("郵便番号入力ページが表示されました")
        
        # ページのHTMLをログに出力
        logging.info(f"郵便番号入力ページのHTML:\n{driver.page_source}")
        
        # 郵便番号入力フィールドを探す
        try:
            # 郵便番号前半3桁を入力
            postal_code_first_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_address_search_zip1"))
            )
            postal_code_first_input.clear()
            postal_code_first_input.send_keys(postal_code_first)
            logging.info(f"郵便番号前半3桁を入力: {postal_code_first}")
            
            # 郵便番号後半4桁を入力
            postal_code_second_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_address_search_zip2"))
            )
            postal_code_second_input.clear()
            postal_code_second_input.send_keys(postal_code_second)
            logging.info(f"郵便番号後半4桁を入力: {postal_code_second}")
            
            # 再検索ボタンをクリック
            search_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "id_address_search_button"))
            )
            search_button.click()
            logging.info("再検索ボタンをクリックしました")
            
            # エラーメッセージの有無を確認
            try:
                
                # エラーメッセージを探す
                error_message = driver.find_elements(By.XPATH, "//*[contains(text(), '入力に誤りがあります')]")
                
                if error_message:
                    error_text = error_message[0].text
                    logging.error(f"郵便番号エラー: {error_text}")
                    return {"status": "error", "message": f"郵便番号エラー: {error_text}"}
            except Exception as e:
                logging.info(f"エラーメッセージの確認中に例外が発生しましたが、処理を続行します: {str(e)}")
            
        except Exception as e:
            logging.error(f"郵便番号入力処理中にエラー: {str(e)}")
            return {"status": "error", "message": f"郵便番号入力処理中にエラーが発生しました: {str(e)}"}
        
        # 住所候補が表示されるのを待つ
        try:
            if progress_callback:
                progress_callback("住所候補を検索中...")
            
            # 住所候補リストが表示されるまで待機
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "addressList"))
            )
            logging.info("住所候補リストが表示されました")
            
            # ページのHTMLをログに出力
            logging.info(f"住所候補リストページのHTML:\n{driver.page_source}")
            
            # 候補リストを取得
            candidates = driver.find_elements(By.CSS_SELECTOR, ".addressList li")
            logging.info(f"{len(candidates)} 件の候補が見つかりました")
            
            # 候補の内容をログ出力
            for i, candidate in enumerate(candidates[:5]):  # 最初の5件のみログ出力
                logging.info(f"候補 {i+1}: {candidate.text.strip()}")
            
            # 有効な候補（テキストが空でない）をフィルタリング
            valid_candidates = [c for c in candidates if c.text.strip()]
            
            if not valid_candidates:
                raise NoSuchElementException("有効な住所候補が見つかりませんでした")
            
            logging.info(f"有効な候補数: {len(valid_candidates)}")
            
            # 住所を選択
            best_candidate, similarity = find_best_address_match(address, valid_candidates)
            
            if best_candidate:
                selected_address = best_candidate.text.strip()
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
                logging.error(f"適切な住所候補が見つかりませんでした。入力住所: {address}")
                raise ValueError("適切な住所候補が見つかりませんでした")
            
            # 番地入力ページが表示されるのを待つ
            WebDriverWait(driver, 10).until(
                EC.url_contains("InputAddressNum")
            )
            logging.info("番地入力ページが表示されました")
            
            # ページのHTMLをログに出力
            logging.info(f"番地入力ページのHTML:\n{driver.page_source}")
            
            # 番地の有無を確認
            try:
                no_number_button = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "noNumberButton"))
                )
                if no_number_button.is_displayed():
                    logging.info("番地なしボタンが見つかりました")
                    
                    # 番地なしボタンをクリック
                    no_number_button.click()
                    logging.info("番地なしボタンをクリックしました")
                    
                    # 戸建てを選択
                    house_type_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "houseTypeButton"))
                    )
                    house_type_button.click()
                    logging.info("戸建てを選択しました")
                    
                    # 次へボタンをクリック
                    next_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "nextButton"))
                    )
                    next_button.click()
                    logging.info("次へボタンをクリックしました")
            except Exception as e:
                logging.info("番地なしボタンが見つからないため、番地入力を行います")
                
                # 番地を入力
                if address_parts['number']:
                    number_parts = address_parts['number'].split('-')
                    
                    # 番地1を入力
                    try:
                        number1_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.NAME, "number1"))
                        )
                        number1_input.clear()
                        number1_input.send_keys(number_parts[0])
                        logging.info(f"番地1を入力: {number_parts[0]}")
                    except Exception as e:
                        logging.error(f"番地1の入力に失敗: {str(e)}")
                        raise
                    
                    # 番地2を入力（存在する場合）
                    if len(number_parts) > 1:
                        try:
                            number2_input = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.NAME, "number2"))
                            )
                            number2_input.clear()
                            number2_input.send_keys(number_parts[1])
                            logging.info(f"番地2を入力: {number_parts[1]}")
                        except Exception as e:
                            logging.error(f"番地2の入力に失敗: {str(e)}")
                            raise
                    
                    # 番地3を入力（存在する場合）
                    if len(number_parts) > 2:
                        try:
                            number3_input = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.NAME, "number3"))
                            )
                            number3_input.clear()
                            number3_input.send_keys(number_parts[2])
                            logging.info(f"番地3を入力: {number_parts[2]}")
                        except Exception as e:
                            logging.error(f"番地3の入力に失敗: {str(e)}")
                            raise
                
                # 戸建てを選択
                house_type_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "houseTypeButton"))
                )
                house_type_button.click()
                logging.info("戸建てを選択しました")
                
                # 次へボタンをクリック
                next_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "nextButton"))
                )
                next_button.click()
                logging.info("次へボタンをクリックしました")
            
            # 建物選択ページが表示されるか確認
            try:
                WebDriverWait(driver, 5).until(
                    EC.url_contains("SelectBuild1")
                )
                logging.info("建物選択ページが表示されました")
                
                # ページのHTMLをログに出力
                logging.info(f"建物選択ページのHTML:\n{driver.page_source}")
                
                # 建物候補が表示されるのを待つ
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "buildingList"))
                )
                
                # 建物候補を取得
                building_candidates = driver.find_elements(By.CSS_SELECTOR, ".buildingList li")
                
                if building_candidates:
                    # 最初の建物を選択
                    building_candidates[0].click()
                    logging.info("建物を選択しました")
                else:
                    logging.info("建物候補が見つかりませんでした")
            except TimeoutException:
                logging.info("建物選択ページは表示されませんでした")
            
            # 結果ページが表示されるのを待つ
            WebDriverWait(driver, 10).until(
                EC.url_contains("ProvideResult")
            )
            logging.info("結果ページが表示されました")
            
            # ページのHTMLをログに出力
            logging.info(f"結果ページのHTML:\n{driver.page_source}")
            
            # 結果を取得
            try:
                result_text = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "resultText"))
                ).text
                
                if "提供エリアです" in result_text:
                    return {"status": "success", "message": "提供エリアです"}
                else:
                    return {"status": "error", "message": "提供エリア外です"}
            except Exception as e:
                logging.error(f"結果の取得に失敗: {str(e)}")
                return {"status": "error", "message": "結果の取得に失敗しました"}
            
        except Exception as e:
            logging.error(f"住所選択処理中にエラー: {str(e)}")
            return {"status": "error", "message": f"住所選択処理中にエラーが発生しました: {str(e)}"}
            
    except Exception as e:
        logging.error(f"検索処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"検索処理中にエラーが発生しました: {str(e)}"}
        
    finally:
        # ブラウザを終了
        if driver and auto_close:
            try:
                driver.quit()
                logging.info("ブラウザを終了しました")
            except Exception as e:
                logging.warning(f"ブラウザの終了中にエラー: {str(e)}") 