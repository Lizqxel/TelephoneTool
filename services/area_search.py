"""
提供エリア検索サービス

このモジュールは、NTT西日本の提供エリア検索を
自動化するための機能を提供します。
"""

import logging
import time
import re
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

from services.web_driver import create_driver
from utils.string_utils import normalize_string, calculate_similarity


def search_service_area(postal_code, address):
    """
    NTT西日本の提供エリア検索を実行する関数
    
    Args:
        postal_code (str): 郵便番号
        address (str): 住所
        
    Returns:
        dict: 検索結果を含む辞書
    """
    logging.info(f"郵便番号 {postal_code}、住所 {address} の処理を開始します")
    
    # 住所を分割（都道府県、市区町村、町名、番地・号）
    address = address.replace('　', ' ')  # 全角スペースを半角に統一
    
    # 番地・号を抽出するための正規表現
    number_pattern = r'([0-9０-９]+(?:[-ー－]+[0-9０-９]+)*)'
    match = re.search(number_pattern, address)
    
    street_number = None
    building_number = None
    base_address = address
    
    if match:
        # 番地・号部分を取得
        numbers = match.group(1)
        # 番地部分より前を基本住所とする
        base_address = address[:match.start()].strip()
        
        # 番地と号を分離
        number_parts = numbers.replace('－', '-').replace('ー', '-').split('-')
        if len(number_parts) >= 1:
            street_number = number_parts[0]
            if len(number_parts) >= 2:
                building_number = number_parts[1]
    
    logging.info(f"住所分割結果 - 基本住所: {base_address}, 番地: {street_number}, 号: {building_number}")
    
    # 郵便番号のフォーマットチェック
    postal_code_clean = postal_code.replace("-", "")
    if len(postal_code_clean) != 7 or not postal_code_clean.isdigit():
        return {"status": "error", "message": "郵便番号は7桁の数字で入力してください。"}
    
    driver = None
    try:
        # WebDriverの作成とサイトアクセス
        driver = create_driver()
        driver.get("https://flets-w.com/cart/")
        
        # 郵便番号入力
        try:
            zip_field = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_tx_ybk_yb']"))
            )
            zip_field.clear()
            zip_field.send_keys(postal_code_clean)
            logging.info(f"郵便番号 {postal_code_clean} を入力しました")
        except Exception as e:
            logging.error(f"郵便番号入力に失敗: {str(e)}")
            raise
        
        # 検索ボタンクリック
        try:
            search_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_bt_ybk_jks']"))
            )
            search_button.click()
            logging.info("検索ボタンをクリックしました")
        except Exception as e:
            logging.error(f"検索ボタンクリックに失敗: {str(e)}")
            raise
        
        # 住所選択
        try:
            # 住所選択モーダルの表示待ち
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, "addressSelectModal"))
            )
            time.sleep(1)
            
            # 住所候補の取得
            candidates = driver.find_elements(By.CSS_SELECTOR, "#addressSelectModal .modal-body div")
            valid_candidates = [c for c in candidates if c.text.strip()]
            
            if not valid_candidates:
                raise NoSuchElementException("住所候補が見つかりません")
            
            # 住所の選択
            normalized_input = normalize_string(base_address)
            best_match = None
            
            for candidate in valid_candidates:
                candidate_text = candidate.text.strip().split('\n')[0]
                if normalize_string(candidate_text) == normalized_input:
                    best_match = candidate
                    break
            
            if not best_match:
                raise ValueError("一致する住所が見つかりません")
            
            # 住所クリック
            try:
                best_match.click()
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", best_match)
                except Exception:
                    ActionChains(driver).move_to_element(best_match).click().perform()
            
            time.sleep(2)
            
        except Exception as e:
            logging.error(f"住所選択に失敗: {str(e)}")
            raise
        
        # 番地入力
        if street_number:
            try:
                street_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "id_tak_tx_bnc_bn"))
                )
                street_input.clear()
                street_input.send_keys(street_number)
                
                confirm_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "id_tak_bt_bnc_jks"))
                )
                confirm_button.click()
                
            except Exception as e:
                logging.error(f"番地入力に失敗: {str(e)}")
                raise
        
        # 号入力
        if building_number:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "DIALOG_ID02"))
                )
                
                building_input = driver.find_element(By.ID, "id_tak_tx_gou_gou")
                building_input.clear()
                building_input.send_keys(building_number)
                
                confirm_button = driver.find_element(By.ID, "id_tak_bt_gou_jks")
                confirm_button.click()
                
                WebDriverWait(driver, 10).until(
                    EC.invisibility_of_element_located((By.ID, "DIALOG_ID02"))
                )
                
                time.sleep(2)
                
            except Exception as e:
                logging.error(f"号入力に失敗: {str(e)}")
                raise
        
        # 結果の取得
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "result-message"))
            )
            
            result_elements = driver.find_elements(By.CLASS_NAME, "result-message")
            result_messages = [elem.text.strip() for elem in result_elements if elem.text.strip()]
            
            if result_messages:
                screenshot_path = "area_search_result.png"
                driver.save_screenshot(screenshot_path)
                
                return {
                    "status": "success",
                    "message": result_messages[0],
                    "screenshot_path": screenshot_path
                }
            else:
                return {
                    "status": "error",
                    "message": "結果を取得できませんでした"
                }
                
        except Exception as e:
            logging.error(f"結果取得に失敗: {str(e)}")
            return {
                "status": "error",
                "message": f"結果取得に失敗: {str(e)}"
            }
            
    except Exception as e:
        logging.error(f"処理中にエラー: {str(e)}")
        if driver:
            driver.save_screenshot("error_screenshot.png")
        return {
            "status": "error",
            "message": f"エラーが発生しました: {str(e)}"
        }
        
    finally:
        if driver:
            driver.quit()
            
    return {
        "status": "success",
        "message": "検索が完了しました",
        "details": {}
    } 