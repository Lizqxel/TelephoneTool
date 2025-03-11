"""
提供エリア検索サービス

このモジュールは、NTT西日本の提供エリア検索を
自動化するための機能を提供します。
"""

import logging
import time
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

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
    
    # 郵便番号のフォーマットチェック
    postal_code_clean = postal_code.replace("-", "")
    if len(postal_code_clean) != 7 or not postal_code_clean.isdigit():
        return {"status": "error", "message": "郵便番号は7桁の数字で入力してください。"}
    
    driver = None
    try:
        # 1. ドライバーを作成してサイトを開く
        driver = create_driver()
        driver.implicitly_wait(0)  # 暗黙の待機を無効化
        driver.get("https://flets-w.com/cart/")
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
            # ページのHTMLを出力してデバッグ
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
            
            # 方法1: XPathで直接取得
            try:
                candidate_container = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='addressSelectModal']/div/div[2]"))
                )
                candidates = candidate_container.find_elements(By.TAG_NAME, "div")
                logging.info(f"方法1: {len(candidates)} 件の候補が見つかりました")
            except Exception as e:
                logging.warning(f"方法1での候補取得に失敗: {str(e)}")
            
            # 方法2: JavaScriptで取得
            if not candidates or all(c.text.strip() == "" for c in candidates):
                try:
                    js_candidates = driver.execute_script("""
                        return Array.from(document.querySelectorAll('#addressSelectModal .modal-body div')).filter(el => el.textContent.trim() !== '');
                    """)
                    if js_candidates:
                        candidates = js_candidates
                        logging.info(f"方法2: {len(candidates)} 件の候補が見つかりました")
                except Exception as e:
                    logging.warning(f"方法2での候補取得に失敗: {str(e)}")
            
            # 方法3: CSSセレクタで取得
            if not candidates or all(c.text.strip() == "" for c in candidates):
                try:
                    css_candidates = driver.find_elements(By.CSS_SELECTOR, "#addressSelectModal .modal-body div")
                    if css_candidates:
                        candidates = css_candidates
                        logging.info(f"方法3: {len(candidates)} 件の候補が見つかりました")
                except Exception as e:
                    logging.warning(f"方法3での候補取得に失敗: {str(e)}")
            
            # 有効な候補（テキストが空でない）をフィルタリング
            valid_candidates = [c for c in candidates if c.text.strip()]
            logging.info(f"有効な候補: {len(valid_candidates)} 件")
            
            if not valid_candidates:
                # 候補が見つからない場合は、モーダル内の全要素を取得して調査
                try:
                    all_elements = driver.find_elements(By.XPATH, "//*[@id='addressSelectModal']//*")
                    logging.info(f"モーダル内の全要素数: {len(all_elements)}")
                    
                    # テキストを持つ要素を探す
                    text_elements = [e for e in all_elements if e.text.strip()]
                    logging.info(f"テキストを持つ要素数: {len(text_elements)}")
                    
                    for i, elem in enumerate(text_elements[:5]):  # 最初の5つだけログ出力
                        logging.info(f"要素 {i+1}: タグ={elem.tag_name}, テキスト='{elem.text.strip()}'")
                    
                    # クリック可能な要素を候補として使用
                    valid_candidates = [e for e in text_elements if e.is_displayed() and e.is_enabled()]
                    logging.info(f"クリック可能な要素数: {len(valid_candidates)}")
                except Exception as e:
                    logging.warning(f"モーダル内の全要素取得に失敗: {str(e)}")
            
            if not valid_candidates:
                raise NoSuchElementException("有効な住所候補が見つかりませんでした")
            
            # 各候補のテキストをログに出力
            for i, candidate in enumerate(valid_candidates[:5]):  # 最初の5つだけログ出力
                candidate_text = candidate.text.strip()
                logging.info(f"候補 {i+1}: '{candidate_text}'")
            
            # 4. 住所を選択（一致するもの、なければ近似値）
            best_candidate = max(
                valid_candidates,
                key=lambda c: calculate_similarity(
                    normalize_string(address),
                    normalize_string(c.text.strip())
                ),
                default=None
            )
            
            if not best_candidate:
                raise NoSuchElementException("一致する住所が見つかりませんでした")
            
            logging.info(f"最適な住所候補: '{best_candidate.text.strip()}'")
            
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
            
            logging.info(f"選択された住所: {best_candidate.text.strip()}")
            
            # 住所選択後の読み込みを待つ
            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element_located((By.ID, "addressSelectModal"))
            )
            logging.info("住所選択モーダルが閉じられました")
            
            # 5. 番地入力画面が表示された場合は、最初の候補を選択
            try:
                # 番地入力ダイアログが表示されるまで待機（ID指定）
                banchi_dialog = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID01"))
                )
                logging.info("番地入力ダイアログが表示されました")
                
                # 番地入力フィールドを探す
                banchi_field = banchi_dialog.find_element(By.XPATH, ".//div/div[2]/div[1]/input")
                banchi_field.clear()
                banchi_field.send_keys("1")  # 仮の番地として1を入力
                logging.info("番地「1」を入力しました")
                
                # 候補リストが表示されるのを待つ
                candidate_list = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='scrollBoxDIALOG_ID01']/ul/li[1]/a"))
                )
                
                # 最初の候補をクリック
                candidate_list.click()
                logging.info(f"番地候補を選択しました: {candidate_list.text}")
                
                # 番地選択後の読み込みを待つ
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.ID, "DIALOG_ID01"))
                )
            except TimeoutException:
                # 番地入力画面が表示されない場合はスキップ
                logging.info("番地入力画面はスキップされました")
            
            # 6. 号入力画面が表示された場合は、最初の候補を選択
            try:
                # 号入力ダイアログが表示されるまで待機（ID指定）
                gou_dialog = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID02"))
                )
                logging.info("号入力ダイアログが表示されました")
                
                # 最初の候補を探して選択
                gou_candidates = gou_dialog.find_elements(By.XPATH, ".//div/div[2]/div")
                
                if gou_candidates:
                    first_gou = gou_candidates[0]
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable(first_gou))
                    first_gou.click()
                    logging.info(f"号候補を選択しました: {first_gou.text}")
                    
                    # 号選択後の読み込みを待つ
                    WebDriverWait(driver, 5).until(
                        EC.invisibility_of_element_located((By.ID, "DIALOG_ID02"))
                    )
            except TimeoutException:
                # 号入力画面が表示されない場合はスキップ
                logging.info("号入力画面はスキップされました")
            
            # 7. 検索ボタンを押す
            try:
                final_search_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='id_tak_bt_nx']"))
                )
                final_search_button.click()
                logging.info("最終検索ボタンをクリックしました")
            except Exception as e:
                logging.error(f"最終検索ボタンが見つかりませんでした: {str(e)}")
                # ページのHTMLを出力してデバッグ
                logging.info(f"ページのHTML: {driver.page_source[:500]}...")
                raise
            
            # 結果が表示されるのを待つ
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='nextForm']/div/div[2]/div/picture/img"))
                )
                logging.info("結果画像が表示されました")
            except Exception as e:
                logging.error(f"結果画像が見つかりませんでした: {str(e)}")
                # ページのHTMLを出力してデバッグ
                logging.info(f"ページのHTML: {driver.page_source[:500]}...")
                # 画像が見つからなくても続行
            
            # 8. 結果を確認
            try:
                # 画像のsrc属性を確認
                try:
                    image_element = driver.find_element(By.XPATH, "//*[@id='nextForm']/div/div[2]/div/picture/img")
                    image_src = image_element.get_attribute("src")
                    logging.info(f"結果画像のsrc: {image_src}")
                    
                    if "ok" in image_src.lower():
                        logging.info("提供可能と判定されました（画像src）")
                        return {"status": "success", "message": "提供可能"}
                except Exception as e:
                    logging.info(f"画像要素が見つからないか、src属性の取得に失敗しました: {str(e)}")
                
                # ページ全体のテキストを取得して判断
                page_text = driver.page_source.lower()
                
                if "ご利用いただけます" in page_text or "提供可能" in page_text or "ok" in page_text:
                    logging.info("提供可能と判定されました（ページテキスト）")
                    return {"status": "success", "message": "提供可能"}
                else:
                    logging.info("提供不可と判定されました")
                    return {"status": "error", "message": "提供不可"}
            except Exception as e:
                logging.error(f"結果の判定に失敗しました: {str(e)}")
                return {"status": "error", "message": f"結果の判定に失敗しました: {str(e)}"}
            
        except TimeoutException as e:
            logging.error(f"住所候補の表示待ちでタイムアウトしました: {str(e)}")
            # ページのHTMLを出力してデバッグ
            logging.info(f"ページのHTML: {driver.page_source[:500]}...")
            return {"status": "error", "message": "住所候補が見つかりませんでした"}
        
    except TimeoutException as e:
        logging.error(f"タイムアウトが発生しました: {str(e)}")
        return {"status": "error", "message": f"処理中にタイムアウトが発生しました"}
        
    except NoSuchElementException as e:
        logging.error(f"要素が見つかりませんでした: {str(e)}")
        return {"status": "error", "message": f"必要な要素が見つかりませんでした"}
        
    except Exception as e:
        logging.error(f"自動化に失敗しました: {str(e)}")
        return {"status": "error", "message": f"エラーが発生しました: {str(e)}"}
        
    finally:
        # ドライバーを閉じる
        if driver:
            driver.quit() 