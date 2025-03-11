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
                # 番地入力ダイアログが表示されるまで待機
                banchi_dialog = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID01"))
                )
                logging.info("番地入力ダイアログが表示されました")
                
                # ダイアログのHTML構造を詳細にログ出力
                dialog_html = banchi_dialog.get_attribute('outerHTML')
                logging.info(f"番地入力ダイアログのHTML構造:\n{dialog_html}")
                
                # iframeが存在する可能性があるため、確認
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes:
                    logging.info(f"ページ内のiframe数: {len(iframes)}")
                    for i, iframe in enumerate(iframes):
                        try:
                            iframe_id = iframe.get_attribute('id')
                            iframe_src = iframe.get_attribute('src')
                            logging.info(f"iframe {i+1}: ID={iframe_id}, src={iframe_src}")
                            
                            # iframeの中身を確認
                            driver.switch_to.frame(iframe)
                            iframe_content = driver.page_source
                            logging.info(f"iframe {i+1} の内容:\n{iframe_content[:1000]}...")  # 最初の1000文字のみ表示
                            driver.switch_to.default_content()
                        except Exception as e:
                            logging.warning(f"iframe {i+1} の情報取得中にエラー: {str(e)}")
                
                # スクリーンショットを撮影（ダイアログ表示直後）
                driver.save_screenshot("debug_banchi_dialog.png")
                logging.info("番地入力ダイアログのスクリーンショットを保存しました")
                
                # 番地選択のUIから直接番地を選択
                try:
                    # 番地のボタンを探す（複数のセレクタを試行）
                    input_street_number = street_number if street_number else "3"
                    button_selectors = [
                        "#DIALOG_ID01 button",
                        "#DIALOG_ID01 .btn",
                        "#DIALOG_ID01 [role='button']",
                        "#DIALOG_ID01 div[onclick]",
                        "//div[@id='DIALOG_ID01']//div[contains(@class, 'clickable')]",
                        "//div[@id='DIALOG_ID01']//div[not(contains(@class, 'header')) and not(contains(@class, 'footer'))]",
                        "//div[@id='scrollBoxDIALOG_ID01']//a",  # 番地候補リンク
                        "//div[@id='DIALOG_ID01']//div[contains(@class, 'number')]",  # 番地ボタンの可能性がある要素
                        "//div[@id='DIALOG_ID01']//div[contains(@class, 'banchi')]",  # 番地関連の要素
                        "//div[@id='DIALOG_ID01']//div[not(ancestor::div[contains(@class, 'header')]) and not(ancestor::div[contains(@class, 'footer')])]"  # ヘッダーとフッター以外の全div
                    ]
                    
                    # ページ全体のHTMLを取得してログ出力
                    page_html = driver.page_source
                    logging.info(f"ページのHTML構造:\n{page_html}")
                    
                    banchi_buttons = []
                    for selector in button_selectors:
                        try:
                            if selector.startswith('//'):
                                elements = driver.find_elements(By.XPATH, selector)
                            else:
                                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                            if elements:
                                banchi_buttons.extend(elements)
                                logging.info(f"セレクタ '{selector}' で {len(elements)} 個の要素が見つかりました")
                                
                                # 各要素の詳細情報をログ出力
                                for element in elements:
                                    try:
                                        element_html = element.get_attribute('outerHTML')
                                        element_text = element.text.strip()
                                        element_classes = element.get_attribute('class')
                                        element_style = element.get_attribute('style')
                                        element_onclick = element.get_attribute('onclick')
                                        element_role = element.get_attribute('role')
                                        
                                        logging.info(f"""要素の詳細:
                                        HTML: {element_html}
                                        テキスト: {element_text}
                                        クラス: {element_classes}
                                        スタイル: {element_style}
                                        onclick: {element_onclick}
                                        role: {element_role}
                                        """)
                                    except Exception as e:
                                        logging.warning(f"要素の詳細情報取得中にエラー: {str(e)}")
                        except Exception as e:
                            logging.warning(f"セレクタ '{selector}' での検索中にエラー: {str(e)}")
                    
                    # 重複を除去
                    banchi_buttons = list(set(banchi_buttons))
                    
                    # 各ボタンの詳細情報をログ出力
                    for button in banchi_buttons:
                        try:
                            text = button.text.strip()
                            html = button.get_attribute('outerHTML')
                            classes = button.get_attribute('class')
                            is_displayed = button.is_displayed()
                            is_enabled = button.is_enabled()
                            
                            logging.info(f"""番地ボタン詳細:
                            テキスト: '{text}'
                            クラス: {classes}
                            表示状態: {is_displayed}
                            有効状態: {is_enabled}
                            HTML: {html}
                            """)
                        except Exception as e:
                            logging.warning(f"ボタン情報取得中にエラー: {str(e)}")
                    
                    # 入力したい番地と一致するボタンを探す
                    best_match_button = None
                    highest_similarity = 0
                    
                    for button in banchi_buttons:
                        try:
                            if not button.is_displayed() or not button.is_enabled():
                                continue
                                
                            button_text = button.text.strip()
                            if not button_text:
                                # テキストが空の場合、div内のテキストを探す
                                button_text = button.get_attribute('textContent').strip()
                            
                            if not button_text:
                                continue
                                
                            similarity = calculate_similarity(
                                normalize_string(input_street_number),
                                normalize_string(button_text)
                            )
                            
                            logging.info(f"番地ボタン '{button_text}' の類似度: {similarity}")
                            
                            if similarity > highest_similarity:
                                highest_similarity = similarity
                                best_match_button = button
                        except Exception as e:
                            logging.warning(f"ボタン類似度計算中にエラー: {str(e)}")
                    
                    if best_match_button:
                        # ボタンが見つかった場合、クリックを試みる
                        button_text = best_match_button.text.strip()
                        logging.info(f"番地ボタン '{button_text}' を選択します")
                        
                        try:
                            # スクロールしてボタンを表示
                            driver.execute_script("arguments[0].scrollIntoView(true);", best_match_button)
                            time.sleep(1)
                            
                            # クリックを試行（複数の方法）
                            try:
                                best_match_button.click()
                                logging.info("通常のクリックで番地を選択しました")
                            except Exception as click_error:
                                logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                try:
                                    driver.execute_script("arguments[0].click();", best_match_button)
                                    logging.info("JavaScriptでクリックしました")
                                except Exception as js_error:
                                    logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                    ActionChains(driver).move_to_element(best_match_button).click().perform()
                                    logging.info("ActionChainsでクリックしました")
                            
                            # クリック後の待機
                            time.sleep(2)
                            
                        except Exception as e:
                            logging.error(f"番地ボタンのクリックに失敗: {str(e)}")
                            raise
                    else:
                        # ボタンが見つからない場合は、入力フィールドを探す（フォールバック）
                        logging.warning("一致する番地ボタンが見つからないため、入力フィールドを使用します")
                        
                        # 番地入力フィールドの検出を試みる（複数のセレクタを使用）
                        selectors = [
                            "//input[@type='text' and ancestor::*[@id='DIALOG_ID01']]",
                            "//*[@id='DIALOG_ID01']//input",
                            "//input[contains(@class, 'banchi-input')]",
                            "//*[@id='DIALOG_ID01']/div/div[2]/div[1]/input"
                        ]
                        
                        banchi_field = None
                        for selector in selectors:
                            try:
                                banchi_field = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, selector))
                                )
                                logging.info(f"番地入力フィールドが見つかりました: {selector}")
                                break
                            except:
                                continue
                        
                        if banchi_field:
                            # 番地入力フィールドをクリアして入力
                            banchi_field.clear()
                            banchi_field.send_keys(input_street_number)
                            logging.info(f"番地「{input_street_number}」を入力しました")
                            time.sleep(1)
                            banchi_field.send_keys(Keys.RETURN)
                            logging.info("Enterキーを送信しました")
                        else:
                            raise NoSuchElementException("番地入力フィールドが見つかりませんでした")
                
                except Exception as e:
                    logging.error(f"番地選択処理中にエラーが発生: {str(e)}")
                    driver.save_screenshot("debug_banchi_error.png")
                    logging.info("エラー発生時のスクリーンショットを保存しました")
                    raise
            
            except TimeoutException as e:
                logging.error(f"番地入力処理でタイムアウトが発生しました: {str(e)}")
                driver.save_screenshot("debug_banchi_timeout.png")
                logging.info("タイムアウト時のスクリーンショットを保存しました")
                raise
            
            except Exception as e:
                logging.error(f"番地入力処理中にエラーが発生しました: {str(e)}")
                driver.save_screenshot("debug_banchi_error.png")
                logging.info("エラー発生時のスクリーンショットを保存しました")
                raise
            
            # 6. 号入力画面が表示された場合は、最初の候補を選択
            try:
                # 号入力ダイアログが表示されるまで待機（ID指定）
                gou_dialog = WebDriverWait(driver, 15).until(
                    EC.visibility_of_element_located((By.ID, "DIALOG_ID02"))
                )
                logging.info("号入力ダイアログが表示されました")
                
                # スクリーンショットを撮影（ダイアログ表示直後）
                driver.save_screenshot("debug_gou_dialog.png")
                logging.info("号入力ダイアログのスクリーンショットを保存しました")
                
                # ダイアログのHTML構造を詳細にログ出力
                dialog_html = gou_dialog.get_attribute('outerHTML')
                logging.info(f"号入力ダイアログのHTML構造:\n{dialog_html}")
                
                # 号選択のUIから直接号を選択
                try:
                    # 号のボタンを探す（複数のセレクタを試行）
                    input_building_number = building_number if building_number else "7"
                    button_selectors = [
                        "#DIALOG_ID02 button",
                        "#DIALOG_ID02 .btn",
                        "#DIALOG_ID02 [role='button']",
                        "#DIALOG_ID02 div[onclick]",
                        "//div[@id='DIALOG_ID02']//div[contains(@class, 'clickable')]",
                        "//div[@id='DIALOG_ID02']//div[not(contains(@class, 'header')) and not(contains(@class, 'footer'))]",
                        "//div[@id='scrollBoxDIALOG_ID02']//a",  # 号候補リンク
                        "//div[@id='DIALOG_ID02']//div[contains(@class, 'number')]",  # 号ボタンの可能性がある要素
                        "//div[@id='DIALOG_ID02']//div[contains(@class, 'gou')]",  # 号関連の要素
                        "//div[@id='DIALOG_ID02']//div[not(ancestor::div[contains(@class, 'header')]) and not(ancestor::div[contains(@class, 'footer')])]"  # ヘッダーとフッター以外の全div
                    ]
                    
                    gou_buttons = []
                    for selector in button_selectors:
                        try:
                            if selector.startswith('//'):
                                elements = driver.find_elements(By.XPATH, selector)
                            else:
                                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                            if elements:
                                gou_buttons.extend(elements)
                                logging.info(f"セレクタ '{selector}' で {len(elements)} 個の要素が見つかりました")
                                
                                # 各要素の詳細情報をログ出力
                                for element in elements:
                                    try:
                                        element_html = element.get_attribute('outerHTML')
                                        element_text = element.text.strip()
                                        element_classes = element.get_attribute('class')
                                        element_style = element.get_attribute('style')
                                        element_onclick = element.get_attribute('onclick')
                                        element_role = element.get_attribute('role')
                                        
                                        logging.info(f"""号要素の詳細:
                                        HTML: {element_html}
                                        テキスト: {element_text}
                                        クラス: {element_classes}
                                        スタイル: {element_style}
                                        onclick: {element_onclick}
                                        role: {element_role}
                                        """)
                                    except Exception as e:
                                        logging.warning(f"要素の詳細情報取得中にエラー: {str(e)}")
                        except Exception as e:
                            logging.warning(f"セレクタ '{selector}' での検索中にエラー: {str(e)}")
                    
                    # 重複を除去
                    gou_buttons = list(set(gou_buttons))
                    
                    # 入力したい号と一致するボタンを探す
                    best_match_button = None
                    highest_similarity = 0
                    
                    for button in gou_buttons:
                        try:
                            if not button.is_displayed() or not button.is_enabled():
                                continue
                                
                            button_text = button.text.strip()
                            if not button_text:
                                # テキストが空の場合、div内のテキストを探す
                                button_text = button.get_attribute('textContent').strip()
                            
                            if not button_text:
                                continue
                                
                            similarity = calculate_similarity(
                                normalize_string(input_building_number),
                                normalize_string(button_text)
                            )
                            
                            logging.info(f"号ボタン '{button_text}' の類似度: {similarity}")
                            
                            if similarity > highest_similarity:
                                highest_similarity = similarity
                                best_match_button = button
                        except Exception as e:
                            logging.warning(f"ボタン類似度計算中にエラー: {str(e)}")
                    
                    if best_match_button:
                        # ボタンが見つかった場合、クリックを試みる
                        button_text = best_match_button.text.strip()
                        logging.info(f"号ボタン '{button_text}' を選択します")
                        
                        try:
                            # スクロールしてボタンを表示
                            driver.execute_script("arguments[0].scrollIntoView(true);", best_match_button)
                            time.sleep(1)
                            
                            # クリックを試行（複数の方法）
                            try:
                                best_match_button.click()
                                logging.info("通常のクリックで号を選択しました")
                            except Exception as click_error:
                                logging.warning(f"通常のクリックに失敗: {str(click_error)}")
                                try:
                                    driver.execute_script("arguments[0].click();", best_match_button)
                                    logging.info("JavaScriptでクリックしました")
                                except Exception as js_error:
                                    logging.warning(f"JavaScriptクリックに失敗: {str(js_error)}")
                                    ActionChains(driver).move_to_element(best_match_button).click().perform()
                                    logging.info("ActionChainsでクリックしました")
                            
                            # クリック後の待機
                            time.sleep(2)
                            
                        except Exception as e:
                            logging.error(f"号ボタンのクリックに失敗: {str(e)}")
                            raise
                    else:
                        # ボタンが見つからない場合は、入力フィールドを探す（フォールバック）
                        logging.warning("一致する号ボタンが見つからないため、入力フィールドを使用します")
                        
                        # 号入力フィールドの検出を試みる（複数のセレクタを使用）
                        selectors = [
                            "//input[@type='text' and ancestor::*[@id='DIALOG_ID02']]",
                            "//*[@id='DIALOG_ID02']//input",
                            "//input[contains(@class, 'gou-input')]",
                            "//*[@id='DIALOG_ID02']/div/div[2]/div[1]/input"
                        ]
                        
                        gou_field = None
                        for selector in selectors:
                            try:
                                gou_field = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, selector))
                                )
                                logging.info(f"号入力フィールドが見つかりました: {selector}")
                                break
                            except:
                                continue
                        
                        if gou_field:
                            # 号入力フィールドをクリアして入力
                            gou_field.clear()
                            gou_field.send_keys(input_building_number)
                            logging.info(f"号「{input_building_number}」を入力しました")
                            time.sleep(1)
                            gou_field.send_keys(Keys.RETURN)
                            logging.info("Enterキーを送信しました")
                        else:
                            raise NoSuchElementException("号入力フィールドが見つかりませんでした")
                
                except Exception as e:
                    logging.error(f"号選択処理中にエラーが発生: {str(e)}")
                    driver.save_screenshot("debug_gou_error.png")
                    logging.info("エラー発生時のスクリーンショットを保存しました")
                    raise
                    
                    # 号選択後の読み込みを待つ
                WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.ID, "DIALOG_ID02"))
                    )
                logging.info("号選択ダイアログが閉じられました")
                
                # 号選択後の画面状態を確認
                time.sleep(2)  # 画面の遷移を待つ
                driver.save_screenshot("debug_after_gou_selection.png")
                logging.info("号選択後の画面状態をスクリーンショットとして保存しました")
                
                # 画面全体のHTMLを取得してログ出力
                page_html = driver.page_source
                logging.info(f"号選択後の画面のHTML構造:\n{page_html}")
                
            except TimeoutException:
                # 号入力画面が表示されない場合はスキップ
                logging.info("号入力画面はスキップされました")
            except Exception as e:
                logging.error(f"号入力処理中にエラーが発生しました: {str(e)}")
                driver.save_screenshot("debug_gou_error.png")
                logging.info("エラー発生時のスクリーンショットを保存しました")
                # エラーが発生しても処理を継続
                pass
            
            # 7. 結果の判定を改善
            try:
                # 結果判定の前にスクリーンショットを撮影
                driver.save_screenshot("debug_result_screen.png")
                logging.info("結果画面のスクリーンショットを保存しました")
                
                # ページ全体のテキストを取得して判断
                page_text = driver.page_source.lower()
                
                # 判定条件を詳細に設定
                availability_indicators = [
                    "ご利用いただけます",
                    "提供可能",
                    "サービスのご利用が可能",
                    "お申し込みいただけます"
                ]
                
                unavailability_indicators = [
                    "ご利用いただけません",
                    "提供不可",
                    "サービスのご利用ができません",
                    "お申し込みいただけません"
                ]
                
                # 利用可能性の判定
                is_available = any(indicator in page_text for indicator in availability_indicators)
                is_unavailable = any(indicator in page_text for indicator in unavailability_indicators)
                
                if is_available and not is_unavailable:
                    logging.info("提供可能と判定されました（テキストベース）")
                    return {"status": "success", "message": "提供可能"}
                elif is_unavailable:
                    logging.info("提供不可と判定されました（テキストベース）")
                    return {"status": "error", "message": "提供不可"}
                else:
                    logging.warning("判定結果が不明確です")
                    return {"status": "error", "message": "判定結果が不明確です"}
                    
            except Exception as e:
                logging.error(f"結果の判定中にエラー: {str(e)}")
                driver.save_screenshot("debug_result_error.png")
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

        # 9. 最終検索ボタンをクリック
        try:
            logging.info("検索結果確認ボタンの検出を開始します")
            
            # 検索結果確認ボタンクリック前のスクリーンショット
            driver.save_screenshot("debug_before_search_confirm.png")
            logging.info("検索結果確認ボタンクリック前のスクリーンショットを保存しました")
            
            # 最終検索ボタンの検出を試みる（複数のセレクタを使用）
            button_selectors = [
                "//button[contains(text(), '検索結果を確認')]",
                "//div[contains(@class, 'search-confirm')]//button",
                "//*[@id='id_tak_bt_nx']",
                "//button[contains(@class, 'next')]",
                "//button[contains(text(), '次へ')]"
            ]
            
            # ボタン検出の待機時間を短縮
            final_search_button = None
            for selector in button_selectors:
                try:
                    element = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    if element.is_displayed() and element.is_enabled():
                        final_search_button = element
                        logging.info(f"検索結果確認ボタンが見つかりました: {selector}")
                        break
                except:
                    continue
            
            if not final_search_button:
                raise NoSuchElementException("検索結果確認ボタンが見つかりませんでした")
            
            # ボタンをクリック
            try:
                # スクロールしてボタンを表示
                driver.execute_script("arguments[0].scrollIntoView(true);", final_search_button)
                time.sleep(0.5)
                
                # クリックを実行
                final_search_button.click()
                logging.info("検索結果確認ボタンをクリックしました")
                
                # クリック後のスクリーンショット
                time.sleep(0.5)
                driver.save_screenshot("debug_after_search_confirm.png")
                logging.info("検索結果確認ボタンクリック後のスクリーンショットを保存しました")
                
            except Exception as e:
                logging.error(f"検索結果確認ボタンのクリックに失敗: {str(e)}")
                raise
            
            # クリック後の画面遷移を待機（タイムアウトを短縮）
            try:
                result_text_selectors = [
                    "//*[contains(text(), 'ご利用いただけます')]",
                    "//*[contains(text(), 'ご利用いただけません')]",
                    "//*[contains(text(), '提供可能')]",
                    "//*[contains(text(), '提供不可')]"
                ]
                
                # より短いタイムアウトで結果を待機
                for selector in result_text_selectors:
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                        logging.info(f"結果テキストが見つかりました: {selector}")
                        break
                    except:
                        continue
            
            except TimeoutException:
                logging.warning("結果テキストの待機中にタイムアウトが発生しました")
                driver.save_screenshot("debug_result_timeout.png")
                logging.info("タイムアウト時のスクリーンショットを保存しました")
            
        except Exception as e:
            logging.error(f"検索結果確認ボタンの操作に失敗: {str(e)}")
            driver.save_screenshot("debug_search_confirm_error.png")
            logging.info("エラー時のスクリーンショットを保存しました")
            raise 