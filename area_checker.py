"""
NTT西日本フレッツ光の提供エリア確認モジュール

このモジュールは、NTT西日本のフレッツ光提供エリア確認サイトにアクセスして、
住所から提供エリア判定を行う機能を提供します。
"""

import time
import json
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

# ロギングの設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AreaChecker:
    """NTT西日本フレッツ光の提供エリア確認クラス"""
    
    def __init__(self):
        """初期化"""
        self.driver = None
        
        # 都道府県コード（NTT西日本エリア）
        self.prefecture_dict = {
            "三重県": "24",
            "滋賀県": "25",
            "京都府": "26",
            "大阪府": "27",
            "兵庫県": "28",
            "奈良県": "29",
            "和歌山県": "30",
            "鳥取県": "31",
            "島根県": "32",
            "岡山県": "33",
            "広島県": "34",
            "山口県": "35",
            "徳島県": "36",
            "香川県": "37",
            "愛媛県": "38",
            "高知県": "39",
            "福岡県": "40",
            "佐賀県": "41",
            "長崎県": "42",
            "熊本県": "43",
            "大分県": "44",
            "宮崎県": "45",
            "鹿児島県": "46",
            "沖縄県": "47"
        }
    
    def setup_driver(self):
        """Seleniumドライバーのセットアップ"""
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # ヘッドレスモードで実行
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--window-size=1920,1080')  # ウィンドウサイズを設定
        chrome_options.add_argument('--start-maximized')  # ウィンドウを最大化
        chrome_options.add_argument('--disable-extensions')  # 拡張機能を無効化
        chrome_options.add_argument('--disable-popup-blocking')  # ポップアップブロックを無効化
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')  # 自動化検出を回避
        chrome_options.add_argument('--lang=ja')  # 言語を日本語に設定
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')  # ユーザーエージェントを設定
        
        # ページロード戦略を設定
        chrome_options.page_load_strategy = 'eager'  # DOMの読み込みが完了したら次に進む
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # タイムアウト設定
        driver.set_page_load_timeout(30)  # ページロードのタイムアウトを30秒に設定
        driver.implicitly_wait(10)  # 暗黙的な待機を10秒に設定
        
        return driver
    
    def check_area(self, prefecture, city, town=None, block=None):
        """
        提供エリアを確認する
        
        Args:
            prefecture (str): 都道府県名
            city (str): 市区町村名
            town (str, optional): 町名
            block (str, optional): 番地
            
        Returns:
            dict: 提供可否の結果を含む辞書
        """
        try:
            logger.info(f"住所 {prefecture}{city}{town or ''}{block or ''} の提供エリア確認を開始します")
            
            # 府県から検索する方法
            result = self.check_area_by_prefecture(prefecture, city, town, block)
            return result
            
        except Exception as e:
            logger.error(f"エラーが発生しました: {str(e)}")
            return {
                'status': 'ERROR',
                'message': f'エラーが発生しました：{str(e)}'
            }
    
    def check_area_by_prefecture(self, prefecture, city, town=None, block=None):
        """
        府県から検索する方法でエリア確認
        
        Args:
            prefecture (str): 都道府県名
            city (str): 市区町村名
            town (str, optional): 町名
            block (str, optional): 番地
            
        Returns:
            dict: 提供可否の結果を含む辞書
        """
        driver = None
        try:
            logger.info(f"府県から検索する方法でエリア確認を開始: {prefecture}{city}{town or ''}{block or ''}")
            driver = self.setup_driver()
            
            # サイトにアクセス
            url = "https://flets-w.com/cart/"
            logger.info(f"サイトにアクセス: {url}")
            driver.get(url)
            
            # ページの読み込みを待機
            time.sleep(5)
            
            # スクリーンショットを取得して状態を確認
            screenshot_path = "initial_page.png"
            driver.save_screenshot(screenshot_path)
            logger.info(f"初期ページのスクリーンショットを保存しました: {screenshot_path}")
            
            # 「府県から検索」リンクをクリック
            try:
                # 新しいHTMLコードに基づいたセレクタを追加
                selectors = [
                    "a.c-link-arrow01",  # 提供されたHTMLコードに基づく
                    "a.c-link-arrow01[href='']",  # href属性が空の場合
                    "//a[contains(@class, 'c-link-arrow01')]",  # XPath
                    "//a[contains(text(), '府県から検索')]",  # テキストで検索
                    "//a[text()='府県から検索']",  # 完全一致テキスト
                    "a[data-target='#tab2']",  # 以前のセレクタ
                    "a[href='#tab2']",
                    "a.tab-link[data-target='#tab2']",
                    "ul.nav-tabs li:nth-child(2) a",
                    "//a[contains(., '府県から検索')]",
                ]
                
                tab_found = False
                for selector in selectors:
                    try:
                        if selector.startswith("//"):
                            # XPathの場合
                            prefecture_tab = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, selector))
                            )
                        else:
                            # CSSセレクタの場合
                            prefecture_tab = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                            )
                        
                        # スクリーンショットを取得して要素の位置を確認
                        element_screenshot_path = f"element_{selectors.index(selector)}.png"
                        driver.save_screenshot(element_screenshot_path)
                        logger.info(f"要素発見時のスクリーンショットを保存しました: {element_screenshot_path}")
                        
                        # 要素の位置までスクロール
                        driver.execute_script("arguments[0].scrollIntoView(true);", prefecture_tab)
                        time.sleep(1)
                        
                        # JavaScriptを使用してクリック
                        driver.execute_script("arguments[0].click();", prefecture_tab)
                        logger.info(f"「府県から検索」リンクをクリック (セレクタ: {selector})")
                        tab_found = True
                        time.sleep(3)
                        
                        # クリック後のスクリーンショットを取得
                        after_click_screenshot = f"after_click_{selectors.index(selector)}.png"
                        driver.save_screenshot(after_click_screenshot)
                        logger.info(f"クリック後のスクリーンショットを保存しました: {after_click_screenshot}")
                        break
                    except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                        logger.warning(f"セレクタ {selector} でリンクが見つからないかクリックできません: {str(e)}")
                        continue
                
                if not tab_found:
                    # リンクが見つからない場合、ページのHTMLを出力して調査
                    logger.error("すべてのセレクタで「府県から検索」リンクが見つかりませんでした")
                    html_content = driver.page_source
                    with open("page_source.html", "w", encoding="utf-8") as f:
                        f.write(html_content)
                    logger.info("ページのHTMLを page_source.html に保存しました")
                    
                    # 直接郵便番号から検索する方法を試す
                    return self.check_area_by_postal_code(prefecture, city, town, block)
                
                # 府県選択画面が表示されるまで待機
                time.sleep(3)
                
                # 府県選択画面のスクリーンショットを取得
                prefecture_screen_path = "prefecture_screen.png"
                driver.save_screenshot(prefecture_screen_path)
                logger.info(f"府県選択画面のスクリーンショットを保存しました: {prefecture_screen_path}")
                
                # 都道府県選択画面のHTMLを保存
                html_content = driver.page_source
                with open("prefecture_selection_page.html", "w", encoding="utf-8") as f:
                    f.write(html_content)
                logger.info("都道府県選択画面のHTMLを prefecture_selection_page.html に保存しました")
                
                # 全てのリンクを取得して調査
                all_links = driver.find_elements(By.TAG_NAME, "a")
                link_texts = [link.text for link in all_links if link.text]
                logger.info(f"ページ上の全てのリンクテキスト: {link_texts}")
                
                # 都道府県コードの辞書
                prefecture_code_dict = {
                    "富山県": "16", "石川県": "17", "福井県": "18", "岐阜県": "21", "静岡県": "22",
                    "愛知県": "23", "三重県": "24", "滋賀県": "25", "京都府": "26", "大阪府": "27",
                    "兵庫県": "28", "奈良県": "29", "和歌山県": "30", "鳥取県": "31", "島根県": "32",
                    "岡山県": "33", "広島県": "34", "山口県": "35", "徳島県": "36", "香川県": "37",
                    "愛媛県": "38", "高知県": "39", "福岡県": "40", "佐賀県": "41", "長崎県": "42",
                    "熊本県": "43", "大分県": "44", "宮崎県": "45", "鹿児島県": "46", "沖縄県": "47"
                }
                
                # 都道府県コードを取得
                prefecture_code = prefecture_code_dict.get(prefecture)
                if not prefecture_code:
                    logger.error(f"都道府県名 {prefecture} に対応するコードが見つかりません")
                    return {
                        'status': 'ERROR',
                        'message': f'都道府県名 {prefecture} に対応するコードが見つかりません'
                    }
                
                # 都道府県選択ドロップダウンを探す
                try:
                    # セレクタを試す
                    prefecture_select_selectors = [
                        "select[name='tak_cb_fkk_fks']",
                        "#'id_tak_cb_fkk_fks'",  # 実際のHTMLにはシングルクォートが含まれている
                        "select[aria-label='選択してください']",
                        ".c-input-pulldown select",
                        "//select[@name='tak_cb_fkk_fks']",  # XPath
                        "//select[@id='id_tak_cb_fkk_fks']",  # XPath
                        "//select[contains(@id, 'tak_cb_fkk_fks')]",  # XPath
                        "//select",  # すべてのセレクト要素
                    ]
                    
                    prefecture_select = None
                    for selector in prefecture_select_selectors:
                        try:
                            if selector.startswith("//"):
                                # XPathの場合
                                prefecture_select = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, selector))
                                )
                            else:
                                # CSSセレクタの場合
                                prefecture_select = WebDriverWait(driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                                )
                            logger.info(f"都道府県選択フィールドを発見 (セレクタ: {selector})")
                            break
                        except (TimeoutException, NoSuchElementException) as e:
                            logger.warning(f"セレクタ {selector} で都道府県選択フィールドが見つかりません: {str(e)}")
                            continue
                    
                    if not prefecture_select:
                        # すべてのセレクト要素を取得
                        select_elements = driver.find_elements(By.TAG_NAME, "select")
                        if select_elements:
                            prefecture_select = select_elements[0]
                            logger.info("最初のセレクト要素を都道府県選択フィールドとして使用します")
                    
                    if prefecture_select:
                        # セレクト要素を取得できた場合
                        select = Select(prefecture_select)
                        
                        # 選択肢の一覧を取得して記録
                        options = [o.text for o in select.options]
                        logger.info(f"利用可能な都道府県選択肢: {options}")
                        
                        # 都道府県を選択
                        try:
                            select.select_by_value(prefecture_code)
                            logger.info(f"都道府県を選択: {prefecture} (コード: {prefecture_code})")
                            time.sleep(2)
                            
                            # 選択後のスクリーンショットを取得
                            after_prefecture_select = "after_prefecture_select.png"
                            driver.save_screenshot(after_prefecture_select)
                            logger.info(f"都道府県選択後のスクリーンショットを保存しました: {after_prefecture_select}")
                            
                            # 選択後に表示される市区町村一覧を待機
                            time.sleep(3)
                            
                            # 市区町村一覧のスクリーンショットを取得
                            city_list_screen_path = "city_list_screen.png"
                            driver.save_screenshot(city_list_screen_path)
                            logger.info(f"市区町村一覧画面のスクリーンショットを保存しました: {city_list_screen_path}")
                            
                            # 市区町村一覧ページのHTMLを保存
                            html_content = driver.page_source
                            with open("city_list_page.html", "w", encoding="utf-8") as f:
                                f.write(html_content)
                            logger.info("市区町村一覧ページのHTMLを city_list_page.html に保存しました")
                            
                            # 市区町村名のリンクをクリック
                            city_link_selectors = [
                                f"//a[contains(text(), '{city}')]",  # テキストで検索
                                f"//a[text()='{city}']",  # 完全一致テキスト
                                f"//li//a[contains(text(), '{city}')]",
                                f"//li//a[text()='{city}']",
                                f"//ul/li//a[contains(text(), '{city}')]",
                                f"//ul/li//a[text()='{city}']",
                                f"//div[@class='candidate_list_wrap']//ul/li//a[contains(text(), '{city}')]",
                                f"//div[@class='candidate_list_wrap']//ul/li//a[text()='{city}']",
                                f"//div[@id='scrollBoxprefectureAddressSearchModalId']//ul/li//a[contains(text(), '{city}')]",
                                f"//div[@id='scrollBoxprefectureAddressSearchModalId']//ul/li//a[text()='{city}']",
                            ]
                            
                            city_link_found = False
                            for selector in city_link_selectors:
                                try:
                                    logger.info(f"市区町村リンクのセレクタを試行: {selector}")
                                    city_links = driver.find_elements(By.XPATH, selector)
                                    
                                    if not city_links:
                                        logger.warning(f"セレクタ {selector} で市区町村リンクが見つかりません")
                                        continue
                                    
                                    logger.info(f"セレクタ {selector} で {len(city_links)} 個のリンクが見つかりました")
                                    
                                    # 市区町村名が完全に一致するリンクを探す
                                    for city_link in city_links:
                                        link_text = city_link.text
                                        logger.info(f"リンクテキスト: '{link_text}'")
                                        
                                        if city in link_text:
                                            # 要素の位置までスクロール
                                            driver.execute_script("arguments[0].scrollIntoView(true);", city_link)
                                            time.sleep(1)
                                            
                                            # クリック前のスクリーンショットを取得
                                            before_city_click = "before_city_click.png"
                                            driver.save_screenshot(before_city_click)
                                            logger.info(f"市区町村クリック前のスクリーンショットを保存しました: {before_city_click}")
                                            
                                            try:
                                                # 通常のクリックを試す
                                                city_link.click()
                                                logger.info(f"市区町村名 '{link_text}' のリンクを通常クリック")
                                            except Exception as e:
                                                logger.warning(f"通常クリックに失敗: {str(e)}")
                                                try:
                                                    # JavaScriptを使用してクリック
                                                    driver.execute_script("arguments[0].click();", city_link)
                                                    logger.info(f"市区町村名 '{link_text}' のリンクをJavaScriptでクリック")
                                                except Exception as e2:
                                                    logger.error(f"JavaScriptクリックにも失敗: {str(e2)}")
                                                    continue
                                            
                                            city_link_found = True
                                            time.sleep(3)
                                            
                                            # クリック後のスクリーンショットを取得
                                            after_city_click = "after_city_click.png"
                                            driver.save_screenshot(after_city_click)
                                            logger.info(f"市区町村クリック後のスクリーンショットを保存しました: {after_city_click}")
                                            break
                                    
                                    if city_link_found:
                                        break
                                except (NoSuchElementException, ElementClickInterceptedException) as e:
                                    logger.warning(f"セレクタ {selector} で市区町村リンクが見つからないかクリックできません: {str(e)}")
                                    continue
                            
                            if not city_link_found:
                                # 市区町村リンクが見つからない場合、全てのリンクを取得して調査
                                logger.error(f"市区町村名 '{city}' のリンクが見つかりませんでした")
                                
                                # 全てのリンクを取得
                                all_links = driver.find_elements(By.TAG_NAME, "a")
                                link_texts = [link.text for link in all_links if link.text]
                                logger.info(f"ページ上の全てのリンクテキスト: {link_texts}")
                                
                                # 類似の市区町村名を探す
                                similar_cities = [link_text for link_text in link_texts if city in link_text]
                                if similar_cities:
                                    logger.info(f"類似の市区町村名が見つかりました: {similar_cities}")
                                    
                                    # 最も類似度の高い市区町村名をクリック
                                    for similar_city in similar_cities:
                                        try:
                                            similar_city_link = driver.find_element(By.XPATH, f"//a[text()='{similar_city}']")
                                            
                                            # 要素の位置までスクロール
                                            driver.execute_script("arguments[0].scrollIntoView(true);", similar_city_link)
                                            time.sleep(1)
                                            
                                            # JavaScriptを使用してクリック
                                            driver.execute_script("arguments[0].click();", similar_city_link)
                                            logger.info(f"類似の市区町村名 '{similar_city}' のリンクをクリック")
                                            city_link_found = True
                                            time.sleep(3)
                                            
                                            # クリック後のスクリーンショットを取得
                                            after_similar_city_click = "after_similar_city_click.png"
                                            driver.save_screenshot(after_similar_city_click)
                                            logger.info(f"類似の市区町村クリック後のスクリーンショットを保存しました: {after_similar_city_click}")
                                            break
                                        except (NoSuchElementException, ElementClickInterceptedException) as e:
                                            logger.warning(f"類似の市区町村名 '{similar_city}' のリンクがクリックできません: {str(e)}")
                                            continue
                                
                                if not city_link_found:
                                    # 市区町村入力フィールドを探す
                                    try:
                                        city_input_selectors = [
                                            "input[name='city']",
                                            "#id_tak_tx_skc_sk",
                                            "input.city-input",
                                            "//input[contains(@id, 'city')]",  # XPath
                                            "//input",  # すべての入力フィールド
                                        ]
                                        
                                        city_input = None
                                        for selector in city_input_selectors:
                                            try:
                                                if selector.startswith("//"):
                                                    # XPathの場合
                                                    city_input = WebDriverWait(driver, 5).until(
                                                        EC.presence_of_element_located((By.XPATH, selector))
                                                    )
                                                else:
                                                    # CSSセレクタの場合
                                                    city_input = WebDriverWait(driver, 5).until(
                                                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                                                    )
                                                logger.info(f"市区町村入力フィールドを発見 (セレクタ: {selector})")
                                                break
                                            except (TimeoutException, NoSuchElementException) as e:
                                                logger.warning(f"セレクタ {selector} で市区町村入力フィールドが見つかりません: {str(e)}")
                                                continue
                                        
                                        if city_input:
                                            # 市区町村名を入力
                                            city_input.clear()
                                            city_input.send_keys(city)
                                            logger.info(f"市区町村名を入力: {city}")
                                            time.sleep(1)
                                            
                                            # 検索ボタンを探して押す
                                            search_button_selectors = [
                                                "button.search-button",
                                                "//button[contains(text(), '検索')]",  # XPath
                                                "//button[contains(@class, 'search')]",  # XPath
                                                "//button",  # すべてのボタン
                                            ]
                                            
                                            for selector in search_button_selectors:
                                                try:
                                                    if selector.startswith("//"):
                                                        # XPathの場合
                                                        search_button = WebDriverWait(driver, 5).until(
                                                            EC.element_to_be_clickable((By.XPATH, selector))
                                                        )
                                                    else:
                                                        # CSSセレクタの場合
                                                        search_button = WebDriverWait(driver, 5).until(
                                                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                                                        )
                                                    
                                                    # JavaScriptを使用してクリック
                                                    driver.execute_script("arguments[0].click();", search_button)
                                                    logger.info(f"検索ボタンをクリック (セレクタ: {selector})")
                                                    city_link_found = True
                                                    time.sleep(3)
                                                    break
                                                except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                                                    logger.warning(f"セレクタ {selector} で検索ボタンが見つからないかクリックできません: {str(e)}")
                                                    continue
                                    except Exception as e:
                                        logger.error(f"市区町村入力処理中にエラーが発生: {str(e)}")
                                
                                if not city_link_found:
                                    # ページのHTMLを出力して調査
                                    html_content = driver.page_source
                                    with open("city_list_page.html", "w", encoding="utf-8") as f:
                                        f.write(html_content)
                                    logger.info("市区町村一覧ページのHTMLを city_list_page.html に保存しました")
                                    
                                    return {
                                        'status': 'ERROR',
                                        'message': f'市区町村名 {city} のリンクが見つかりませんでした'
                                    }
                        except Exception as e:
                            logger.error(f"都道府県選択処理中にエラーが発生: {str(e)}")
                            # スクリーンショットを取得
                            screenshot_path = "prefecture_error.png"
                            driver.save_screenshot(screenshot_path)
                            logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
                            return {
                                'status': 'ERROR',
                                'message': f'都道府県選択処理中にエラーが発生しました: {str(e)}'
                            }
                    else:
                        logger.error("都道府県選択フィールドが見つかりませんでした")
                        return {
                            'status': 'ERROR',
                            'message': '都道府県選択フィールドが見つかりませんでした'
                        }
                except Exception as e:
                    logger.error(f"都道府県選択処理中にエラーが発生: {str(e)}")
                    # スクリーンショットを取得
                    screenshot_path = "prefecture_error.png"
                    driver.save_screenshot(screenshot_path)
                    logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
                    return {
                        'status': 'ERROR',
                        'message': f'都道府県選択処理中にエラーが発生しました: {str(e)}'
                    }
            
            except Exception as e:
                logger.error(f"府県から検索するリンクの処理中にエラーが発生: {str(e)}")
                # スクリーンショットを取得
                screenshot_path = "tab_error.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
                
                # 直接郵便番号から検索する方法を試す
                return self.check_area_by_postal_code(prefecture, city, town, block)
        
            # デフォルトの結果（ここに到達することはほぼないはず）
            return {
                'status': 'ERROR',
                'message': f'{prefecture}{city}{town or ""}{block or ""} の提供可否を判定できませんでした'
            }
        except Exception as e:
            logger.error(f"府県から検索する方法でエラーが発生: {str(e)}")
            # スクリーンショットを取得
            if driver:
                screenshot_path = "general_error.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
            return {
                'status': 'ERROR',
                'message': f'エラーが発生しました：{str(e)}'
            }
        
        finally:
            if driver:
                driver.quit()
    
    def check_area_by_postal_code(self, prefecture, city, town=None, block=None):
        """
        郵便番号から検索する方法でエリア確認（フォールバック）
        
        Args:
            prefecture (str): 都道府県名
            city (str): 市区町村名
            town (str, optional): 町名
            block (str, optional): 番地
            
        Returns:
            dict: 提供可否の結果を含む辞書
        """
        driver = None
        try:
            # 住所から郵便番号を推定
            address = f"{prefecture}{city}{town or ''}{block or ''}"
            logger.info(f"郵便番号から検索する方法でエリア確認を開始: {address}")
            
            # 郵便番号を推定（実際のアプリケーションでは郵便番号DBを使用するか、APIを利用）
            # ここでは簡易的に「000-0000」を使用
            postal_code = "000-0000"
            logger.info(f"推定郵便番号: {postal_code}")
            
            driver = self.setup_driver()
            
            # サイトにアクセス
            url = "https://flets-w.com/cart/"
            logger.info(f"サイトにアクセス: {url}")
            driver.get(url)
            
            # ページの読み込みを待機
            time.sleep(5)
            
            # 郵便番号入力フィールドを探す
            postal_input_selectors = [
                "#id_tak_tx_ybk_yb",
                "input[name='zip_code']",
                "input.zip-code-input",
                "//input[contains(@placeholder, '郵便番号')]",  # XPath
            ]
            
            postal_input = None
            for selector in postal_input_selectors:
                try:
                    if selector.startswith("//"):
                        # XPathの場合
                        postal_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                    else:
                        # CSSセレクタの場合
                        postal_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                    logger.info(f"郵便番号入力フィールドを発見 (セレクタ: {selector})")
                    break
                except (TimeoutException, NoSuchElementException) as e:
                    logger.warning(f"セレクタ {selector} で郵便番号入力フィールドが見つかりません: {str(e)}")
                    continue
            
            if not postal_input:
                logger.error("郵便番号入力フィールドが見つかりません")
                # スクリーンショットを取得
                screenshot_path = "postal_input_error.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
                return {
                    'status': 'ERROR',
                    'message': '郵便番号入力フィールドが見つかりませんでした'
                }
            
            # 郵便番号を入力
            postal_input.clear()
            postal_input.send_keys(postal_code.replace('-', ''))
            logger.info(f"郵便番号を入力: {postal_code}")
            time.sleep(2)
            
            # 検索ボタンをクリック
            search_button_selectors = [
                "#id_tak_bt_ybk_jks",
                "button.search-button",
                "button[type='submit']",
                "//button[contains(text(), '検索')]",  # XPath
            ]
            
            search_button = None
            for selector in search_button_selectors:
                try:
                    if selector.startswith("//"):
                        # XPathの場合
                        search_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        # CSSセレクタの場合
                        search_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    logger.info(f"検索ボタンを発見 (セレクタ: {selector})")
                    break
                except (TimeoutException, NoSuchElementException) as e:
                    logger.warning(f"セレクタ {selector} で検索ボタンが見つかりません: {str(e)}")
                    continue
            
            if not search_button:
                logger.error("検索ボタンが見つかりません")
                # スクリーンショットを取得
                screenshot_path = "search_button_error.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
                return {
                    'status': 'ERROR',
                    'message': '検索ボタンが見つかりませんでした'
                }
            
            search_button.click()
            logger.info("郵便番号検索ボタンをクリック")
            time.sleep(3)
            
            # 町名・番地入力画面が表示されるまで待機
            time.sleep(3)
            
            # 町名・番地入力画面のスクリーンショットを取得
            town_screen_path = "town_screen.png"
            driver.save_screenshot(town_screen_path)
            logger.info(f"町名・番地入力画面のスクリーンショットを保存しました: {town_screen_path}")
            
            # 町名・番地入力画面のHTMLを保存
            html_content = driver.page_source
            with open("town_page.html", "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info("町名・番地入力画面のHTMLを town_page.html に保存しました")
            
            # 町名の頭文字に応じたカテゴリをクリック
            town_category_found = False
            if town:
                # 町名の頭文字を判定
                first_char = town[0]
                category_key = None
                
                # 頭文字からカテゴリキーを判定
                if 'あ' <= first_char <= 'お' or 'ア' <= first_char <= 'オ' or 'a' <= first_char.lower() <= 'o':
                    category_key = "a"  # あ行
                elif 'か' <= first_char <= 'ご' or 'カ' <= first_char <= 'ゴ' or 'k' == first_char.lower() or 'g' == first_char.lower():
                    category_key = "ka"  # か行
                elif 'さ' <= first_char <= 'そ' or 'サ' <= first_char <= 'ソ' or 's' == first_char.lower() or 'z' == first_char.lower():
                    category_key = "sa"  # さ行
                elif 'た' <= first_char <= 'ど' or 'タ' <= first_char <= 'ド' or 't' == first_char.lower() or 'd' == first_char.lower():
                    category_key = "ta"  # た行
                elif 'な' <= first_char <= 'の' or 'ナ' <= first_char <= 'ノ' or 'n' == first_char.lower():
                    category_key = "na"  # な行
                elif 'は' <= first_char <= 'ぽ' or 'ハ' <= first_char <= 'ポ' or 'h' == first_char.lower() or 'b' == first_char.lower() or 'p' == first_char.lower():
                    category_key = "ha"  # は行
                elif 'ま' <= first_char <= 'も' or 'マ' <= first_char <= 'モ' or 'm' == first_char.lower():
                    category_key = "ma"  # ま行
                elif 'や' <= first_char <= 'よ' or 'ヤ' <= first_char <= 'ヨ' or 'y' == first_char.lower():
                    category_key = "ya"  # や行
                elif 'ら' <= first_char <= 'ろ' or 'ラ' <= first_char <= 'ロ' or 'r' == first_char.lower():
                    category_key = "ra"  # ら行
                elif 'わ' <= first_char <= 'ん' or 'ワ' <= first_char <= 'ン' or 'w' == first_char.lower():
                    category_key = "wa"  # わ行
                else:
                    # デフォルトは「あ行」
                    category_key = "a"
                
                logger.info(f"町名 '{town}' の頭文字 '{first_char}' からカテゴリキー '{category_key}' を判定")
                
                # カテゴリをクリック
                category_selectors = [
                    f"//li[@data-kana-order-key='{category_key}']/a",
                    f"//li[@data-kana-order-key='{category_key}']",
                    f"//li[contains(@class, '{category_key}')]/a",
                    f"//a[contains(@href, '{category_key}')]",
                ]
                
                for selector in category_selectors:
                    try:
                        category_elements = driver.find_elements(By.XPATH, selector)
                        
                        if not category_elements:
                            logger.warning(f"セレクタ {selector} でカテゴリ要素が見つかりません")
                            continue
                        
                        for category_element in category_elements:
                            try:
                                # 要素の位置までスクロール
                                driver.execute_script("arguments[0].scrollIntoView(true);", category_element)
                                time.sleep(1)
                                
                                # JavaScriptを使用してクリック
                                driver.execute_script("arguments[0].click();", category_element)
                                logger.info(f"カテゴリ '{category_key}' をクリック (テキスト: {category_element.text})")
                                town_category_found = True
                                time.sleep(2)
                                
                                # クリック後のスクリーンショットを取得
                                after_category_click = "after_category_click.png"
                                driver.save_screenshot(after_category_click)
                                logger.info(f"カテゴリクリック後のスクリーンショットを保存しました: {after_category_click}")
                                break
                            except Exception as e:
                                logger.warning(f"カテゴリ要素のクリックに失敗: {str(e)}")
                                continue
                        
                        if town_category_found:
                            break
                    except Exception as e:
                        logger.warning(f"カテゴリ要素の検索に失敗: {str(e)}")
                        continue
                
                # カテゴリが見つからない場合は全てのリンクを取得して調査
                if not town_category_found:
                    logger.warning(f"カテゴリ '{category_key}' が見つかりませんでした")
                    
                    # 全てのリンクを取得
                    all_links = driver.find_elements(By.TAG_NAME, "a")
                    link_texts = [link.text for link in all_links if link.text]
                    logger.info(f"ページ上の全てのリンクテキスト: {link_texts}")
                
                # 町名リンクをクリック
                town_link_found = False
                if town:
                    town_link_selectors = [
                        f"//a[contains(text(), '{town}')]",  # テキストで検索
                        f"//a[text()='{town}']",  # 完全一致テキスト
                        f"//li//a[contains(text(), '{town}')]",
                        f"//li//a[text()='{town}']",
                        f"//div[@class='candidate_list_wrap']//ul/li//a[contains(text(), '{town}')]",
                        f"//div[@class='candidate_list_wrap']//ul/li//a[text()='{town}']",
                        f"//div[@id='scrollBoxprefectureAddressSearchModalId']//ul/li//a[contains(text(), '{town}')]",
                        f"//div[@id='scrollBoxprefectureAddressSearchModalId']//ul/li//a[text()='{town}']",
                    ]
                    
                    for selector in town_link_selectors:
                        try:
                            logger.info(f"町名リンクのセレクタを試行: {selector}")
                            town_links = driver.find_elements(By.XPATH, selector)
                            
                            if not town_links:
                                logger.warning(f"セレクタ {selector} で町名リンクが見つかりません")
                                continue
                            
                            logger.info(f"セレクタ {selector} で {len(town_links)} 個のリンクが見つかりました")
                            
                            # 町名が完全に一致するリンクを探す
                            for town_link in town_links:
                                link_text = town_link.text
                                logger.info(f"リンクテキスト: '{link_text}'")
                                
                                if town in link_text:
                                    # 要素の位置までスクロール
                                    driver.execute_script("arguments[0].scrollIntoView(true);", town_link)
                                    time.sleep(1)
                                    
                                    # クリック前のスクリーンショットを取得
                                    before_town_click = "before_town_click.png"
                                    driver.save_screenshot(before_town_click)
                                    logger.info(f"町名クリック前のスクリーンショットを保存しました: {before_town_click}")
                                    
                                    try:
                                        # 通常のクリックを試す
                                        town_link.click()
                                        logger.info(f"町名 '{link_text}' のリンクを通常クリック")
                                    except Exception as e:
                                        logger.warning(f"通常クリックに失敗: {str(e)}")
                                        try:
                                            # JavaScriptを使用してクリック
                                            driver.execute_script("arguments[0].click();", town_link)
                                            logger.info(f"町名 '{link_text}' のリンクをJavaScriptでクリック")
                                        except Exception as e2:
                                            logger.error(f"JavaScriptクリックにも失敗: {str(e2)}")
                                            continue
                                    
                                    town_link_found = True
                                    time.sleep(3)
                                    
                                    # クリック後のスクリーンショットを取得
                                    after_town_click = "after_town_click.png"
                                    driver.save_screenshot(after_town_click)
                                    logger.info(f"町名クリック後のスクリーンショットを保存しました: {after_town_click}")
                                    break
                            
                            if town_link_found:
                                break
                        except (NoSuchElementException, ElementClickInterceptedException) as e:
                            logger.warning(f"セレクタ {selector} で町名リンクが見つからないかクリックできません: {str(e)}")
                            continue
            
            # 番地入力フィールドを探す
            block_input_found = False
            block_input_selectors = [
                "input[name='block']",
                "#id_tak_tx_skc_bn",
                "input.block-input",
                "//input[contains(@id, 'block')]",  # XPath
            ]
            
            for selector in block_input_selectors:
                try:
                    if selector.startswith("//"):
                        # XPathの場合
                        block_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                    else:
                        # CSSセレクタの場合
                        block_input = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                    
                    if block and block_input:
                        # 番地を入力
                        block_input.clear()
                        block_input.send_keys(block)
                        logger.info(f"番地を入力: {block}")
                        block_input_found = True
                        break
                except (TimeoutException, NoSuchElementException) as e:
                    logger.warning(f"セレクタ {selector} で番地入力フィールドが見つかりません: {str(e)}")
                    continue
            
            # 「検索結果を確認」ボタンを探して押す
            confirm_button_selectors = [
                "button.btn_search_result",
                "//button[contains(text(), '検索結果を確認')]",  # XPath
                "//button[contains(@class, 'btn_search_result')]",  # XPath
                "//button[contains(@class, 'search-result')]",  # XPath
                "//button[contains(@class, 'btn-primary')]",  # XPath
                "//button",  # すべてのボタン
            ]
            
            confirm_button_found = False
            for selector in confirm_button_selectors:
                try:
                    if selector.startswith("//"):
                        # XPathの場合
                        confirm_buttons = driver.find_elements(By.XPATH, selector)
                    else:
                        # CSSセレクタの場合
                        confirm_buttons = driver.find_elements(By.CSS_SELECTOR, selector)
                    
                    if not confirm_buttons:
                        logger.warning(f"セレクタ {selector} で検索結果確認ボタンが見つかりません")
                        continue
                    
                    # ボタンのテキストを確認
                    for confirm_button in confirm_buttons:
                        button_text = confirm_button.text
                        logger.info(f"ボタンのテキスト: {button_text}")
                        
                        if "検索結果" in button_text or "確認" in button_text or button_text == "":
                            # 要素の位置までスクロール
                            driver.execute_script("arguments[0].scrollIntoView(true);", confirm_button)
                            time.sleep(1)
                            
                            # JavaScriptを使用してクリック
                            driver.execute_script("arguments[0].click();", confirm_button)
                            logger.info(f"検索結果確認ボタンをクリック (テキスト: {button_text})")
                            confirm_button_found = True
                            time.sleep(5)
                            
                            # クリック後のスクリーンショットを取得
                            after_confirm_click = "after_confirm_click.png"
                            driver.save_screenshot(after_confirm_click)
                            logger.info(f"検索結果確認ボタンクリック後のスクリーンショットを保存しました: {after_confirm_click}")
                            break
                    
                    if confirm_button_found:
                        break
                except (NoSuchElementException, ElementClickInterceptedException) as e:
                    logger.warning(f"セレクタ {selector} で検索結果確認ボタンがクリックできません: {str(e)}")
                    continue
            
            # 結果ページのHTMLを保存
            html_content = driver.page_source
            with open("result_page.html", "w", encoding="utf-8") as f:
                f.write(html_content)
            logger.info("結果ページのHTMLを result_page.html に保存しました")
            
            # 結果を判定
            result_selectors = [
                "//div[contains(@class, 'result-ok')]",  # 提供可能
                "//div[contains(@class, 'result-ng')]",  # 提供不可
                "//div[contains(text(), '提供可能')]",
                "//div[contains(text(), '提供不可')]",
                "//p[contains(text(), '提供可能')]",
                "//p[contains(text(), '提供不可')]",
                "//span[contains(text(), '提供可能')]",
                "//span[contains(text(), '提供不可')]",
            ]
            
            result_found = False
            for selector in result_selectors:
                try:
                    result_elements = driver.find_elements(By.XPATH, selector)
                    
                    if result_elements:
                        result_text = result_elements[0].text
                        logger.info(f"結果テキスト: {result_text}")
                        
                        if "提供可能" in result_text or "ご利用いただけます" in result_text:
                            logger.info("提供可能と判定されました")
                            return {
                                'status': 'OK',
                                'message': f'{prefecture}{city}{town or ""}{block or ""} は提供可能エリアです'
                            }
                        elif "提供不可" in result_text or "ご利用いただけません" in result_text:
                            logger.info("提供不可と判定されました")
                            return {
                                'status': 'NG',
                                'message': f'{prefecture}{city}{town or ""}{block or ""} は提供不可エリアです'
                            }
                        
                        result_found = True
                        break
                except NoSuchElementException as e:
                    logger.warning(f"セレクタ {selector} で結果要素が見つかりません: {str(e)}")
                    continue
            
            if not result_found:
                # 結果が見つからない場合は、ページ全体のテキストから判定を試みる
                page_text = driver.find_element(By.TAG_NAME, "body").text
                
                if "提供可能" in page_text or "ご利用いただけます" in page_text:
                    logger.info("ページテキストから提供可能と判定されました")
                    return {
                        'status': 'OK',
                        'message': f'{prefecture}{city}{town or ""}{block or ""} は提供可能エリアです'
                    }
                elif "提供不可" in page_text or "ご利用いただけません" in page_text:
                    logger.info("ページテキストから提供不可と判定されました")
                    return {
                        'status': 'NG',
                        'message': f'{prefecture}{city}{town or ""}{block or ""} は提供不可エリアです'
                    }
                else:
                    logger.warning("結果を判定できませんでした")
                    return {
                        'status': 'ERROR',
                        'message': f'{prefecture}{city}{town or ""}{block or ""} の提供可否を判定できませんでした'
                    }
            
            # デフォルトの結果（ここに到達することはほぼないはず）
            return {
                'status': 'ERROR',
                'message': f'{prefecture}{city}{town or ""}{block or ""} の提供可否を判定できませんでした'
            }
        
        except Exception as e:
            logger.error(f"郵便番号から検索する方法でエラーが発生: {str(e)}")
            # スクリーンショットを取得
            if driver:
                screenshot_path = "postal_error.png"
                driver.save_screenshot(screenshot_path)
                logger.info(f"エラー時のスクリーンショットを保存しました: {screenshot_path}")
            return {
                'status': 'ERROR',
                'message': f'エラーが発生しました：{str(e)}'
            }
        
        finally:
            if driver:
                driver.quit()

# 単体テスト用
if __name__ == "__main__":
    checker = AreaChecker()
    
    # 入力を受け付ける
    prefecture = input("都道府県名を入力してください（例: 大阪府）: ")
    city = input("市区町村名を入力してください（例: 大阪市中央区）: ")
    town = input("町名を入力してください（省略可）: ") or None
    block = input("番地を入力してください（省略可）: ") or None
    
    result = checker.check_area(prefecture, city, town, block)
    print(json.dumps(result, ensure_ascii=False, indent=2)) 