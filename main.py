"""
コールセンター業務効率化ツール

このスクリプトは、コールセンター業務の効率化を目的としたGUIアプリケーションです。
PySide6を使用してUIを構築し、Google Spreadsheetsとの連携機能を提供します。

主な機能：
- 顧客情報の入力
- CTIフォーマットの生成
- スプレッドシートへのデータ転記
- クリップボード監視機能
- 提供エリア検索機能
"""

import sys
import logging
from PySide6.QtWidgets import QApplication
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from ui.main_window import MainWindow


# ログの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('app.log', encoding='utf-8')
    ]
)


def debug_address_selection(driver):
    """
    番地選択画面の要素を確認するデバッグ関数
    
    Args:
        driver: WebDriverインスタンス
    """
    try:
        # 番地選択画面の表示を待機
        logging.info("番地選択画面の要素を確認中...")
        time.sleep(2)  # 画面の表示を待機
        
        # 現在のフレーム情報を出力
        try:
            current_frame = driver.execute_script("return self.name")
            logging.info(f"現在のフレーム名: {current_frame}")
        except:
            logging.info("現在のフレーム名を取得できません")
        
        # 画面上の全要素を取得（より広範な属性を確認）
        all_elements = driver.find_elements(By.XPATH, "//*[not(self::script) and not(self::style)]")
        visible_elements = [elem for elem in all_elements if elem.is_displayed()]
        
        logging.info(f"画面上の表示要素数: {len(visible_elements)}")
        
        # 番地に関連する要素を特に詳しく確認
        address_related_elements = driver.find_elements(
            By.XPATH,
            "//*[contains(text(), '番地') or contains(text(), '住所') or "
            "contains(@placeholder, '番地') or contains(@aria-label, '番地') or "
            "contains(@name, 'address') or contains(@class, 'address')]"
        )
        
        logging.info(f"番地関連の要素数: {len(address_related_elements)}")
        for elem in address_related_elements:
            if elem.is_displayed():
                logging.info("=== 番地関連要素の詳細 ===")
                logging.info(f"タグ名: {elem.tag_name}")
                logging.info(f"テキスト: {elem.text.strip() if elem.text else 'なし'}")
                logging.info(f"type属性: {elem.get_attribute('type')}")
                logging.info(f"class属性: {elem.get_attribute('class')}")
                logging.info(f"name属性: {elem.get_attribute('name')}")
                logging.info(f"id属性: {elem.get_attribute('id')}")
                logging.info(f"placeholder属性: {elem.get_attribute('placeholder')}")
                logging.info(f"aria-label属性: {elem.get_attribute('aria-label')}")
                logging.info(f"表示状態: {elem.is_displayed()}")
                logging.info(f"クリック可能: {elem.is_enabled()}")
                try:
                    rect = elem.rect
                    logging.info(f"位置情報: x={rect['x']}, y={rect['y']}, "
                               f"width={rect['width']}, height={rect['height']}")
                except:
                    logging.info("位置情報: 取得できません")
                logging.info("========================")
        
        # リスト要素（番地選択肢）の確認
        list_elements = driver.find_elements(
            By.XPATH,
            "//ul/li | //select/option | //div[@role='listbox']//div[@role='option']"
        )
        
        logging.info(f"リスト要素数: {len(list_elements)}")
        for elem in list_elements:
            if elem.is_displayed():
                logging.info(f"リスト要素: テキスト='{elem.text.strip()}', "
                           f"タグ名={elem.tag_name}, "
                           f"クラス={elem.get_attribute('class')}")
        
        # 入力フィールドの確認（番地入力用）
        input_elements = driver.find_elements(By.TAG_NAME, "input")
        logging.info(f"入力フィールド数: {len(input_elements)}")
        for input_elem in input_elements:
            if input_elem.is_displayed():
                logging.info("=== 入力フィールドの詳細 ===")
                logging.info(f"type属性: {input_elem.get_attribute('type')}")
                logging.info(f"id属性: {input_elem.get_attribute('id')}")
                logging.info(f"name属性: {input_elem.get_attribute('name')}")
                logging.info(f"placeholder: {input_elem.get_attribute('placeholder')}")
                logging.info(f"value: {input_elem.get_attribute('value')}")
                logging.info("========================")
        
        # デバッグ用に画面のスクリーンショットを保存
        driver.save_screenshot("address_selection_debug.png")
        logging.info("スクリーンショットを保存しました: address_selection_debug.png")
        
        # ページのHTMLソースを保存（デバッグ用）
        with open("page_source_debug.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        logging.info("ページソースを保存しました: page_source_debug.html")
        
        # デバッグ確認用の待機時間
        logging.info("デバッグ確認のため10秒間待機します...")
        time.sleep(10)
        
    except Exception as e:
        logging.error(f"デバッグ中にエラーが発生: {str(e)}")
        import traceback
        logging.error(traceback.format_exc())


def check_west_japan(postal_code):
    """
    西日本エリアの郵便番号チェックを行う関数
    
    Args:
        postal_code (str): チェックする郵便番号
    
    Returns:
        bool: 提供エリアの場合はTrue、それ以外はFalse
    """
    try:
        # Chromeドライバーの設定
        chrome_options = Options()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.implicitly_wait(10)
        
        # URLにアクセス
        url = "https://flets-w.com/cart/"
        driver.get(url)
        
        # 郵便番号入力欄を探す
        postal_input = None
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        logging.info(f"フレーム数: {len(frames)}")
        
        target_frame = None
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                postal_input = driver.find_element(By.CSS_SELECTOR, "input[type='text']")
                if postal_input:
                    target_frame = frame
                    break
            except:
                driver.switch_to.default_content()
                continue
        
        if not postal_input:
            raise Exception("郵便番号入力欄が見つかりませんでした")
        
        # 郵便番号を入力
        postal_code = postal_code.replace("-", "")
        postal_input.clear()
        postal_input.send_keys(postal_code)
        
        # 検索ボタンをクリック
        search_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        search_button.click()
        
        # メインフレームに戻る
        driver.switch_to.default_content()
        
        # 番地選択画面の処理
        logging.info("番地選択画面の処理を開始します")
        time.sleep(3)  # 画面遷移を待機
        
        # 全てのフレームをチェック
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        logging.info(f"番地選択画面のフレーム数: {len(frames)}")
        
        address_frame_found = False
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                logging.info(f"フレームの切り替え - src: {frame.get_attribute('src')}")
                
                # フレーム内の要素を確認
                elements = driver.find_elements(By.XPATH, "//*[contains(text(), '番地') or contains(text(), '住所')]")
                if elements:
                    logging.info("番地関連の要素が見つかりました")
                    address_frame_found = True
                    break
                
                driver.switch_to.default_content()
            except Exception as e:
                logging.error(f"フレーム切り替え中のエラー: {str(e)}")
                driver.switch_to.default_content()
                continue
        
        if not address_frame_found:
            logging.error("番地選択用のフレームが見つかりませんでした")
            return False
        
        # デバッグ関数を呼び出し
        debug_address_selection(driver)
        
        # 番地選択画面の処理を続行
        try:
            # 番地入力に関連する要素を探す（複数のセレクタを試行）
            selectors = [
                "input[type='text']",
                "input[placeholder*='番地']",
                "input[aria-label*='番地']",
                "input[name*='address']",
                "input.address-input"
            ]
            
            address_input = None
            for selector in selectors:
                try:
                    address_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    if address_input:
                        logging.info(f"番地入力フィールドが見つかりました（セレクタ: {selector}）")
                        break
                except TimeoutException:
                    continue
            
            if not address_input:
                raise TimeoutException("番地入力フィールドが見つかりませんでした")
            
            # 番地入力フィールドの属性を確認
            logging.info(f"番地入力フィールド - ID: {address_input.get_attribute('id')}, "
                        f"Name: {address_input.get_attribute('name')}, "
                        f"Type: {address_input.get_attribute('type')}")
            
        except TimeoutException as e:
            logging.error(f"番地入力フィールドの検出に失敗: {str(e)}")
            return False
        
        # 処理を続行...
        
    except Exception as e:
        logging.error(f"エラーが発生しました: {str(e)}")
        return False
    finally:
        # ブラウザを閉じる前に10秒待機（デバッグ用）
        time.sleep(10)
        driver.quit()
    
    return True


def main():
    """アプリケーションのメイン関数"""
    # アプリケーションの作成
    app = QApplication(sys.argv)
    
    # メインウィンドウの作成と表示
    window = MainWindow()
    window.show()
    
    # アプリケーションの実行
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 