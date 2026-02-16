"""
MapFan検索サービス

このモジュールは、MapFanで住所検索を実行し、
検索結果の詳細画面へ遷移したURLを取得する機能を提供します。
"""

import logging
import time
import threading
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from services.web_driver import create_driver, load_browser_settings


class MapfanService:
    """MapFan検索サービス"""

    LEFT_SEARCH_BUTTON_LOCATORS: List[Tuple[str, str]] = [
        (By.XPATH, "//button[normalize-space()='検索' and not(@disabled)]"),
        (By.XPATH, "//a[normalize-space()='検索']"),
        (By.XPATH, "//*[self::button or self::a][contains(normalize-space(.), '検索')]")
    ]

    SEARCH_INPUT_LOCATORS: List[Tuple[str, str]] = [
        (By.XPATH, "//input[@type='text' and @maxlength='150']"),
        (By.XPATH, "//input[@type='text' and @role='combobox']"),
        (By.XPATH, "//input[@type='search' and contains(@placeholder, 'スポット/住所/駅/TEL')]") ,
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL')]") ,
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL/郵便番号')]") ,
        (By.CSS_SELECTOR, "input[placeholder='スポット/住所/駅/TEL/郵便番号']"),
        (By.CSS_SELECTOR, "input[name*='word']"),
        (By.CSS_SELECTOR, "input[id*='word']"),
        (By.CSS_SELECTOR, "input[placeholder*='検索']"),
        (By.CSS_SELECTOR, "input[type='search']"),
        (By.CSS_SELECTOR, "input[placeholder*='住所']"),
        (By.CSS_SELECTOR, "input[aria-label*='住所']"),
        (By.XPATH, "//input[contains(@placeholder, '住所') or contains(@aria-label, '住所')]"),
        (By.XPATH, "//header//input[@type='text' or @type='search']"),
    ]

    SEARCH_BUTTON_LOCATORS: List[Tuple[str, str]] = [
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL')]/following-sibling::*[self::button or self::a][1]"),
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL')]/ancestor::*[self::form or self::div][1]//*[self::button or self::a][1]"),
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL/郵便番号')]/following-sibling::*[self::button or self::a][1]"),
        (By.XPATH, "//input[contains(@placeholder, 'スポット/住所/駅/TEL/郵便番号')]/ancestor::*[self::form or self::div][1]//*[self::button or self::a][1]"),
        (By.CSS_SELECTOR, "button[type='submit']"),
        (By.CSS_SELECTOR, "button[aria-label*='検索']"),
        (By.CSS_SELECTOR, "button[title*='検索']"),
        (By.XPATH, "//button[contains(normalize-space(.), '検索')]") ,
        (By.XPATH, "//header//*[self::button or self::a][contains(@class, 'search')]") ,
    ]

    SEARCH_RESULT_READY_LOCATORS: List[Tuple[str, str]] = [
        (By.XPATH, "//button[contains(@aria-label, '詳細') or contains(@title, '詳細')]") ,
        (By.XPATH, "//button[normalize-space()='i']"),
        (By.XPATH, "//div[contains(@class, 'spot')]//*[self::button or self::a][contains(@class, 'info')]") ,
        (By.XPATH, "//*[self::button or self::a][contains(@class, 'info') or contains(@aria-label, '情報')]") ,
        (By.XPATH, "//li[contains(@class, 'result')] | //div[contains(@class, 'result')]"),
    ]

    DETAIL_BUTTON_LOCATORS: List[Tuple[str, str]] = [
        (By.XPATH, "//button[contains(@aria-label, '詳細') or contains(@title, '詳細')]") ,
        (By.XPATH, "//button[contains(@aria-label, '情報') or contains(@title, '情報')]") ,
        (By.XPATH, "//button[normalize-space()='i']"),
        (By.XPATH, "//*[self::button or self::a][.//*[normalize-space()='i']]") ,
        (By.XPATH, "//*[self::button or self::a][contains(@class, 'info')]") ,
        (By.CSS_SELECTOR, "button[class*='info']"),
        (By.CSS_SELECTOR, "a[class*='info']"),
    ]

    def __init__(
        self,
        base_url: str = "https://mapfan.com/map",
        timeout: int = 30,
        debug: bool = True,
        detailed_logging: bool = False
    ):
        self.base_url = base_url
        self.timeout = timeout
        self.debug = debug
        self.detailed_logging = detailed_logging
        self._cancel_event: Optional[threading.Event] = None
        self._active_driver = None

    def request_cancel(self) -> None:
        try:
            if self._cancel_event is not None:
                self._cancel_event.set()
            if self._active_driver is not None:
                try:
                    self._active_driver.quit()
                except Exception:
                    pass
        except Exception:
            pass

    def _is_cancelled(self) -> bool:
        return bool(self._cancel_event is not None and self._cancel_event.is_set())

    def _check_cancel(self) -> None:
        if self._is_cancelled():
            raise TimeoutException("MapFan処理がキャンセルされました")

    def get_detail_url_from_address(
        self,
        address: str,
        auto_close: Optional[bool] = None,
        force_headless: Optional[bool] = None,
        cancel_event: Optional[threading.Event] = None
    ) -> Optional[str]:
        """
        住所検索から詳細画面へ遷移し、遷移先URLを取得します。

        Args:
            address (str): 検索に使う住所文字列
            auto_close (Optional[bool]): 処理後にブラウザを閉じるかどうか

        Returns:
            Optional[str]: 遷移先URL。取得できない場合はNone
        """
        normalized_address = (address or "").strip()
        if not normalized_address:
            logging.info("MapFan検索をスキップ: 住所情報が未入力です")
            return None

        settings = load_browser_settings()
        headless = settings.get("headless", True)
        if force_headless is not None:
            headless = bool(force_headless)
        elif self.debug:
            headless = False

        if auto_close is None:
            auto_close = settings.get("auto_close", False)

        self._cancel_event = cancel_event
        driver = None
        try:
            self._check_cancel()
            driver = create_driver(headless=headless, page_load_strategy="eager")
            self._active_driver = driver
            driver.implicitly_wait(0)
            wait = WebDriverWait(driver, self.timeout)

            navigation_urls = [self.base_url]
            if headless and self.base_url != "https://mapfan.com/":
                navigation_urls.append("https://mapfan.com/")

            page_title = ""
            for index, target_url in enumerate(navigation_urls, start=1):
                self._check_cancel()
                driver.get(target_url)
                WebDriverWait(driver, 12).until(
                    lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
                )

                page_title = (driver.title or "").strip()
                logging.info(f"MapFanを表示しました: title={page_title}, url={target_url}")

                if not self._is_mapfan_block_page(driver, page_title):
                    break

                self._log_block_page_diagnostics(driver, phase=f"initial-load-{index}")
                if index < len(navigation_urls):
                    logging.warning("MapFanブロックページを検出。ヘッドレスのまま別URLで再試行します")
                    continue

                logging.error("MapFanのブロック/エラーページを検出しました（ヘッドレス実行のまま終了）")
                return None

            self._check_cancel()
            self._input_address_and_search(driver, wait, normalized_address)
            previous_url = driver.current_url

            detail_clicked = self._click_left_panel_info_button(driver, normalized_address)
            if not detail_clicked:
                detail_button = self._find_first_clickable(driver, self.DETAIL_BUTTON_LOCATORS, wait, timeout_per_locator=2)
                if detail_button is None:
                    logging.error("MapFanの詳細ボタン（iマーク）を見つけられませんでした")
                    return None
                self._safe_click(driver, detail_button)
                logging.info("汎用ロケータで詳細ボタンをクリックしました")
            else:
                logging.info("左パネルのiボタンをクリックしました")

            wait.until(lambda d: d.current_url != previous_url)
            detail_url = driver.current_url
            logging.info(f"MapFan詳細URL取得成功: {detail_url}")
            return detail_url

        except TimeoutException:
            if self._is_cancelled():
                logging.info("MapFan処理をキャンセルしました")
            else:
                logging.error("MapFan操作中にタイムアウトが発生しました")
            if driver is not None:
                self._log_block_page_diagnostics(driver, phase="timeout")
            return None
        except Exception as e:
            logging.error(f"MapFan詳細URL取得中にエラーが発生しました: {str(e)}")
            if driver is not None:
                self._log_block_page_diagnostics(driver, phase="exception")
            return None
        finally:
            self._active_driver = None
            if driver is not None and auto_close:
                try:
                    driver.quit()
                except Exception as e:
                    logging.warning(f"MapFan用WebDriver終了時にエラーが発生しました: {str(e)}")

    def _input_address_and_search(self, driver: WebDriver, wait: WebDriverWait, address: str) -> None:
        self._check_cancel()
        self._log_search_inputs(driver, "左検索ボタン押下前")
        self._dismiss_blocking_overlay(driver)

        if not self._click_left_search_button(driver):
            raise TimeoutException("左側の検索ボタンをクリックできませんでした")

        time.sleep(0.2)
        self._dismiss_blocking_overlay(driver)
        self._log_search_inputs(driver, "左検索ボタン押下後")

        search_input = self._find_visible_search_input_after_left_click(driver)
        if search_input is None:
            raise TimeoutException("左検索UIの入力欄を見つけられませんでした")

        logging.info("MapFan検索欄を検出しました")
        self._log_element_summary(search_input, "選択した検索欄")

        value_after_input = ""
        for index in range(5):
            self._check_cancel()
            try:
                self._dismiss_blocking_overlay(driver)
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", search_input)
                try:
                    search_input.click()
                except Exception:
                    driver.execute_script("arguments[0].focus();", search_input)

                search_input.send_keys(Keys.CONTROL, "a")
                search_input.send_keys(Keys.BACKSPACE)
                search_input.send_keys(address)

                # Angular/Material向けにイベントを明示発火
                driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                    search_input
                )

                value_after_input = (search_input.get_attribute("value") or "").strip()
                if value_after_input:
                    logging.info(f"検索欄への入力に成功しました (試行 {index + 1}/5): {value_after_input}")
                    break
                logging.warning(f"検索欄入力後の値が空でした (試行 {index + 1}/5)")
            except Exception as input_error:
                logging.warning(f"検索欄入力時に例外が発生しました (試行 {index + 1}/5): {input_error}")
                time.sleep(0.1)

        if not value_after_input:
            raise TimeoutException("MapFan検索欄への住所入力後、値が反映されませんでした")

        logging.info(f"MapFan検索欄に住所を入力しました: {value_after_input}")

        # 検索ボタン押下直前に値を再確認（消えるケースの対策）
        current_value = (search_input.get_attribute("value") or "").strip()
        if not current_value:
            try:
                search_input.send_keys(address)
                current_value = (search_input.get_attribute("value") or "").strip()
                logging.info(f"検索押下前に住所を再入力しました: {current_value}")
            except Exception:
                pass

        search_button = self._find_search_submit_button(driver, wait, search_input)
        if search_button is not None:
            try:
                button_text = (search_button.text or "").strip().lower()
            except Exception:
                button_text = ""

            if button_text in ("clear", "×", "x"):
                search_input.send_keys(Keys.ENTER)
                logging.info("clearボタンを回避し、Enterで検索しました")
            else:
                self._safe_click(driver, search_button)
                logging.info("検索ボタンをクリックしました")
        else:
            search_input.send_keys(Keys.ENTER)
            logging.info("検索ボタンが見つからないためEnterで検索しました")

        self._wait_for_search_result(driver, wait)
        logging.info("MapFan検索結果の表示を確認しました")

    def _click_left_search_button(self, driver: WebDriver) -> bool:
        try:
            js_button = driver.execute_script(
                "const nodes = Array.from(document.querySelectorAll('button, a'));"
                "for (const el of nodes) {"
                "  const r = el.getBoundingClientRect();"
                "  const st = window.getComputedStyle(el);"
                "  if (st.display === 'none' || st.visibility === 'hidden') continue;"
                "  if (r.width < 20 || r.height < 20) continue;"
                "  if (r.left > 320) continue;"
                "  const label = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).replace(/\\s+/g, '');"
                "  if (label.includes('検索')) return el;"
                "}"
                "return null;"
            )
            if js_button is not None:
                self._safe_click(driver, js_button)
                logging.info("左側の検索ボタンをクリックしました")
                return True
        except Exception:
            pass

        for by, selector in self.LEFT_SEARCH_BUTTON_LOCATORS:
            try:
                elements = driver.find_elements(by, selector)
                for element in elements:
                    try:
                        if not element.is_displayed() or not element.is_enabled():
                            continue
                        if int(element.location.get("x", 9999)) > 300:
                            continue
                        self._safe_click(driver, element)
                        logging.info("左側の検索ボタンをクリックしました")
                        return True
                    except Exception:
                        continue
            except Exception:
                continue
        return False

    def _find_visible_search_input_after_left_click(self, driver: WebDriver):
        end_time = time.time() + 10
        while time.time() < end_time:
            self._check_cancel()
            try:
                candidates = driver.find_elements(By.XPATH, "//input[@type='search' or @type='text']")
                best = None
                best_score = -1
                for element in candidates:
                    try:
                        if not element.is_displayed() or not element.is_enabled():
                            continue
                        location = element.location
                        size = element.size
                        x_pos = int(location.get('x', 9999))
                        y_pos = int(location.get('y', 9999))
                        width = int(size.get('width', 0))
                        height = int(size.get('height', 0))
                        if x_pos > 420 or y_pos > 220:
                            continue
                        if width < 160 or height < 18:
                            continue

                        placeholder = (element.get_attribute('placeholder') or '').strip()
                        input_type = (element.get_attribute('type') or '').strip()
                        score = width
                        if input_type == 'search':
                            score += 300
                        if 'スポット' in placeholder or '住所' in placeholder:
                            score += 1000

                        if score > best_score:
                            best = element
                            best_score = score
                    except Exception:
                        continue

                if best is not None:
                    logging.info(f"左検索UI入力欄を採用: score={best_score}")
                    return best
            except Exception:
                pass
            time.sleep(0.2)
        return None

    def _submit_search_via_js(self, driver: WebDriver, address: str) -> bool:
        """初期画面の上部入力欄へJSで直接入力し、検索ボタンを押下する"""
        try:
            result = driver.execute_script(
                "const keyword = arguments[0];"
                "const inputs = Array.from(document.querySelectorAll('input[type=\"text\"], input[type=\"search\"]'));"
                "let target = null;"
                "let bestScore = -1;"
                "for (const el of inputs) {"
                "  const rect = el.getBoundingClientRect();"
                "  const style = window.getComputedStyle(el);"
                "  const visible = rect.width > 220 && rect.height > 18 && rect.top >= 0 && rect.top < 220 && style.display !== 'none' && style.visibility !== 'hidden';"
                "  if (!visible || el.disabled) continue;"
                "  const ph = (el.getAttribute('placeholder') || '');"
                "  let score = rect.width;"
                "  if ((el.getAttribute('type') || '') === 'search') score += 200;"
                "  if (ph.includes('スポット') || ph.includes('住所')) score += 800;"
                "  if (score > bestScore) { bestScore = score; target = el; }"
                "}"
                "if (!target) return {ok:false, reason:'input_not_found'};"
                "const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;"
                "setter.call(target, keyword);"
                "target.dispatchEvent(new Event('input', {bubbles:true}));"
                "target.dispatchEvent(new Event('change', {bubbles:true}));"
                "const val = (target.value || '').trim();"
                "if (!val) return {ok:false, reason:'value_not_set'};"
                "const inputRect = target.getBoundingClientRect();"
                "const candidates = Array.from(document.querySelectorAll('button, a, [role=\"button\"]'));"
                "let btn = null;"
                "let best = Infinity;"
                "for (const c of candidates) {"
                "  const r = c.getBoundingClientRect();"
                "  const s = window.getComputedStyle(c);"
                "  if (r.width < 14 || r.height < 14) continue;"
                "  if (s.display === 'none' || s.visibility === 'hidden') continue;"
                "  if (r.top > inputRect.bottom + 60 || r.bottom < inputRect.top - 60) continue;"
                "  if (r.left < inputRect.right - 20) continue;"
                "  const dx = r.left - inputRect.right;"
                "  const dy = Math.abs((r.top + r.height/2) - (inputRect.top + inputRect.height/2));"
                "  const dist = Math.max(0, dx) + dy * 2;"
                "  if (dist < best) { best = dist; btn = c; }"
                "}"
                "if (btn) { btn.click(); return {ok:true, mode:'button', value:val}; }"
                "target.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', code:'Enter', bubbles:true}));"
                "target.dispatchEvent(new KeyboardEvent('keyup', {key:'Enter', code:'Enter', bubbles:true}));"
                "return {ok:true, mode:'enter', value:val};",
                address
            )

            if result and result.get("ok"):
                logging.info(f"JS検索入力成功: value='{result.get('value')}', mode={result.get('mode')}")
                return True

            logging.info(f"JS検索入力は未完了: {result}")
            return False
        except Exception as e:
            logging.info(f"JS入力検索ルートで例外: {e}")
            return False

    def _wait_initial_global_input(self, driver: WebDriver, timeout: int = 8):
        end_time = time.time() + timeout
        while time.time() < end_time:
            element = self._find_initial_global_input(driver)
            if element is not None:
                logging.info("初期画面候補探索(Python)で入力欄を検出しました")
                return element
            time.sleep(0.3)
        return None

    def _find_initial_input_direct(self, driver: WebDriver):
        """初期画面で見えている上部入力欄をSeleniumで直接取得"""
        try:
            # 初期画面では type=text の入力欄が主検索欄として現れる
            candidates = driver.find_elements(By.XPATH, "//input[@type='text']")
            if not candidates:
                candidates = driver.find_elements(By.XPATH, "//input[@type='search']")
            best = None
            best_score = -1
            for element in candidates:
                try:
                    if not element.is_displayed() or not element.is_enabled():
                        continue
                    loc = element.location
                    size = element.size
                    width = int(size.get('width', 0))
                    height = int(size.get('height', 0))
                    top = int(loc.get('y', 9999))
                    if top > 260 or width < 220 or height < 18:
                        continue
                    score = width
                    input_type = (element.get_attribute('type') or '').strip()
                    if input_type == 'search':
                        score += 200
                    placeholder = (element.get_attribute('placeholder') or '').strip()
                    if 'スポット' in placeholder or '住所' in placeholder:
                        score += 800
                    if score > best_score:
                        best = element
                        best_score = score
                except Exception:
                    continue
            if best is not None:
                self._log_element_summary(best, "初期画面入力欄(直接)を採用")
                logging.info(f"初期画面入力欄候補を採用: score={best_score}")
            return best
        except Exception:
            return None

    def _find_initial_global_input(self, driver: WebDriver):
        try:
            candidates = driver.find_elements(By.XPATH, "//input[@type='text' or @type='search']")
            best = None
            best_score = -1

            for element in candidates:
                try:
                    if not element.is_displayed() or not element.is_enabled():
                        continue

                    location = element.location
                    size = element.size
                    width = size.get('width', 0)
                    height = size.get('height', 0)
                    y_pos = location.get('y', 9999)

                    # 初期画面上部の広い検索欄を優先
                    if y_pos > 220 or width < 220 or height < 18:
                        continue

                    placeholder = (element.get_attribute("placeholder") or "").strip()
                    input_type = (element.get_attribute("type") or "").strip()

                    score = width
                    if input_type == 'search':
                        score += 200
                    if 'スポット' in placeholder or '住所' in placeholder:
                        score += 800

                    if score > best_score:
                        best_score = score
                        best = element
                except Exception:
                    continue

            if best is not None:
                self._log_element_summary(best, "初期画面候補から使用要素を決定")
            return best
        except Exception:
            return None

    def _find_initial_global_input_js(self, driver: WebDriver):
        """JSで初期画面の上部入力欄を幾何情報から特定"""
        try:
            element = driver.execute_script(
                "const inputs = Array.from(document.querySelectorAll('input[type=\"text\"], input[type=\"search\"]'));"
                "let best = null;"
                "let bestScore = -1;"
                "for (const el of inputs) {"
                "  const rect = el.getBoundingClientRect();"
                "  const style = window.getComputedStyle(el);"
                "  const visible = rect.width > 220 && rect.height > 18 && rect.top >= 0 && rect.top < 220 && style.display !== 'none' && style.visibility !== 'hidden';"
                "  if (!visible || el.disabled) continue;"
                "  const ph = (el.getAttribute('placeholder') || '');"
                "  let score = rect.width;"
                "  if ((el.getAttribute('type') || '') === 'search') score += 200;"
                "  if (ph.includes('スポット') || ph.includes('住所')) score += 800;"
                "  if (score > bestScore) { bestScore = score; best = el; }"
                "}"
                "return best;"
            )
            if element is not None:
                self._log_element_summary(element, "初期画面候補(JS)から使用要素を決定")
            return element
        except Exception as e:
            logging.info(f"初期画面入力欄のJS探索に失敗: {e}")
            return None

    def _find_top_search_input(self, driver: WebDriver):
        """上部の住所検索ボックスを優先的に取得"""
        try:
            element = driver.execute_script(
                "const inputs = Array.from(document.querySelectorAll('input[type=\"search\"], input[type=\"text\"]'));"
                "let best = null;"
                "let bestScore = -1;"
                "for (const el of inputs) {"
                "  const rect = el.getBoundingClientRect();"
                "  const style = window.getComputedStyle(el);"
                "  const visible = rect.width > 180 && rect.height > 18 && rect.top >= 0 && rect.top < 240 && style.display !== 'none' && style.visibility !== 'hidden';"
                "  if (!visible || el.disabled) continue;"
                "  const ph = (el.getAttribute('placeholder') || '');"
                "  let score = rect.width;"
                "  if ((el.getAttribute('type') || '') === 'search') score += 100;"
                "  if (ph.includes('スポット') || ph.includes('住所') || ph.includes('検索')) score += 500;"
                "  if (score > bestScore) { bestScore = score; best = el; }"
                "}"
                "return best;"
            )

            if element is not None:
                self._log_element_summary(element, "上部検索欄候補から使用要素を決定しました")
                return element
            logging.info("上部検索欄候補が見つかりませんでした")
        except Exception:
            pass
        return None

    def _log_search_inputs(self, driver: WebDriver, phase: str) -> None:
        if not self.detailed_logging:
            return
        try:
            inputs = driver.find_elements(By.XPATH, "//input")
            logging.info(f"[MapFan入力欄デバッグ] {phase}: input要素数={len(inputs)}")
            for index, element in enumerate(inputs[:12], start=1):
                try:
                    placeholder = (element.get_attribute("placeholder") or "").strip()
                    aria_label = (element.get_attribute("aria-label") or "").strip()
                    input_type = (element.get_attribute("type") or "").strip()
                    value = (element.get_attribute("value") or "").strip()
                    is_displayed = element.is_displayed()
                    is_enabled = element.is_enabled()
                    location = element.location
                    size = element.size
                    logging.info(
                        f"[input {index}] type={input_type} displayed={is_displayed} enabled={is_enabled} "
                        f"placeholder='{placeholder}' aria='{aria_label}' value='{value}' "
                        f"x={location.get('x')} y={location.get('y')} w={size.get('width')} h={size.get('height')}"
                    )
                except Exception as item_error:
                    logging.info(f"[input {index}] 情報取得失敗: {item_error}")
        except Exception as e:
            logging.warning(f"入力欄デバッグログ取得に失敗: {e}")

    def _log_element_summary(self, element, label: str) -> None:
        if not self.detailed_logging:
            return
        try:
            placeholder = (element.get_attribute("placeholder") or "").strip()
            aria_label = (element.get_attribute("aria-label") or "").strip()
            value = (element.get_attribute("value") or "").strip()
            logging.info(
                f"{label}: displayed={element.is_displayed()} enabled={element.is_enabled()} "
                f"placeholder='{placeholder}' aria='{aria_label}' value='{value}'"
            )
        except Exception as e:
            logging.info(f"{label}: 要素サマリ取得に失敗: {e}")

    def _find_search_submit_button(self, driver: WebDriver, wait: WebDriverWait, search_input):
        """上部検索欄に対応する検索ボタンを取得"""
        try:
            # 入力欄右側の虫眼鏡アイコンを最優先で取得（clearボタン除外）
            button = driver.execute_script(
                "const input = arguments[0];"
                "const r = input.getBoundingClientRect();"
                "const x = Math.floor(r.right + 12);"
                "const y = Math.floor(r.top + r.height / 2);"
                "let el = document.elementFromPoint(x, y);"
                "while (el && !['BUTTON','A'].includes(el.tagName)) { el = el.parentElement; }"
                "if (!el) return null;"
                "const er = el.getBoundingClientRect();"
                "if (er.left > 430 || er.top > 220) return null;"
                "const t = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).toLowerCase();"
                "if (t.includes('clear') || t.includes('クリア') || t.includes('close') || t.includes('閉じる')) return null;"
                "return el;",
                search_input
            )
            if button is not None:
                try:
                    tag = button.tag_name
                    text = (button.text or "").strip()
                    logging.info(f"検索ボタン候補を採用: tag={tag}, text='{text}'")
                except Exception:
                    pass
                return button

            button = driver.execute_script(
                "const input = arguments[0];"
                "const container = input.closest('form, header, div') || document;"
                "const all = Array.from(container.querySelectorAll('button, a'));"
                "let found = null;"
                "let best = Infinity;"
                "const ir = input.getBoundingClientRect();"
                "for (const el of all) {"
                "  const fr = el.getBoundingClientRect();"
                "  if (fr.left > 430 || fr.top > 220) continue;"
                "  if (fr.width < 10 || fr.height < 10) continue;"
                "  const text = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).toLowerCase();"
                "  if (text.includes('clear') || text.includes('クリア') || text.includes('close') || text.includes('閉じる')) continue;"
                "  const isSearchLike = text.includes('検索') || text.includes('search') || text.includes('虫眼鏡') || text.includes('magnifier') || text.includes('find');"
                "  if (!isSearchLike && fr.left < ir.right - 4) continue;"
                "  const dx = Math.max(0, fr.left - ir.right);"
                "  const dy = Math.abs((fr.top + fr.height/2) - (ir.top + ir.height/2));"
                "  const dist = dx + dy * 2;"
                "  if (dist < best) { best = dist; found = el; }"
                "}"
                "return found;",
                search_input
            )
            if button is not None:
                try:
                    tag = button.tag_name
                    text = (button.text or "").strip()
                    logging.info(f"検索ボタン候補(フォールバック)を採用: tag={tag}, text='{text}'")
                except Exception:
                    pass
                return button
        except Exception:
            pass

        return self._find_first_clickable(driver, self.SEARCH_BUTTON_LOCATORS, wait, timeout_per_locator=1)

    def _find_neighbor_search_button(self, driver: WebDriver, search_input):
        """検索入力欄の隣接アイコン/ボタンを取得"""
        try:
            neighbor = driver.execute_script(
                "const input = arguments[0];"
                "const root = input.parentElement || document;"
                "let btn = root.querySelector('button, a[role=\"button\"], a');"
                "if (!btn) {"
                "  const parent = input.closest('form, header, div') || document;"
                "  btn = parent.querySelector('button, a[role=\"button\"], a');"
                "}"
                "return btn;",
                search_input
            )
            if neighbor is not None:
                return neighbor
        except Exception:
            pass
        return None

    def _click_left_panel_info_button(self, driver: WebDriver, address: str) -> bool:
        """左パネル検索結果のiボタンを優先クリックする"""
        try:
            key = (address or "").replace(" ", "").replace("　", "")
            key = key[:12] if key else ""

            # 即時チェック（最速経路）
            immediate_candidate = self._find_left_panel_info_button(driver, key)
            if immediate_candidate is not None:
                try:
                    t = (immediate_candidate.text or "").strip()
                    a = (immediate_candidate.get_attribute("aria-label") or "").strip()
                    title = (immediate_candidate.get_attribute("title") or "").strip()
                    logging.info(f"iボタン候補を採用: text='{t}', aria='{a}', title='{title}'")
                except Exception:
                    pass
                self._safe_click(driver, immediate_candidate)
                return True

            # まずは少し待って結果行の描画を待機
            end_time = time.time() + 2.5
            while time.time() < end_time:
                self._check_cancel()
                try:
                    candidate = self._find_left_panel_info_button(driver, key)

                    if candidate is not None:
                        try:
                            t = (candidate.text or "").strip()
                            a = (candidate.get_attribute("aria-label") or "").strip()
                            title = (candidate.get_attribute("title") or "").strip()
                            logging.info(f"iボタン候補を採用: text='{t}', aria='{a}', title='{title}'")
                        except Exception:
                            pass
                        self._safe_click(driver, candidate)
                        return True
                except Exception:
                    pass

                time.sleep(0.1)

            return False
        except Exception as e:
            logging.warning(f"左パネルiボタンの探索中にエラー: {e}")
            return False

    def _find_left_panel_info_button(self, driver: WebDriver, key: str = ""):
        try:
            return driver.execute_script(
                "const key = arguments[0];"
                "const nodes = Array.from(document.querySelectorAll('button, a'));"
                "let best = null;"
                "let bestScore = -1;"
                "for (const el of nodes) {"
                "  const r = el.getBoundingClientRect();"
                "  const st = window.getComputedStyle(el);"
                "  if (r.width < 14 || r.height < 14) continue;"
                "  if (st.display === 'none' || st.visibility === 'hidden') continue;"
                "  if (r.left > 500 || r.top > window.innerHeight - 20) continue;"
                "  const label = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).toLowerCase();"
                "  const isInfo = label.includes(' i ') || label.trim() === 'i' || label.includes('詳細') || label.includes('情報') || label.includes('info');"
                "  if (!isInfo) continue;"
                "  let score = 100;"
                "  if (label.includes('詳細') || label.includes('情報') || label.includes('info')) score += 80;"
                "  const row = el.closest('li, article, section, div');"
                "  if (row) {"
                "    const text = (row.textContent || '').replace(/\s+/g, '');"
                "    if (key && text.includes(key)) score += 120;"
                "  }"
                "  score += Math.max(0, 500 - r.left) * 0.01;"
                "  if (score > bestScore) { bestScore = score; best = el; }"
                "}"
                "return best;",
                key
            )
        except Exception:
            return None

    def _wait_for_search_result(self, driver: WebDriver, wait: WebDriverWait) -> None:
        end_time = time.time() + 12
        while time.time() < end_time:
            self._check_cancel()
            try:
                ready = driver.execute_script(
                    "if (document.querySelector(\"button[aria-label*='詳細'], button[title*='詳細']\")) return true;"
                    "if (document.querySelector(\"button[aria-label*='情報'], button[title*='情報']\")) return true;"
                    "if (document.querySelector(\"button, a\")) {"
                    "  const nodes = Array.from(document.querySelectorAll('button, a'));"
                    "  for (const el of nodes) {"
                    "    const txt = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).toLowerCase();"
                    "    if (txt.includes('info') || txt.includes('情報') || txt.trim() === 'i') return true;"
                    "  }"
                    "}"
                    "if (document.querySelector(\"li[class*='result'], div[class*='result']\")) return true;"
                    "return false;"
                )
                if ready:
                    return
            except Exception:
                pass
            time.sleep(0.15)
        raise TimeoutException("MapFan検索結果の表示を確認できませんでした")

    def _enter_map_view_if_needed(self, driver: WebDriver) -> None:
        return

    def _dismiss_blocking_overlay(self, driver: WebDriver) -> None:
        try:
            overlays = driver.find_elements(By.CSS_SELECTOR, "div.cdk-overlay-backdrop.cdk-overlay-backdrop-showing")
            if not overlays:
                return
            for overlay in overlays:
                try:
                    driver.execute_script("arguments[0].click();", overlay)
                except Exception:
                    continue
            time.sleep(0.2)
        except Exception:
            pass

    def _is_mapfan_block_page(self, driver: WebDriver, title: str = "") -> bool:
        try:
            normalized_title = (title or "").strip().lower()
            if "request could not be satisfied" in normalized_title:
                return True
            if normalized_title == "error":
                return True

            body_text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
            if "the request could not be satisfied" in body_text:
                return True
            if "generated by cloudfront" in body_text:
                return True
            if "access denied" in body_text and "cloudfront" in body_text:
                return True
        except Exception:
            pass
        return False

    def _log_block_page_diagnostics(self, driver: WebDriver, phase: str = "") -> None:
        try:
            now = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            output_dir = Path("logs") / "mapfan_debug"
            output_dir.mkdir(parents=True, exist_ok=True)

            safe_phase = (phase or "unknown").replace(" ", "_")
            base_name = f"mapfan_{safe_phase}_{now}"

            current_url = ""
            title = ""
            body_text = ""
            html = ""
            try:
                current_url = driver.current_url or ""
                title = driver.title or ""
                body_text = (driver.find_element(By.TAG_NAME, "body").text or "")[:500]
                html = driver.page_source or ""
            except Exception:
                pass

            html_path = output_dir / f"{base_name}.html"
            png_path = output_dir / f"{base_name}.png"

            if html:
                try:
                    html_path.write_text(html, encoding="utf-8")
                except Exception:
                    pass

            try:
                driver.save_screenshot(str(png_path))
            except Exception:
                pass

            logging.warning(
                f"[MapFan診断] phase={phase}, current_url={current_url}, title={title}, body_head={body_text}"
            )
            logging.warning(
                f"[MapFan診断] artifacts html={html_path}, screenshot={png_path}"
            )
        except Exception as e:
            logging.warning(f"MapFan診断ログ保存に失敗: {e}")

    def _find_first_clickable(
        self,
        driver: WebDriver,
        locators: List[Tuple[str, str]],
        wait: WebDriverWait,
        timeout_per_locator: int = 3
    ):
        for by, selector in locators:
            try:
                return WebDriverWait(driver, timeout_per_locator).until(EC.element_to_be_clickable((by, selector)))
            except Exception:
                continue
        return None

    def _find_first_present(self, driver: WebDriver, locators: List[Tuple[str, str]]):
        for by, selector in locators:
            try:
                elements = driver.find_elements(by, selector)
                if elements:
                    return elements[0]
            except Exception:
                continue
        return None

    def _safe_click(self, driver: WebDriver, element) -> None:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        try:
            element.click()
            return
        except Exception:
            pass

        try:
            driver.execute_script("arguments[0].click();", element)
            return
        except Exception:
            pass

        ActionChains(driver).move_to_element(element).click().perform()


def get_mapfan_detail_url(
    address: str,
    debug: bool = True,
    auto_close: Optional[bool] = None,
    force_headless: Optional[bool] = None,
    cancel_event: Optional[threading.Event] = None
) -> Optional[str]:
    """MapFan詳細URL取得の簡易エントリーポイント"""
    service = MapfanService(debug=debug)
    return service.get_detail_url_from_address(
        address=address,
        auto_close=auto_close,
        force_headless=force_headless,
        cancel_event=cancel_event
    )