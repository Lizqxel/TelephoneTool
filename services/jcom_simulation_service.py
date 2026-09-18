"""SeleniumによるJ:COM料金シミュレーション実画面の状態遷移。"""

from __future__ import annotations

from datetime import datetime
import base64
from io import BytesIO
import logging
import os
from pathlib import Path
import re
import tempfile
import threading
import time
from typing import Callable, Iterable, Optional, Tuple

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType
from PIL import Image
from services.web_driver import _resolve_chromedriver_executable

from services.jcom_simulation_models import (
    AddressApproximation,
    AddressCandidate,
    AddressCandidateRequest,
    DiscountLine,
    JcomSearchCriteria,
    JcomSimulationResult,
    JST,
    ResidenceType,
    SimulationStatus,
)
from utils.jcom_address_matcher import (
    MISSING_ADDRESS,
    NEXT_WITH_CURRENT,
    choose_address_candidate,
    classify_special_candidate,
    confirmed_address_matches,
    normalize_address_for_match,
    remaining_address,
)


JCOM_URL = "https://onlineshop.jcom.co.jp/"
YEN_RE = re.compile(r"(?P<sign>[-−▲]?)\s*(?P<amount>\d[\d,，]*)\s*円")


class JcomCancelled(Exception):
    pass


class JcomOutOfArea(Exception):
    pass


class JcomAddressNotFound(Exception):
    pass


class JcomCourseUnavailable(Exception):
    pass


class JcomSiteError(Exception):
    pass


def address_zero_result_is_final(body_text: str) -> bool:
    """J:COMの非同期検索中に一時表示される0件を確定結果と混同しない。"""
    text = body_text or ""
    has_zero = bool(re.search(r"郵便番号の検索結果は\s*0件", text))
    loading = any(
        marker in text
        for marker in ("物件情報を取得しています", "住所取得中")
    )
    return has_zero and not loading


def _yen(text: str) -> Optional[int]:
    match = YEN_RE.search(text or "")
    if not match:
        return None
    amount = int(match.group("amount").replace(",", "").replace("，", ""))
    return -amount if match.group("sign") else amount


def extract_pricing_from_text(text: str) -> dict:
    """料金結果領域だけの可視テキストから、既知項目と未知行を抽出する。"""
    raw = (text or "").replace("\r", "")
    if not raw.strip() or "----" in raw:
        raise ValueError("料金結果が読み込み中または空です。")
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw.splitlines() if line.strip()]

    data = {
        "next_month_price_yen": None,
        "discount_period_text": "",
        "base_price_yen": None,
        "contract_period": "",
        "line_type": "",
        "course": "",
        "discounts": [],
        "benefits_text": "",
        "notes_text": "",
    }
    seen_discount = set()
    benefit_lines, note_lines = [], []
    monthly_heading_index = next(
        (index for index, line in enumerate(lines) if "月額利用料金" in line or "加入翌月" in line),
        None,
    )
    if monthly_heading_index is not None:
        monthly_end = next(
            (
                index for index in range(monthly_heading_index + 1, len(lines))
                if "料金内訳" in lines[index]
            ),
            min(len(lines), monthly_heading_index + 7),
        )
        monthly_lines = lines[monthly_heading_index + 1:monthly_end]
        for index, line in enumerate(monthly_lines):
            if re.fullmatch(r"\d[\d,，]*", line) and any(
                "カ月" in following for following in monthly_lines[index + 1:index + 3]
            ):
                data["next_month_price_yen"] = int(
                    line.replace(",", "").replace("，", "")
                )
                break
        if data["next_month_price_yen"] is None:
            data["next_month_price_yen"] = next(
                (_yen(line) for line in monthly_lines if _yen(line) is not None),
                None,
            )

    for index, line in enumerate(lines):
        context = " ".join(lines[max(0, index - 1): index + 2])
        if "J:COM NET 光(N)" in line:
            data["line_type"] = "J:COM NET 光(N)"
        if ("光(N)" in line and "1G" in line) or "[ネット]1Gコース" in line:
            data["course"] = "光(N) 1Gコース"
            data["line_type"] = "J:COM NET 光(N)"
        if "光 1Gコース on auひかり" in line:
            data["course"] = "光 1Gコース on auひかり"
            data["line_type"] = "光 on auひかり"
        if not data["contract_period"]:
            contract = re.search(r"(?:\d+カ月|\d+年)契約", line)
            if contract:
                data["contract_period"] = contract.group(0)
        if data["next_month_price_yen"] is None and (
            "加入翌月" in context or "月額" in context
        ) and "基本料金" not in context:
            data["next_month_price_yen"] = _yen(line) if _yen(line) is not None else _yen(context)
        if "基本料金" in line or (
            index and lines[index - 1].rstrip("：:") == "基本料金"
        ):
            value = _yen(line)
            if value is None and index + 1 < len(lines):
                value = _yen(lines[index + 1])
            if value is not None:
                data["base_price_yen"] = value
        if data["base_price_yen"] is None and (
            "[テレビ]なし" in line or "[ネット]1Gコース" in line
        ):
            for following in lines[index + 1:index + 5]:
                value = _yen(following)
                if value is not None:
                    data["base_price_yen"] = value
                    break
        if data["base_price_yen"] is None and "光 1Gコース on auひかり" in line:
            for following in lines[index + 1:index + 5]:
                if "割" in following or "特典" in following:
                    break
                value = _yen(following)
                if value is not None and value >= 0:
                    data["base_price_yen"] = value
                    break
        period = re.search(r"\d+\s*カ月間", line)
        if period and not data["discount_period_text"]:
            data["discount_period_text"] = period.group(0).replace(" ", "")
        if "割" in line or "値引" in line:
            block_lines = [line]
            for following in lines[index + 1:index + 3]:
                if "割" in following or "値引" in following:
                    break
                block_lines.append(following)
            discount_block = " ".join(block_lines)
            amount = _yen(discount_block)
            if amount is None:
                continue
            block_period = re.search(r"\d+\s*カ月間", discount_block)
            key = re.sub(r"\s+", "", discount_block)
            if key not in seen_discount:
                seen_discount.add(key)
                data["discounts"].append(
                    DiscountLine(
                        name=line.split("：", 1)[0],
                        amount_yen=amount,
                        period_text=(block_period.group(0) if block_period else ""),
                        raw_text=discount_block,
                    )
                )
        if "キャッシュバック" in line or "特典" in line or "エントリー" in line:
            benefit_lines.append(line)
        if any(word in line for word in ("注意", "条件", "別途", "対象外")):
            note_lines.append(line)
    data["benefits_text"] = "\n".join(dict.fromkeys(benefit_lines))
    data["notes_text"] = "\n".join(dict.fromkeys(note_lines))
    if data["next_month_price_yen"] is None and data["base_price_yen"] is None:
        raise ValueError("料金金額を構造化できませんでした。")
    return data


class JcomSimulationService:
    """検索1件につき専用WebDriverを所有する。"""

    def __init__(
        self,
        cancel_event: threading.Event,
        progress: Callable[[str], None],
        candidate_resolver: Optional[Callable[[AddressCandidateRequest], Optional[str]]] = None,
        driver_factory: Optional[Callable[[], object]] = None,
        timeout: int = 25,
        screenshot_dir: Optional[Path] = None,
    ):
        self.cancel_event = cancel_event
        self.progress = progress
        self.candidate_resolver = candidate_resolver
        self.driver_factory = driver_factory
        self.timeout = timeout
        self.screenshot_dir = screenshot_dir or self._default_screenshot_dir()
        self.driver = None
        self._profile_dir = None
        self._selected_course_name = "光(N) 1Gコース"
        self._selected_line_type = "J:COM NET 光(N)"
        # Workerから設定される段階通知。料金テキストを画像生成より先に返す。
        self.result_ready_callback = None
        self.screenshot_ready_callback = None

    @staticmethod
    def _default_screenshot_dir() -> Path:
        base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
        return base / "TelephoneTool" / "jcom_results"

    def _prune_screenshots(self):
        """住所を含みうる画像を7日・最大20件に制限する。"""
        try:
            files = sorted(
                self.screenshot_dir.glob("*.png"),
                key=lambda path: path.stat().st_mtime,
                reverse=True,
            )
            cutoff = time.time() - 7 * 24 * 60 * 60
            for index, path in enumerate(files):
                if index >= 20 or path.stat().st_mtime < cutoff:
                    path.unlink(missing_ok=True)
        except Exception as exc:
            logging.warning("J:COM結果画像の保存期間整理に失敗: %s", type(exc).__name__)

    def _capture_result_screenshot(
        self, element, path: Path, confirmed_address: str = ""
    ) -> bool:
        """サイトの確定住所欄と料金領域を、読みやすい1枚にまとめる。"""
        self._screenshot_address_included = False
        hidden_marker = "data-telephonetool-capture-hidden"
        try:
            address_element = None
            if confirmed_address:
                address_boxes = self.driver.execute_script(
                    """
                    return Array.from(document.querySelectorAll('.selected-address-box'))
                        .filter((el) => el.getClientRects().length &&
                            getComputedStyle(el).visibility !== 'hidden')
                        .map((el) => ({element: el, text: el.innerText || ''}));
                    """
                )
                expected = normalize_address_for_match(confirmed_address)
                address_element = next(
                    (
                        item["element"] for item in address_boxes
                        if expected and expected in normalize_address_for_match(item["text"])
                    ),
                    None,
                )
            metrics = self.driver.execute_script(
                """
                const target = arguments[0];
                const marker = arguments[1];
                const address = arguments[2];
                document.querySelectorAll('*').forEach((node) => {
                    const style = window.getComputedStyle(node);
                    if ((style.position === 'fixed' || style.position === 'sticky') &&
                        !node.contains(target) && !target.contains(node) &&
                        !(address && (node.contains(address) || address.contains(node)))) {
                        node.setAttribute(marker, node.style.visibility || '');
                        node.style.visibility = 'hidden';
                    }
                });
                const bounds = (el) => {
                    if (!el) return null;
                    const rect = el.getBoundingClientRect();
                    return {
                        x: Math.max(0, rect.left + window.scrollX),
                        y: Math.max(0, rect.top + window.scrollY),
                        width: Math.ceil(Math.max(rect.width, el.scrollWidth)),
                        height: Math.ceil(Math.max(rect.height, el.scrollHeight))
                    };
                };
                return {
                    result: bounds(target),
                    address: bounds(address)
                };
                """,
                element,
                hidden_marker,
                address_element,
            )

            def capture(bounds):
                width = min(4000, max(1, int(bounds["width"])))
                height = min(12000, max(1, int(bounds["height"])))
                captured = self.driver.execute_cdp_cmd(
                    "Page.captureScreenshot",
                    {
                        "format": "png",
                        "fromSurface": True,
                        "captureBeyondViewport": True,
                        "clip": {
                            "x": float(bounds["x"]),
                            "y": float(bounds["y"]),
                            "width": float(width),
                            "height": float(height),
                            "scale": 1,
                        },
                    },
                )
                return base64.b64decode(captured.get("data", ""))

            result_png = capture(metrics["result"])
            if not result_png:
                return False
            address_bounds = metrics.get("address")
            if address_bounds and address_bounds["width"] > 0 and address_bounds["height"] > 0:
                try:
                    address_png = capture(address_bounds)
                    with Image.open(BytesIO(address_png)) as address_image, Image.open(
                        BytesIO(result_png)
                    ) as result_image:
                        width = max(address_image.width, result_image.width)
                        gap = 16
                        combined = Image.new(
                            "RGB", (width, address_image.height + gap + result_image.height), "white"
                        )
                        combined.paste(address_image.convert("RGB"), (0, 0))
                        combined.paste(
                            result_image.convert("RGB"), (0, address_image.height + gap)
                        )
                        combined.save(path, format="PNG")
                    self._screenshot_address_included = True
                except Exception as exc:
                    logging.warning("J:COM住所欄の画像取得に失敗: %s", type(exc).__name__)
                    path.write_bytes(result_png)
            else:
                path.write_bytes(result_png)
            return path.is_file() and path.stat().st_size > 0
        except Exception as exc:
            logging.warning("J:COM結果全体の範囲取得に失敗: %s", type(exc).__name__)
            try:
                return bool(element.screenshot(str(path)))
            except Exception:
                return False
        finally:
            try:
                self.driver.execute_script(
                    """
                    const marker = arguments[0];
                    document.querySelectorAll('[' + marker + ']').forEach((node) => {
                        node.style.visibility = node.getAttribute(marker) || '';
                        node.removeAttribute(marker);
                    });
                    """,
                    hidden_marker,
                )
            except Exception:
                pass

    def _create_driver(self):
        if self.driver_factory:
            return self.driver_factory()
        self._profile_dir = tempfile.TemporaryDirectory(prefix="telephonetool-jcom-")
        options = Options()
        options.page_load_strategy = "eager"
        options.add_argument(f"--user-data-dir={self._profile_dir.name}")
        options.add_argument("--incognito")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--lang=ja-JP")
        options.add_argument("--disable-sync")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        # Chrome自身の診断ログ（DevTools/GCM/Updater/最適化モデル）を
        # TelephoneToolのコンソールへ流さない。ページ通信は無効にしない。
        options.add_argument("--log-level=3")
        options.add_argument("--silent")
        options.add_argument("--disable-background-networking")
        options.add_argument("--disable-component-update")
        options.add_argument(
            "--disable-features="
            "OptimizationGuideModelDownloading,OptimizationHints,"
            "OptimizationTargetPrediction,OptimizationGuideOnDeviceModel,PushMessaging"
        )
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        # 既存の提供判定ブラウザと同じ初期サイズにする。
        options.add_argument("--window-size=800,600")
        options.add_experimental_option(
            "prefs",
            {
                "profile.managed_default_content_settings.images": 1,
                "profile.default_content_setting_values.images": 1,
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
            },
        )
        try:
            driver = webdriver.Chrome(
                service=Service(log_output=os.devnull),
                options=options,
            )
        except WebDriverException:
            # Selenium Managerが利用できない配布環境では既存依存の
            # webdriver-managerでChrome本体に合うドライバーを解決する。
            installed = ChromeDriverManager(chrome_type=ChromeType.GOOGLE).install()
            executable = _resolve_chromedriver_executable(installed)
            driver = webdriver.Chrome(
                service=Service(executable, log_output=os.devnull),
                options=options,
            )
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(0)
        return driver

    def _check_cancelled(self):
        if self.cancel_event.is_set():
            raise JcomCancelled()

    def _wait(self, condition, timeout: Optional[int] = None):
        end = time.monotonic() + (timeout or self.timeout)
        last_error = None
        next_url_check = 0.0
        while (now := time.monotonic()) < end:
            self._check_cancelled()
            if now >= next_url_check:
                current_url = self.driver.current_url
                if "/Err" in current_url:
                    match = re.search(r"[?&]errorId=([^&]+)", current_url)
                    raise JcomSiteError(match.group(1) if match else "不明")
                next_url_check = now + 1.0
            try:
                value = condition(self.driver)
                if value:
                    return value
            except (NoSuchElementException, StaleElementReferenceException) as exc:
                last_error = exc
            self.cancel_event.wait(0.2)
        raise TimeoutException(str(last_error or "待機がタイムアウトしました"))

    @staticmethod
    def _visible(elements):
        return [element for element in elements if element.is_displayed()]

    def _click_if_not_selected(self, by: By, selector: str):
        # React画面では実inputをCSSで隠し、labelだけを表示する箇所がある。
        # DOM上の操作可能なinputを対象にし、表示状態そのものは要求しない。
        element = self._wait(
            lambda d: next((e for e in d.find_elements(by, selector) if e.is_enabled()), None)
        )
        if not element.is_selected():
            self.driver.execute_script("arguments[0].click()", element)
        return self._wait(
            lambda d: next(
                (e for e in d.find_elements(by, selector) if e.is_enabled() and e.is_selected()),
                None,
            ),
            5,
        )

    def _click_button_with_text(self, texts: Iterable[str], scope=None):
        root = scope or self.driver
        expected = tuple(re.sub(r"\s+", "", text) for text in texts)

        def find(_):
            # 多数のリンクを1件ずつSeleniumへ問い合わせず、同一DOM状態で探す。
            return self.driver.execute_script(
                """
                const root = arguments[0] || document;
                const expected = arguments[1];
                return Array.from(root.querySelectorAll('button, a, [role="button"]'))
                    .find((el) => {
                        if (!el.getClientRects().length ||
                            getComputedStyle(el).visibility === 'hidden' ||
                            el.matches(':disabled') ||
                            el.getAttribute('aria-disabled') === 'true') return false;
                        const label = (el.innerText || el.textContent || '')
                            .replace(/\\s+/g, '');
                        return expected.some((text) => label.includes(text));
                    }) || null;
                """,
                None if root is self.driver else root,
                expected,
            )

        element = self._wait(find)
        self.driver.execute_script("arguments[0].click()", element)
        return element

    def _setup_area(self, criteria: JcomSearchCriteria):
        self.progress("J:COM画面を開いています")
        self.driver.get(JCOM_URL)
        self.progress("対象サービスを設定しています")
        self._click_if_not_selected(By.CSS_SELECTOR, "input#select-net")
        for selector in ("#select-tv", "#select-phone", "#select-electricity"):
            for element in self.driver.find_elements(By.CSS_SELECTOR, selector):
                if element.is_selected():
                    self.driver.execute_script("arguments[0].click()", element)

        zip_input = self._wait(lambda d: next((e for e in d.find_elements(By.CSS_SELECTOR, "input#zip-code") if e.is_displayed()), None))
        zip_input.clear()
        zip_input.send_keys(criteria.postal_code)
        if re.sub(r"\D", "", zip_input.get_attribute("value") or "") != criteria.postal_code:
            raise RuntimeError("郵便番号の入力状態を確認できません。")
        self._click_if_not_selected(
            By.CSS_SELECTOR,
            f"input[name='house-type'][value='{criteria.residence_type.value}']",
        )
        self._click_button_with_text(("エリアを設定する", "エリアを検索"))

    def _candidate_elements(self, selected_address: str = ""):
        # 表示中のモーダル/住所選択領域だけを対象にする。
        # 本文・現在のモーダル・候補を同一のDOM状態から一度に読む。
        snapshot = self.driver.execute_script(
            r"""
            const selected = arguments[0] || '';
            const visible = (el) => !!el.getClientRects().length &&
                getComputedStyle(el).visibility !== 'hidden';
            const roots = Array.from(document.querySelectorAll(
                '[role="dialog"], .modal, .address-modal')).filter(visible);
            const scope = roots.at(-1) || document;
            // 横スライドで隠れた前段列もgetClientRects()は残るため、
            // 選択済み住所があれば現在の列だけを候補として読む。
            const columns = scope.querySelector('.address-narrow-down');
            let active = scope;
            if (columns) {
                const selector = selected
                    ? '.js-narrow-down-2nd' : '.js-narrow-down-first';
                active = columns.querySelector(selector) || scope;
            }
            const collect = (selector) => Array.from(active.querySelectorAll(selector))
                .filter(visible);
            let nodes = [];
            if (selected.split('|').at(-1).endsWith('番地')) {
                nodes = collect('.collapse-container.is-open '
                    + '.sub-address-link-list .link-bullet-black');
            }
            if (!nodes.length) {
                nodes = collect('a.link-bullet-black, button.link-bullet-black, '
                    + '.link-bullet-black, .accordion-header.collapse-trigger, a:not([href])');
            }
            for (const el of Array.from(scope.querySelectorAll('button, a')).filter(visible)) {
                    const text = (el.innerText || '').replace(/\\s+/g, ' ').trim();
                    if ((text.includes('表示中の住所') && text.includes('次へ')) ||
                        text.includes('住所・物件がない') ||
                        text.includes('住所および物件が見つからない') ||
                        text.includes('該当する住所がない')) {
                        if (!nodes.includes(el)) nodes.push(el);
                    }
            }
            return {
                bodyText: document.body.innerText || '',
                entries: nodes.map((el) => ({
                    element: el,
                    text: (el.innerText || '').replace(/\\s+/g, ' ').trim(),
                    className: el.className || '',
                    noHref: !el.hasAttribute('href'),
                })).filter((item) => item.text),
            };
            """,
            selected_address,
        )
        self._last_address_body_text = snapshot["bodyText"]
        result = []
        for entry in snapshot["entries"]:
            element = entry["element"]
            text = entry["text"]
            special = classify_special_candidate(text)
            if (
                special
                or "link-bullet-black" in entry["className"]
                or "accordion-header" in entry["className"]
                or entry["noHref"]
            ):
                element._jcom_candidate_text = text
                result.append(element)
        return None, result

    @staticmethod
    def _candidate_text(element):
        return getattr(element, "_jcom_candidate_text", None) or element.text

    @staticmethod
    def _identical_duplicate_candidate_id(candidates, element_by_id):
        """表示とDOMが同一の町域重複だけを自動で1件にまとめる。"""
        ordinary = [
            item for item in candidates
            if not item.special_action
            and item.text not in ("戻る", "閉じる", "再設定")
        ]
        if len(ordinary) != 2 or ordinary[0].text != ordinary[1].text:
            return None
        first = element_by_id[ordinary[0].candidate_id]
        second = element_by_id[ordinary[1].candidate_id]
        if first.get_attribute("outerHTML") == second.get_attribute("outerHTML"):
            return ordinary[0].candidate_id
        return None

    @staticmethod
    def _is_building_choice(criteria, selected_address, candidates):
        if criteria.residence_type is not ResidenceType.APARTMENT:
            return False
        last_selected = normalize_address_for_match(
            selected_address.split("|")[-1]
        )
        if not re.fullmatch(r"\d+(?:番地|番|号)?", last_selected):
            return False
        ordinary = [
            item for item in candidates
            if not item.special_action
            and item.text not in ("戻る", "閉じる", "再設定")
        ]
        return bool(ordinary) and all(
            not re.fullmatch(r"\d+(?:番地|番|号|号室)?", normalize_address_for_match(item.text))
            for item in ordinary
        )

    @staticmethod
    def _is_room_choice(criteria, candidates, building_selected=False):
        """建物選択後の数値候補だけを部屋番号選択として扱う。"""
        if criteria.residence_type is not ResidenceType.APARTMENT:
            return False
        ordinary = [
            item for item in candidates
            if not item.special_action
            and item.text not in ("戻る", "閉じる", "再設定")
        ]
        if not ordinary:
            return False
        normalized = [normalize_address_for_match(item.text) for item in ordinary]
        numeric_rooms = all(
            re.fullmatch(r"\d+(?:号|号室)?", text) for text in normalized
        )
        explicit_rooms = all(
            re.fullmatch(r"\d+号室", text) for text in normalized
        )
        return numeric_rooms and (building_selected or explicit_rooms)

    @staticmethod
    def _manual_choice_candidates(candidates):
        """実候補を先に、表示中住所で次へを最後に並べる。"""
        ordinary = [
            item for item in candidates
            if not item.special_action
            and item.text not in ("戻る", "閉じる", "再設定")
        ]
        next_actions = [
            item for item in candidates
            if item.special_action == NEXT_WITH_CURRENT
        ]
        return ordinary + next_actions

    def _not_join_visible(self) -> bool:
        inputs = self.driver.find_elements(By.CSS_SELECTOR, "input#notJoin")
        if not inputs:
            return False
        labels = self.driver.find_elements(By.CSS_SELECTOR, "label[for='notJoin']")
        if any(label.is_displayed() for label in labels):
            return True
        try:
            container = inputs[0].find_element(By.XPATH, "ancestor::div[.//*[contains(normalize-space(.), 'J:COMのサービス')]][1]")
            return container.is_displayed()
        except NoSuchElementException:
            return inputs[0].is_displayed()

    def _read_confirmed_address(self, fallback: str) -> str:
        selectors = (
            "#serviceSelectResult",
            ".serviceSelectResult",
            "[class*='serviceSelectResult']",
        )
        ignored = ("検討しているサービス", "ネット", "テレビ", "固定電話", "でんき")
        for selector in selectors:
            for element in self.driver.find_elements(By.CSS_SELECTOR, selector):
                if not element.is_displayed():
                    continue
                for line in element.text.splitlines():
                    candidate = line.strip()
                    if candidate and candidate not in ignored and candidate != "再設定":
                        return candidate
        return fallback

    def _select_address(self, criteria: JcomSearchCriteria) -> Tuple[str, bool]:
        self.progress("住所候補を確認しています")
        selected = ""
        matched_selected = ""
        partial = False
        building_selected = False
        previous_signatures = set()
        for stage_number in range(20):
            self._check_cancelled()
            zero_result_state = {"since": None}

            def candidates_ready(_):
                root, elements = self._candidate_elements(selected)
                body = self._last_address_body_text
                if "提供エリア外" in body or "サービス提供エリア外" in body:
                    raise JcomOutOfArea()
                actionable = []
                for element in elements:
                    text = re.sub(r"\s+", "", self._candidate_text(element) or "")
                    special = classify_special_candidate(text)
                    if special == MISSING_ADDRESS or text in ("戻る", "閉じる"):
                        continue
                    actionable.append(element)
                if actionable:
                    zero_result_state["since"] = None
                    return root, elements
                if address_zero_result_is_final(body):
                    now = time.monotonic()
                    if zero_result_state["since"] is None:
                        zero_result_state["since"] = now
                    elif now - zero_result_state["since"] >= 2.0:
                        raise JcomAddressNotFound()
                else:
                    zero_result_state["since"] = None
                return None

            try:
                root, elements = self._wait(candidates_ready, 20)
            except TimeoutException:
                # モーダルが閉じ、利用状況へ進める状態なら住所確定。
                if self._not_join_visible():
                    confirmed = self._read_confirmed_address(criteria.address_original)
                    confirmed_full = confirmed_address_matches(
                        criteria.address_original,
                        confirmed,
                        criteria.address_components,
                    )
                    return confirmed, partial or not confirmed_full
                raise

            candidates = []
            element_by_id = {}
            for index, element in enumerate(elements):
                candidate_id = f"{stage_number}:{index}"
                text = re.sub(r"\s+", " ", self._candidate_text(element) or "").strip()
                candidate = AddressCandidate(candidate_id, text, classify_special_candidate(text))
                candidates.append(candidate)
                element_by_id[candidate_id] = element
            signature = tuple((item.text, item.special_action) for item in candidates)
            state_signature = (selected, signature)
            if state_signature in previous_signatures:
                raise RuntimeError("同じ住所候補が繰り返されたため停止しました。")
            previous_signatures.add(state_signature)

            selection_kind = ""
            if self._is_building_choice(criteria, selected, candidates):
                selection_kind = "building"
            elif self._is_room_choice(criteria, candidates, building_selected):
                selection_kind = "room"
            if selection_kind:
                candidate_id = None
                decision = choose_address_candidate(
                    criteria.address_original,
                    matched_selected,
                    candidates,
                    criteria.address_components,
                )
            else:
                decision = choose_address_candidate(
                    criteria.address_original,
                    matched_selected,
                    candidates,
                    criteria.address_components,
                )
                candidate_id = decision.candidate_id
            if decision is not None and decision.requires_user:
                # 郵便番号検索が同一DOMの町域候補を二重に返す場合がある。
                # 表示・属性・要素構造が完全に同じ2件に限り、重複表示として
                # 先頭を採用する。異なる候補は従来どおり利用者に確認する。
                duplicate_id = self._identical_duplicate_candidate_id(
                    candidates, element_by_id
                )
                if duplicate_id is not None:
                    candidate_id = duplicate_id
            if selection_kind:
                if self.candidate_resolver is None:
                    target = "建物名" if selection_kind == "building" else "部屋番号"
                    raise RuntimeError(f"集合住宅の{target}を選択してください。")
                request = AddressCandidateRequest(
                    request_id=criteria.request_id,
                    generation=criteria.generation,
                    stage_id=f"address-{stage_number}",
                    selected_address=selected.replace("|", " → "),
                    remaining_address=remaining_address(criteria.address_original, matched_selected),
                    candidates=self._manual_choice_candidates(candidates),
                    selection_kind=selection_kind,
                )
                candidate_id = self.candidate_resolver(request)
                if candidate_id is None:
                    raise JcomCancelled()
            elif candidate_id is None:
                if decision.no_match:
                    raise JcomAddressNotFound()
                raise RuntimeError(
                    "住所候補を一意に特定できませんでした。入力住所を確認してください。"
                )
            chosen = next((item for item in candidates if item.candidate_id == candidate_id), None)
            if chosen is None:
                raise RuntimeError("選択された住所候補は既に無効です。")
            if chosen.special_action == MISSING_ADDRESS:
                raise RuntimeError("住所・物件なしの問い合わせ導線は自動操作できません。")
            if chosen.special_action == NEXT_WITH_CURRENT:
                if remaining_address(criteria.address_original, matched_selected):
                    partial = True
            else:
                selected = f"{selected}|{chosen.text}" if selected else chosen.text
                matched_text = chosen.text
                if selection_kind == "building":
                    building_selected = True
                if (
                    selection_kind == "room"
                    and decision is not None
                    and decision.candidate_id != chosen.candidate_id
                ):
                    expected = decision.approximate_from
                    if not expected and decision.candidate_id:
                        expected_candidate = next(
                            (
                                item for item in candidates
                                if item.candidate_id == decision.candidate_id
                            ),
                            None,
                        )
                        expected = expected_candidate.text if expected_candidate else ""
                    if expected:
                        self._address_approximations.append(
                            AddressApproximation(expected, chosen.text)
                        )
                        matched_text = expected
                        partial = True
                elif decision is not None and decision.approximate_from:
                    self._address_approximations.append(
                        AddressApproximation(decision.approximate_from, chosen.text)
                    )
                    matched_text = decision.approximate_from
                    partial = True
                matched_selected = (
                    f"{matched_selected}|{matched_text}"
                    if matched_selected else matched_text
                )

            old_signature = signature
            element = element_by_id[candidate_id]
            self.driver.execute_script("arguments[0].click()", element)

            def stage_changed(_):
                if self._not_join_visible():
                    return True
                try:
                    _, new_elements = self._candidate_elements(selected)
                    new_signature = tuple(
                        (
                            re.sub(r"\s+", " ", self._candidate_text(e) or "").strip(),
                            classify_special_candidate(self._candidate_text(e) or ""),
                        )
                        for e in new_elements
                    )
                    return new_signature and new_signature != old_signature and "住所取得中" not in self._last_address_body_text
                except StaleElementReferenceException:
                    return False

            self._wait(stage_changed, 20)
            if self._not_join_visible():
                confirmed = self._read_confirmed_address(criteria.address_original)
                confirmed_full = confirmed_address_matches(
                    criteria.address_original,
                    confirmed,
                    criteria.address_components,
                )
                return confirmed, partial or not confirmed_full
        raise RuntimeError("住所選択の段階数が上限を超えました。")

    def _open_simulation(self):
        self.progress("提供サービスを確認しています")
        self._click_if_not_selected(By.CSS_SELECTOR, "input#notJoin")

        def result_scope(_):
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            if "提供エリア外" in body_text or "ご提供できません" in body_text:
                raise JcomOutOfArea()
            selectors = (
                "#service-result",
                ".service-result",
                "[class*='serviceSelectResult']",
                ".area-result",
                "main",
            )
            for selector in selectors:
                for element in self.driver.find_elements(By.CSS_SELECTOR, selector):
                    if element.is_displayed() and "料金シミュレーション" in element.text:
                        return element
            return None

        scope = self._wait(result_scope, 30)
        self._click_button_with_text(("料金シミュレーション",), scope=scope)
        self._wait(lambda d: "/Simulation/" in d.current_url, 30)

    def _apply_age_if_present(self, criteria: JcomSearchCriteria) -> bool:
        self.progress("年齢条件を確認しています")

        def page_ready(_):
            url = self.driver.current_url
            if "/Simulation/Simulation03" in url:
                return "age"
            if "/Simulation/Simulation05" in url:
                return "course"
            if self.driver.find_elements(By.CSS_SELECTOR, "input[name='course-net']"):
                return "course"
            return None

        state = self._wait(page_ready, 30)
        if state == "course":
            return False

        target_label = criteria.age_bracket.value
        labels = self._visible(self.driver.find_elements(By.CSS_SELECTOR, "label"))
        label = next((item for item in labels if target_label in re.sub(r"\s+", "", item.text)), None)
        if label is None:
            raise RuntimeError(f"年齢区分「{target_label}」を選択できません。")
        target_id = label.get_attribute("for")
        input_element = self.driver.find_element(By.ID, target_id) if target_id else label.find_element(By.CSS_SELECTOR, "input")
        if not input_element.is_selected():
            self.driver.execute_script("arguments[0].click()", label)
        if target_id:
            self._wait(
                lambda d: next(
                    (e for e in d.find_elements(By.ID, target_id) if e.is_selected()),
                    None,
                ),
                5,
            )
        else:
            self._wait(lambda _: input_element.is_selected(), 5)
        self._click_button_with_text(("料金シミュレーションを開始", "シミュレーションを開始"))
        self._wait(lambda d: "/Simulation/Simulation05" in d.current_url, 30)
        return True

    def _select_course_and_result(self):
        self.progress("光(N) 1Gコースを選択しています")
        unavailable_state = {"since": None, "signature": None}

        def exact_course_input(driver):
            all_course_inputs = driver.find_elements(
                By.CSS_SELECTOR, "input[name='course-net']"
            )
            listed_courses = []
            for input_element in all_course_inputs:
                if input_element.get_attribute("value") != "1G":
                    continue
                input_id = input_element.get_attribute("id")
                label_text = ""
                if input_id:
                    labels = driver.find_elements(By.CSS_SELECTOR, f"label[for='{input_id}']")
                    label_text = " ".join(label.text for label in labels)
                try:
                    label_text += " " + input_element.find_element(
                        By.XPATH, "ancestor::*[self::label or self::div][1]"
                    ).text
                except NoSuchElementException:
                    pass
                normalized = re.sub(r"\s+", "", label_text).lower()
                listed_courses.append((normalized, input_element.is_enabled()))
                if (
                    input_element.is_enabled()
                    and "光(n)1gコース" in normalized
                    and "auひかり" not in normalized
                ):
                    return input_element
            # Simulation05の枠やauひかり等だけが先に表示される場合もある。
            # コース一覧の状態が8秒間変わらないことを確認してから判定不能にする。
            body_text = driver.find_element(By.TAG_NAME, "body").text
            if (
                "/Simulation/Simulation05" in driver.current_url
                and "コースをお選びください" in body_text
            ):
                signature = tuple(listed_courses)
                now = time.monotonic()
                if signature != unavailable_state["signature"]:
                    unavailable_state["signature"] = signature
                    unavailable_state["since"] = now
                elif now - unavailable_state["since"] >= 8.0:
                    raise JcomCourseUnavailable(
                        "判定不能：光(N) 1Gコースを確認できませんでした。"
                        "auひかり1Gへの代替はしません。"
                    )
            else:
                unavailable_state["since"] = None
                unavailable_state["signature"] = None
            return None

        target = self._wait(
            exact_course_input,
            30,
        )
        if not target.is_selected():
            self.driver.execute_script("arguments[0].click()", target)
        target_id = target.get_attribute("id")
        if target_id:
            self._wait(
                lambda d: next(
                    (e for e in d.find_elements(By.ID, target_id) if e.is_selected()),
                    None,
                ),
                5,
            )
        else:
            self._wait(lambda _: target.is_selected(), 5)

        stability = {"text": None, "since": 0.0}
        result_selector = ".simulation-result-wrap, #simulation-result"

        def selected_course_is_visible(element):
            text = element.text
            return "J:COM NET 光(N)" in text and "1Gコース" in text

        def stable_result(_):
            elements = self._visible(self.driver.find_elements(By.CSS_SELECTOR, result_selector))
            if not elements:
                return None
            text = elements[-1].text.strip()
            if not text or "----" in text or "シミュレーション結果" not in text:
                return None
            if text != stability["text"]:
                stability["text"] = text
                stability["since"] = time.monotonic()
                return None
            if time.monotonic() - stability["since"] < 0.8:
                return None
            return elements[-1]

        result_element = self._wait(stable_result, 40)

        triggers = self._visible(
            result_element.find_elements(
                By.XPATH,
                ".//*[self::button or self::a or self::div]"
                "[normalize-space()='料金内訳' or normalize-space()='内訳を見る']",
            )
        )
        if not triggers:
            triggers = self._visible(
                self.driver.find_elements(
                    By.CSS_SELECTOR,
                    "div[title='料金内訳'].collapse-trigger, [title='料金内訳']",
                )
            )
        if triggers:
            trigger = triggers[-1]
            parent_text_before = result_element.text
            if "基本料金" not in parent_text_before or "内訳" not in parent_text_before:
                self.driver.execute_script("arguments[0].click()", trigger)
            result_element = self._wait(
                lambda d: next(
                    (
                        element for element in self._visible(
                            d.find_elements(By.CSS_SELECTOR, result_selector)
                        )
                        if (
                            selected_course_is_visible(element)
                            and "----" not in element.text
                        )
                    ),
                    None,
                ),
                15,
            )
        elif not selected_course_is_visible(result_element):
            raise JcomCourseUnavailable(
                "判定不能：料金結果が光(N) 1Gコースであることを確認できませんでした。"
            )
        return result_element

    def run(self, criteria: JcomSearchCriteria) -> JcomSimulationResult:
        self._selected_course_name = "光(N) 1Gコース"
        self._selected_line_type = "J:COM NET 光(N)"
        self._address_approximations = []
        result = JcomSimulationResult(
            request_id=criteria.request_id,
            generation=criteria.generation,
            acquired_at=datetime.now(JST),
            input_address=criteria.address_original,
            confirmed_address="",
            residence_type=criteria.residence_type,
            age=criteria.age,
            calculated_age_bracket=criteria.age_bracket,
            applied_age_bracket=None,
            age_selection_used=False,
            selected_service=criteria.service,
            line_type="",
            course=criteria.course,
            address_approximations=self._address_approximations,
        )
        try:
            self.driver = self._create_driver()
            self._setup_area(criteria)
            result.confirmed_address, result.partial_address = self._select_address(criteria)
            if (
                not result.partial_address
                and not confirmed_address_matches(
                    criteria.address_original,
                    result.confirmed_address,
                    criteria.address_components,
                )
            ):
                raise JcomAddressNotFound()
            self._open_simulation()
            result.age_selection_used = self._apply_age_if_present(criteria)
            if result.age_selection_used:
                result.applied_age_bracket = criteria.age_bracket
            element = self._select_course_and_result()
            result.raw_text = element.text
            result.source_url = self.driver.current_url
            parsed = extract_pricing_from_text(result.raw_text)
            for key, value in parsed.items():
                setattr(result, key, value)
            result.course = self._selected_course_name
            result.line_type = self._selected_line_type
            result.selected_service = "ネットのみ"
            result.status = SimulationStatus.PARTIAL if result.partial_address else SimulationStatus.SUCCESS
            result.screenshot_pending = True
            self.progress("料金結果を取得しました")
            if callable(self.result_ready_callback):
                try:
                    self.result_ready_callback(result)
                except Exception:
                    logging.exception("J:COM料金結果の先行通知に失敗")

            image_note = ""
            try:
                self.screenshot_dir.mkdir(parents=True, exist_ok=True)
                self._prune_screenshots()
                path = self.screenshot_dir / f"{criteria.request_id}.png"
                if self._capture_result_screenshot(element, path, result.confirmed_address):
                    result.screenshot_path = str(path)
                    if not self._screenshot_address_included:
                        image_note = "画像にサイトの住所欄を含められませんでした"
                else:
                    image_note = "画像未取得"
            except Exception as exc:
                logging.warning("J:COM料金領域の画像取得に失敗: %s", type(exc).__name__)
                image_note = "画像未取得"
            if image_note:
                result.notes_text = "\n".join(
                    part for part in (result.notes_text, image_note) if part
                )
            result.screenshot_pending = False
            if callable(self.screenshot_ready_callback):
                try:
                    self.screenshot_ready_callback(result)
                except Exception:
                    logging.exception("J:COMスクリーンショット結果の通知に失敗")
        except JcomCancelled:
            result.status = SimulationStatus.CANCELLED
            result.error_message = "検索をキャンセルしました。"
        except JcomAddressNotFound:
            result.status = SimulationStatus.ADDRESS_NOT_FOUND
            result.error_message = (
                "J:COMの住所検索で該当する住所・物件が見つかりませんでした"
                f"（選択中: {criteria.residence_type.label}）。"
                "住居タイプが正しいか確認してください。"
            )
        except JcomOutOfArea:
            result.status = SimulationStatus.OUT_OF_AREA
            result.error_message = "J:COMの提供対象外です。"
        except JcomCourseUnavailable as exc:
            result.status = SimulationStatus.COURSE_UNAVAILABLE
            result.error_message = str(exc)
            if result.partial_address:
                result.error_message += (
                    " サイト確定住所は入力住所と異なるため、"
                    "この判定は入力住所そのものの提供可否を示しません。"
                )
        except JcomSiteError as exc:
            result.status = SimulationStatus.COMMUNICATION_ERROR
            result.error_message = (
                "J:COMサイト側で受付エラーが発生しました"
                f"（エラーコード: {exc}）。時間を空けて再検索してください。"
            )
        except (TimeoutException, WebDriverException) as exc:
            result.status = SimulationStatus.COMMUNICATION_ERROR
            result.error_message = "J:COM画面の応答を確認できませんでした。"
            logging.warning("J:COM通信/待機エラー: %s", type(exc).__name__)
        except ValueError as exc:
            result.status = SimulationStatus.EXTRACTION_ERROR
            result.error_message = f"構造化取得に失敗: {exc}"
        except Exception as exc:
            result.status = SimulationStatus.ERROR
            result.error_message = str(exc)
            logging.exception("J:COM料金シミュレーションに失敗")
        finally:
            if self.driver is not None:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None
            if self._profile_dir is not None:
                try:
                    self._profile_dir.cleanup()
                except Exception:
                    pass
                self._profile_dir = None
        return result
