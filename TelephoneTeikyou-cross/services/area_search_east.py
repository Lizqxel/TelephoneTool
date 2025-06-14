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
from bs4 import BeautifulSoup
from datetime import datetime

from services.web_driver import create_driver, load_browser_settings
from utils.string_utils import normalize_string, calculate_similarity
from utils.address_utils import normalize_address
from services.area_search import take_full_page_screenshot, check_cancellation, CancellationError

# グローバル変数でブラウザドライバーを保持
global_driver = None

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
        # 住所を正規化（normalize_stringを使用して漢数字変換も含める）
        address = normalize_string(address)
        logging.info(f"正規化後の住所: {address}")
        
        # 都道府県を抽出
        prefecture_match = re.match(r'^(.+?[都道府県])', address)
        prefecture = prefecture_match.group(1) if prefecture_match else None
        block = None
        number_part = None
        town = ""
        
        if prefecture:
            remaining_address = address[len(prefecture):].strip()
            # 市区町村を抽出（郡がある場合も考慮）
            city_match = re.match(r'^(.+?郡.+?[町村]|.+?[市区町村])', remaining_address)
            city = city_match.group(1) if city_match else None
            
            if city:
                # 残りの住所から基本住所と番地を分離
                remaining = remaining_address[len(city):].strip()
                
                # 特殊な表記（大字、字、甲乙丙丁）を含む部分を抽出
                # 町名抽出のための正規表現を拡張
                # 例: "巻堀字巻堀88" → 町名: "巻堀字巻堀" 番地: "88"
                town_match = re.match(r'^(.+?字.+?)(\d+.*)?$', remaining)
                if town_match:
                    town = town_match.group(1)
                    number_part = town_match.group(2).strip() if town_match.group(2) else None
                else:
                    # 改良された丁目認識ロジック
                    chome_match = re.search(r'(\d+)丁目', remaining)
                    if chome_match:
                        block = chome_match.group(1)
                        after_chome = remaining[remaining.find('丁目') + 2:].strip()
                        number_match = re.search(r'(\d+(?:[-－]\d+)?)', after_chome)
                        if number_match:
                            number_part = number_match.group(1)
                        else:
                            number_part = None
                        
                        # 町名の抽出を改良
                        # 「恵比寿四丁目」→ 町名: "恵比寿", 丁目: "4"
                        # 「西岡四条1丁目」→ 町名: "西岡四条", 丁目: "1"
                        full_chome_text = chome_match.group(1) + '丁目'
                        chome_start = remaining.find(full_chome_text)
                        
                        # 丁目より前の部分を町名として抽出
                        town_candidate = remaining[:chome_start].strip()
                        
                        # 条がある場合は条も含めて町名とする
                        if '条' in town_candidate:
                            town = town_candidate
                        else:
                            # 通常の場合は丁目の前の部分が町名
                            town = town_candidate
                    else:
                        # 住所パターンの智能判定
                        # パターン1: 「町名＋1-2桁数字－番地」形式（丁目あり）
                        # 例: "外川町4-11162" → 町名: "外川町", 丁目: "4", 番地: "11162"
                        town_with_block_pattern = r'^(.+?町)(\d{1,2})[-－](\d{3,}(?:[-－]\d+)*)$'
                        town_block_match = re.match(town_with_block_pattern, remaining)
                        
                        if town_block_match:
                            town = town_block_match.group(1)
                            block = town_block_match.group(2)
                            number_part = town_block_match.group(3)
                            logging.info(f"丁目形式を検出: 町名={town}, 丁目={block}, 番地={number_part}")
                        else:
                            # パターン2: 「町名＋3桁以上数字－数字」形式（番地－号）
                            # 例: "北堀1870-1" → 町名: "北堀", 番地: "1870", 号: "1"
                            town_with_number_pattern = r'^(.+?)(\d{3,})[-－](\d+(?:[-－]\d+)*)$'
                            town_number_match = re.match(town_with_number_pattern, remaining)
                            
                            if town_number_match:
                                town = town_number_match.group(1).strip()
                                block = None  # 丁目はなし
                                number_part = f"{town_number_match.group(2)}-{town_number_match.group(3)}"
                                logging.info(f"番地－号形式を検出: 町名={town}, 番地={number_part}")
                            else:
                                # パターン3: より汎用的なパターン（従来のフォールバック）
                                general_pattern = r'^(.+?)(\d+)[-－](\d+(?:[-－]\d+)*)$'
                                general_match = re.match(general_pattern, remaining)
                                
                                if general_match:
                                    first_number = general_match.group(2)
                                    # 数字の桁数で判断
                                    if len(first_number) <= 2:
                                        # 1-2桁なら丁目として扱う
                                        town = general_match.group(1).strip()
                                        block = first_number
                                        number_part = general_match.group(3)
                                        logging.info(f"汎用丁目形式を検出: 町名={town}, 丁目={block}, 番地={number_part}")
                                    else:
                                        # 3桁以上なら番地として扱う
                                        town = general_match.group(1).strip()
                                        block = None
                                        number_part = f"{first_number}-{general_match.group(3)}"
                                        logging.info(f"汎用番地形式を検出: 町名={town}, 番地={number_part}")
                                else:
                                    double_hyphen_match = re.search(r'(\d+)-(\d+)-(\d+)', remaining)
                                    if double_hyphen_match:
                                        block = double_hyphen_match.group(1)
                                        number_part = f"{double_hyphen_match.group(2)}-{double_hyphen_match.group(3)}"
                                        town = remaining[:double_hyphen_match.start()].strip()
                                    else:
                                        number_match = re.search(r'(\d+(?:[-－]\d+)?)', remaining)
                                        if number_match:
                                            number_part = number_match.group(1)
                                            town = remaining[:number_match.start()].strip()
                                        else:
                                            number_part = None
                                            town = remaining
                
                result = {
                    'prefecture': prefecture,
                    'city': city,
                    'town': town if town else "",
                    'block': block,
                    'number': number_part,
                    'building_id': None
                }
                
                # 分割結果をログ出力
                logging.info("住所分割結果:")
                logging.info(f"  都道府県: {result['prefecture']}")
                logging.info(f"  市区町村: {result['city']}")
                logging.info(f"  町名: {result['town']}")
                logging.info(f"  丁目: {result['block']}")
                logging.info(f"  番地: {result['number']}")
                
                return result
            
            return {
                'prefecture': prefecture,
                'city': remaining_address if prefecture else None,
                'town': "",
                'block': None,
                'number': None,
                'building_id': None
            }
        
        return {
            'prefecture': None,
            'city': None,
            'town': "",
            'block': None,
            'number': None,
            'building_id': None
        }
        
    except Exception as e:
        logging.error(f"住所の分割中にエラー: {str(e)}")
        return None

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
    best_match_score = -1  # 詳細なマッチスコア
    
    # 入力住所を分割
    input_parts = split_address(input_address)
    if not input_parts:
        logging.error("入力住所の分割に失敗しました")
        return None, -1
        
    # 基本住所を構築（番地と号を除く）
    base_input_address = f"{input_parts['prefecture']}{input_parts['city']}{input_parts['town']}"
    if input_parts['block']:
        base_input_address += f"{input_parts['block']}丁目"
    
    logging.info(f"基本住所（比較用）: {base_input_address}")
    
    # 入力住所から字名を抽出（原文から直接抽出）
    input_aza_name = None
    original_input_raw = input_address  # 正規化前の原文
    logging.info(f"原文（正規化前）: {original_input_raw}")
    
    if '字' in original_input_raw:
        # 「字」以降、数字（ただし地名として使われる漢数字は除く）が現れるまでの部分を抽出
        # 改良：「丁目」「番」「号」などの前、または半角数字の前で区切る
        aza_pattern = r'字([^0-9０-９\-－−ー番号]+?)(?=\d|番|号|$)'
        logging.info(f"字名抽出パターン: {aza_pattern}")
        aza_match = re.search(aza_pattern, original_input_raw)
        if aza_match:
            input_aza_name = aza_match.group(1).strip()
            logging.info(f"入力住所の字名（正規表現で抽出）: '{input_aza_name}'")
        else:
            logging.info(f"正規表現での字名抽出に失敗。フォールバック処理を実行します。")
            # フォールバック：「字」以降の部分を取得
            aza_index = original_input_raw.find('字')
            if aza_index != -1:
                after_aza = original_input_raw[aza_index + 1:]
                logging.info(f"「字」以降の文字列: '{after_aza}'")
                # 最初の半角数字の位置を探す
                digit_match = re.search(r'\d', after_aza)
                if digit_match:
                    input_aza_name = after_aza[:digit_match.start()].strip()
                    logging.info(f"数字前までの文字列: '{input_aza_name}'")
                else:
                    # 数字が見つからない場合は全体を字名とする
                    input_aza_name = after_aza.strip()
                    logging.info(f"数字なし、全体を字名とする: '{input_aza_name}'")
    else:
        logging.info(f"「字」が含まれていません")
    
    logging.info(f"最終的に抽出された入力住所の字名: '{input_aza_name}'")
    
    # 正規化後の住所（住所比較用）
    original_input = normalize_string(input_address)
    logging.info(f"原文（正規化後）: {original_input}")
    
    if not input_aza_name:
        logging.info("入力住所から字名を抽出できませんでした")
    
    candidate_scores = []  # 各候補のスコアを記録
    
    for candidate in candidates:
        try:
            candidate_text = candidate.text.strip()
            # 候補の住所も分割して基本住所を取得
            candidate_parts = split_address(candidate_text)
            if not candidate_parts:
                continue
                
            base_candidate_address = f"{candidate_parts['prefecture']}{candidate_parts['city']}{candidate_parts['town']}"
            if candidate_parts['block']:
                base_candidate_address += f"{candidate_parts['block']}丁目"
            
            # 基本住所での比較
            _, similarity = is_address_match(base_input_address, base_candidate_address)
            
            logging.info(f"候補 '{candidate_text}' の基本住所: {base_candidate_address}")
            logging.info(f"基本住所での類似度: {similarity}")
            
            # 詳細なマッチスコアを計算
            match_score = similarity
            
            # 字名の完全一致チェック（原文同士で比較）
            aza_match_bonus = 0
            candidate_aza_name = None
            
            if '字' in candidate_text:
                # 候補からも字名を抽出
                logging.info(f"候補原文: '{candidate_text}'")
                aza_pattern = r'字([^0-9０-９\-－−ー番号]+?)(?=\d|番|号|$)'
                candidate_aza_match = re.search(aza_pattern, candidate_text)
                if candidate_aza_match:
                    candidate_aza_name = candidate_aza_match.group(1).strip()
                    logging.info(f"候補の字名（正規表現）: '{candidate_aza_name}'")
                else:
                    # フォールバック：「字」以降の部分を取得
                    aza_index = candidate_text.find('字')
                    if aza_index != -1:
                        after_aza = candidate_text[aza_index + 1:]
                        # 最初の半角数字の位置を探す
                        digit_match = re.search(r'\d', after_aza)
                        if digit_match:
                            candidate_aza_name = after_aza[:digit_match.start()].strip()
                        else:
                            candidate_aza_name = after_aza.strip()
                        logging.info(f"候補の字名（フォールバック）: '{candidate_aza_name}'")
                
                if candidate_aza_name:
                    logging.info(f"字名比較: 入力='{input_aza_name}' vs 候補='{candidate_aza_name}'")
                    
                    if input_aza_name and candidate_aza_name:
                        if input_aza_name == candidate_aza_name:
                            aza_match_bonus = 1.0  # 字名完全一致で決定的なボーナス
                            logging.info(f"★字名完全一致ボーナス: {aza_match_bonus} ('{input_aza_name}' == '{candidate_aza_name}')")
                        elif input_aza_name in candidate_aza_name or candidate_aza_name in input_aza_name:
                            aza_match_bonus = 0.5  # 字名部分一致でボーナス
                            logging.info(f"字名部分一致ボーナス: {aza_match_bonus}")
                        else:
                            logging.info(f"字名不一致: ボーナスなし")
            
            # より厳密な文字列一致チェック
            exact_match_bonus = 0
            normalized_input = original_input  # 既に正規化済み
            normalized_candidate = normalize_string(candidate_text)
            
            # 重要キーワードの完全一致をチェック（字名一致がない場合のみ）
            if aza_match_bonus == 0:
                input_keywords = set(re.findall(r'[一-龯]{2,}', normalized_input))  # 2文字以上の漢字部分
                candidate_keywords = set(re.findall(r'[一-龯]{2,}', normalized_candidate))
                
                # 共通キーワードの重要度を計算
                common_keywords = input_keywords & candidate_keywords
                for keyword in common_keywords:
                    keyword_bonus = len(keyword) * 0.01  # 字名一致がない場合のみ適用
                    exact_match_bonus += keyword_bonus
                    logging.info(f"共通キーワード '{keyword}' によるボーナス: {keyword_bonus}")
            
            # 最長共通部分文字列の長さによるボーナス（字名一致がない場合のみ）
            lcs_bonus = 0
            if aza_match_bonus == 0:
                lcs_length = calculate_longest_common_substring(normalized_input, normalized_candidate)
                if lcs_length >= 4:  # 4文字以上の共通部分文字列
                    lcs_bonus = (lcs_length - 3) * 0.01  # 字名一致がない場合のみ適用
                    logging.info(f"最長共通部分文字列長: {lcs_length}, ボーナス: {lcs_bonus}")
            
            # 候補リストでの順序によるボーナス（微細な差のみ）
            position_bonus = 0
            try:
                candidate_index = candidates.index(candidate)
                # 最初の10件には順序ボーナスを付与（非常に微細な差）
                if candidate_index < 10:
                    position_bonus = (10 - candidate_index) * 0.0001  # より小さな値に調整
                    logging.info(f"順序ボーナス（{candidate_index + 1}位）: {position_bonus}")
            except ValueError:
                pass
            
            # 最終的なマッチスコアを計算
            final_match_score = match_score + aza_match_bonus + exact_match_bonus + lcs_bonus + position_bonus
            
            candidate_scores.append({
                'candidate': candidate,
                'similarity': similarity,
                'match_score': final_match_score,
                'text': candidate_text,
                'aza_bonus': aza_match_bonus,
                'exact_match_bonus': exact_match_bonus,
                'lcs_bonus': lcs_bonus,
                'position_bonus': position_bonus
            })
            
            logging.info(f"候補 '{candidate_text}' の詳細スコア:")
            logging.info(f"  基本類似度: {similarity}")
            logging.info(f"  字名ボーナス: {aza_match_bonus}")
            logging.info(f"  完全一致ボーナス: {exact_match_bonus}")
            logging.info(f"  最長共通部分ボーナス: {lcs_bonus}")
            logging.info(f"  順序ボーナス: {position_bonus}")
            logging.info(f"  最終スコア: {final_match_score}")
            
        except Exception as e:
            logging.warning(f"候補の処理中にエラー: {str(e)}")
            continue
    
    # スコアが最も高い候補を選択
    if candidate_scores:
        # 最終スコアでソート（降順）
        candidate_scores.sort(key=lambda x: x['match_score'], reverse=True)
        
        best_result = candidate_scores[0]
        best_candidate = best_result['candidate']
        best_similarity = best_result['similarity']
        best_match_score = best_result['match_score']
        
        logging.info(f"最適な候補が決定されました:")
        logging.info(f"  候補: {best_result['text']}")
        logging.info(f"  基本類似度: {best_similarity}")
        logging.info(f"  最終スコア: {best_match_score}")
        
        # 上位候補が複数ある場合の詳細ログ
        if len(candidate_scores) > 1:
            score_diff = candidate_scores[0]['match_score'] - candidate_scores[1]['match_score']
            if score_diff < 0.01:  # スコア差が0.01未満の場合
                logging.warning(f"僅差の候補が複数存在します（スコア差: {score_diff}）:")
                for i, score_info in enumerate(candidate_scores[:3]):  # 上位3つまで表示
                    logging.warning(f"  {i+1}位: {score_info['text']} (スコア: {score_info['match_score']})")
    
    if best_candidate and best_similarity >= 0.5:  # 最低限の類似度しきい値
        logging.info(f"最適な候補が見つかりました（類似度: {best_similarity}）: {best_candidate.text.strip()}")
        return best_candidate, best_similarity
    
    return None, best_similarity

def calculate_longest_common_substring(str1, str2):
    """
    2つの文字列の最長共通部分文字列の長さを計算
    
    Args:
        str1 (str): 文字列1
        str2 (str): 文字列2
        
    Returns:
        int: 最長共通部分文字列の長さ
    """
    if not str1 or not str2:
        return 0
    
    max_length = 0
    len1, len2 = len(str1), len(str2)
    
    # 動的プログラミングを使用
    dp = [[0] * (len2 + 1) for _ in range(len1 + 1)]
    
    for i in range(1, len1 + 1):
        for j in range(1, len2 + 1):
            if str1[i-1] == str2[j-1]:
                dp[i][j] = dp[i-1][j-1] + 1
                max_length = max(max_length, dp[i][j])
            else:
                dp[i][j] = 0
    
    return max_length

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
    
    # デバッグログ：入力値の確認
    logging.info(f"=== 検索開始 ===")
    logging.info(f"入力郵便番号（変換前）: {postal_code}")
    logging.info(f"入力住所（変換前）: {address}")
    
    # 郵便番号と住所の正規化
    try:
        postal_code = normalize_address(postal_code)
        address = normalize_address(address)
        logging.info(f"正規化後郵便番号: {postal_code}")
        logging.info(f"正規化後住所: {address}")
    except Exception as e:
        logging.error(f"正規化処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"住所の正規化に失敗しました: {str(e)}"}
    
    # 住所を分割
    if progress_callback:
        progress_callback("住所情報を解析中...")
    
    # キャンセルチェック
    check_cancellation()
    
    address_parts = split_address(address)
    if not address_parts:
        logging.error("住所の分割に失敗しました")
        return {"status": "error", "message": "住所の分割に失敗しました。"}
        
    # 基本住所を構築（番地と号を除く）
    base_address = f"{address_parts['prefecture']}{address_parts['city']}{address_parts['town']}"
    if address_parts['block']:
        base_address += f"{address_parts['block']}丁目"
    
    logging.info(f"住所分割結果 - 基本住所: {base_address}, 番地: {address_parts['number']}")
    
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
        
        # キャンセルチェック
        check_cancellation()
        
        driver.get("https://flets.com/app2/search_c.html")
        logging.info("サイトにアクセスしました")
        

        
        # 郵便番号入力ページが表示されるのを待つ
        # 郵便番号入力フィールドが表示されるまで待機
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "FIELD_ZIP1"))
        )
        logging.info("郵便番号入力ページが表示されました")
        

        
        # 郵便番号入力フィールドを探す
        try:
            # キャンセルチェック
            check_cancellation()
            
            # 郵便番号前半3桁を入力
            postal_code_first_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.NAME, "FIELD_ZIP1"))
            )
            postal_code_first_input.clear()
            postal_code_first_input.send_keys(postal_code_first)
            logging.info(f"郵便番号前半3桁を入力: {postal_code_first}")
            
            # 郵便番号後半4桁を入力
            postal_code_second_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.NAME, "FIELD_ZIP2"))
            )
            postal_code_second_input.clear()
            postal_code_second_input.send_keys(postal_code_second)
            logging.info(f"郵便番号後半4桁を入力: {postal_code_second}")
            
            # 再検索ボタンをクリック
            search_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "cao_submit"))
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
            
            # キャンセルチェック
            check_cancellation()
            
            # 住所候補リストが表示されるまで待機
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "btn_list"))
            )
            logging.info("住所候補リストが表示されました")
            

            
            # 候補リストを取得
            candidates = driver.find_elements(By.CSS_SELECTOR, ".btn_list li.addressInfo")
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
                
                try:
                    # 住所選択前に必要な情報を取得（要素が無効になる前に）
                    selected_address_text = best_candidate.text.strip()
                    logging.info(f"住所選択前に取得した住所テキスト: {selected_address_text}")
                    
                    # JavaScriptを使用してクリックを実行
                    driver.execute_script("arguments[0].click();", best_candidate)
                    logging.info("JavaScriptを使用して住所を選択しました")
                    time.sleep(2)

                    # 住所選択後の丁目・番地調整処理
                    logging.info(f"選択された住所: {selected_address_text}")
                    
                    # 選択された候補の住所構造を分析
                    selected_parts = split_address(selected_address_text)
                    if selected_parts:
                        logging.info(f"選択された住所の分割結果 - 丁目: {selected_parts.get('block')}")
                        
                        # 入力住所では丁目があるが、選択された候補では丁目がない場合
                        if address_parts.get('block') and not selected_parts.get('block'):
                            original_block = address_parts.get('block')
                            original_number = address_parts.get('number', '')
                            
                            logging.info(f"丁目・番地調整が必要です:")
                            logging.info(f"  入力住所の丁目: {original_block}")
                            logging.info(f"  入力住所の番地: {original_number}")
                            logging.info(f"  選択候補の丁目: {selected_parts.get('block')}")
                            
                            # 丁目を番地の先頭に移動
                            if original_number:
                                adjusted_number = f"{original_block}-{original_number}"
                            else:
                                adjusted_number = str(original_block)
                            
                            # address_partsを更新
                            address_parts['block'] = None
                            address_parts['number'] = adjusted_number
                            
                            logging.info(f"丁目・番地を調整しました:")
                            logging.info(f"  調整前 - 丁目: {original_block}, 番地: {original_number}")
                            logging.info(f"  調整後 - 丁目: {address_parts['block']}, 番地: {address_parts['number']}")
                            
                            # 調整後の値を確認
                            logging.info(f"調整後のaddress_parts全体: {address_parts}")
                        else:
                            logging.info("丁目・番地の調整は不要です")
                            logging.info(f"  入力住所の丁目: {address_parts.get('block')}")
                            logging.info(f"  選択候補の丁目: {selected_parts.get('block')}")
                    else:
                        logging.warning("選択された住所の分割に失敗しました")

                    # 住所選択直後に「要調査」文言や特殊URLを即時チェック
                    page_title = driver.title
                    current_url = driver.current_url
                    page_text = driver.find_element(By.TAG_NAME, "body").text
                    if ("詳しい状況確認が必要" in page_text or
                        "InfoSpecialAddressCollabo" in current_url):
                        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                        screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        logging.info("要調査文言またはInfoSpecialAddressCollabo遷移を検出。即時要調査判定で返却")
                        return {
                            "status": "investigation",
                            "message": "要調査",
                            "details": {
                                "判定結果": "要調査",
                                "提供エリア": "詳しい状況確認が必要です",
                                "備考": "ご指定の住所は『光アクセスサービス』の詳しい状況確認が必要です。"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    
                    # 番地入力画面への遷移を待機
                    WebDriverWait(driver, 15).until(
                        EC.url_contains("/cao/InputAddressNum")
                    )
                    logging.info("番地入力画面への遷移を確認しました")
                    
                    # ページの読み込み完了を待機
                    WebDriverWait(driver, 20).until(
                        lambda d: d.execute_script('return document.readyState') == 'complete'
                    )
                    logging.info("番地入力ページの読み込みが完了しました")
                    
                    if progress_callback:
                        progress_callback("番地を入力中...")
                    
                    # 番地入力画面の処理に進む
                    result = handle_address_number_input(driver, address_parts, progress_callback)
                    logging.info(f"番地入力処理の結果: {result}")
                    return result
                    
                except Exception as e:
                    logging.error(f"住所選択処理に失敗: {str(e)}")
                    take_full_page_screenshot(driver, "debug_address_select_error.png")
                    raise
            else:
                logging.error(f"適切な住所候補が見つかりませんでした。入力住所: {address}")
                raise ValueError("適切な住所候補が見つかりませんでした")
            
        except Exception as e:
            logging.error(f"住所選択処理中にエラー: {str(e)}")
            return {"status": "error", "message": f"住所選択処理中にエラーが発生しました: {str(e)}"}
            
    except CancellationError as e:
        logging.info("★★★ 提供エリア検索がキャンセルされました ★★★")
        return {"status": "cancelled", "message": "検索がキャンセルされました"}
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

def find_input_element(driver, attempt_count=0):
    """
    番地入力フォームを見つけるためのヘルパー関数
    
    Args:
        driver: WebDriverインスタンス
        attempt_count: 試行回数
        
    Returns:
        element: 見つかった要素、見つからない場合はNone
    """
    try:
        # iframeの確認
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            for iframe in iframes:
                try:
                    driver.switch_to.frame(iframe)
                    element = driver.find_element(By.NAME, "banchi1to3manualAddressNum1")
                    if element.is_displayed():
                        return element
                except:
                    pass
                finally:
                    driver.switch_to.default_content()
        
        # 複数の方法で要素を探す
        selectors = [
            (By.NAME, "banchi1to3manualAddressNum1"),
            (By.ID, "id_banchi1to3manualAddressNum1"),
            (By.CSS_SELECTOR, "input[name='banchi1to3manualAddressNum1']"),
            (By.CSS_SELECTOR, "input[type='tel'][name='banchi1to3manualAddressNum1']"),
            (By.XPATH, "//input[@name='banchi1to3manualAddressNum1']"),
            (By.XPATH, "//div[contains(@class, '_input')]//input[1]")
        ]
        
        for by, selector in selectors:
            try:
                element = driver.find_element(by, selector)
                if element.is_displayed():
                    return element
            except:
                continue
        
        return None
    except Exception as e:
        logging.warning(f"要素検索中にエラー: {str(e)}")
        return None

def debug_page_state(driver, context=""):
    """
    ページの状態をデバッグ出力する
    
    Args:
        driver: WebDriverインスタンス
        context: デバッグコンテキストの説明
    """
    try:
        logging.info(f"=== デバッグ情報（{context}）===")
        logging.info(f"現在のURL: {driver.current_url}")
        logging.info(f"ページタイトル: {driver.title}")
        
        # ページの読み込み状態
        ready_state = driver.execute_script('return document.readyState')
        logging.info(f"ページの読み込み状態: {ready_state}")
        
        # DOMの準備状態
        is_dom_loaded = driver.execute_script('return document.readyState === "complete" || document.readyState === "interactive"')
        logging.info(f"DOMの準備完了: {is_dom_loaded}")
        
        # iframeの確認
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        logging.info(f"iframe数: {len(iframes)}")
        
        # エラーメッセージの確認
        error_elements = driver.find_elements(By.CLASS_NAME, "error")
        if error_elements:
            logging.info("エラーメッセージが見つかりました:")
            for error in error_elements:
                logging.info(f"エラー: {error.text}")
        
    except Exception as e:
        logging.error(f"デバッグ情報の取得中にエラー: {str(e)}")

def handle_address_number_input(driver, address_parts, progress_callback=None):
    """
    番地入力画面の処理を行う
    
    Args:
        driver: WebDriverインスタンス
        address_parts: 分割された住所情報
        progress_callback: 進捗コールバック関数
    """
    try:
        # ブラウザ設定を読み込む（show_popup用）
        show_popup = True  # デフォルト値
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    browser_settings = settings.get("browser_settings", {})
                    show_popup = browser_settings.get("show_popup", True)
                    logging.info(f"ポップアップ表示設定を読み込みました: {show_popup}")
        except Exception as e:
            logging.warning(f"ブラウザ設定の読み込みに失敗しました: {str(e)}")

        logging.info("=== 番地入力画面の処理開始 ===")
        debug_page_state(driver, "番地入力画面_初期状態")

        # ページ読み込み完了まで待機
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "id_form_main"))
        )
        debug_page_state(driver, "番地入力画面_読み込み完了後")
        
        # address_partsの内容を確認
        logging.info(f"番地入力処理開始時のaddress_parts:")
        logging.info(f"  都道府県: {address_parts.get('prefecture')}")
        logging.info(f"  市区町村: {address_parts.get('city')}")
        logging.info(f"  町名: {address_parts.get('town')}")
        logging.info(f"  丁目: {address_parts.get('block')}")
        logging.info(f"  番地: {address_parts.get('number')}")


        # 番地がない場合のチェックボックス処理
        if not address_parts.get('number'):
            logging.info("番地が指定されていないため、「番地・号が無い」をチェック")
            try:
                checkbox = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "id_banchi1to3Fixed"))
                )
                if not checkbox.is_selected():
                    driver.execute_script("arguments[0].click();", checkbox)
                logging.info("「番地・号が無い」をチェックしました")
                

            except Exception as e:
                logging.error(f"「番地・号が無い」チェックボックスの操作に失敗: {str(e)}")
                debug_page_state(driver, "チェックボックス操作_失敗")
                raise
        else:
            # 番地入力フィールドの処理
            number_parts = address_parts['number'].split('-')
            logging.info(f"入力する番地: {number_parts}")

            try:
                # 番地1の入力
                number1_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "id_banchi1to3manualAddressNum1"))
                )
                # フォーカスを設定してからクリア
                driver.execute_script("arguments[0].focus();", number1_input)
                number1_input.clear()
                number1_input.send_keys(number_parts[0])
                logging.info(f"番地1を入力: {number_parts[0]}")

                # 番地2の入力（存在する場合）
                if len(number_parts) > 1:
                    number2_input = driver.find_element(By.ID, "id_banchi1to3manualAddressNum2")
                    driver.execute_script("arguments[0].focus();", number2_input)
                    number2_input.clear()
                    number2_input.send_keys(number_parts[1])
                    logging.info(f"番地2を入力: {number_parts[1]}")

                # 番地3の入力（存在する場合）
                if len(number_parts) > 2:
                    number3_input = driver.find_element(By.ID, "id_banchi1to3manualAddressNum3")
                    driver.execute_script("arguments[0].focus();", number3_input)
                    number3_input.clear()
                    number3_input.send_keys(number_parts[2])
                    logging.info(f"番地3を入力: {number_parts[2]}")
                
      

            except Exception as e:
                logging.error(f"番地入力に失敗: {str(e)}")
                debug_page_state(driver, "番地入力_失敗")
                raise

        # 住居タイプの選択（デフォルトで戸建てを選択）
        try:
            house_type = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id_buildType_1"))
            )
            if not house_type.is_selected():
                driver.execute_script("arguments[0].click();", house_type)
            logging.info("住居タイプ: 戸建てを選択")
            


            # 次へボタンをクリック
            next_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "id_nextButton"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(1)  # スクロール完了を待つ
            driver.execute_script("arguments[0].click();", next_button)
            logging.info("次へボタンをクリックしました")

            # 分岐1: 該当する番地・号が見つからない場合のモーダル処理
            try:
                # モーダルメッセージの確認（短い待機時間）
                modal_message = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "p.ico_att.jc_c"))
                )
                if "該当する番地・号が見つかりませんでした" in modal_message.text:
                    logging.info("番地・号未発見モーダルを検出しました")
                    
                    # モーダルの「次へ」ボタンをクリック
                    modal_next_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "id_modalNextButton"))
                    )
                    driver.execute_script("arguments[0].click();", modal_next_button)
                    logging.info("モーダルの次へボタンをクリックしました")
                    
                    # SelectAddressNum1ページへの遷移を待機
                    WebDriverWait(driver, 10).until(
                        EC.url_contains("SelectAddressNum1")
                    )
                    logging.info("番地選択画面(SelectAddressNum1)へ遷移しました")
                    
                    # 番地選択処理
                    return handle_address_number_selection(driver, address_parts, progress_callback, show_popup)
                    
            except TimeoutException:
                logging.info("番地・号未発見モーダルは表示されませんでした")
                # モーダルが表示されない場合は次の処理に進む
                pass
            except Exception as e:
                logging.warning(f"モーダル処理中にエラー: {str(e)}")
                # エラーが発生した場合も次の処理に進む
                pass

            # 分岐2: 建物選択画面が表示されたかチェック
            try:
                # 建物選択画面のURLを確認
                WebDriverWait(driver, 10).until(
                    lambda d: "SelectBuild1" in d.current_url
                )
                logging.info("建物選択画面が表示されました")

                # 建物選択画面が表示された時点で集合住宅と判定
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = f"debug_apartment_confirmation_{timestamp}.png"
                take_full_page_screenshot(driver, screenshot_path)
                return {
                    "status": "apartment",
                    "message": "集合住宅",
                    "details": {
                        "判定結果": "NG",
                        "提供エリア": "集合住宅",
                        "備考": "集合住宅のため、判定を終了します"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }

            except TimeoutException:
                # 建物選択画面が表示されない場合は通常の結果ページへの遷移を待機
                logging.info("建物選択画面はスキップされました")
                WebDriverWait(driver, 10).until(
                    EC.url_contains("ProvideResult")
                )
                logging.info("結果ページへ遷移しました")

        except Exception as e:
            logging.error(f"住居タイプの選択または次へボタンのクリックに失敗: {str(e)}")
            debug_page_state(driver, "住居タイプ選択_次へボタン_失敗")
            raise

        # 結果ページへの遷移を待機
        try:
            WebDriverWait(driver, 10).until(
                EC.url_contains("ProvideResult")
            )
            logging.info("結果ページへ遷移しました")
            
            
            debug_page_state(driver, "結果ページ_表示")

            # 結果テキストの取得を修正
            try:
                # まず、ローディング表示が消えるのを待つ
                WebDriverWait(driver, 10).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "loading"))
                )
                
                # 結果テキストを取得（複数の方法で試行）
                result_text = None
                
                # 結果テキストの取得方法を追加
                selectors = [
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//h1/following-sibling::div"),
                    (By.CLASS_NAME, "resultText"),
                    (By.XPATH, "//div[contains(text(), 'フレッツ光') and contains(text(), 'エリア')]"),
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[contains(text(), 'エリア')]"),
                    (By.XPATH, "//h1[contains(text(), 'エリア')]"),
                    (By.XPATH, "//div[contains(@class, 'result')]"),
                    (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[not(@class)]")
                ]

                for selector_type, selector in selectors:
                    try:
                        element = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((selector_type, selector))
                        )
                        if element and element.is_displayed():
                            result_text = element.text.strip()
                            logging.info(f"結果テキストを取得: {result_text} (セレクター: {selector})")
                            if result_text:
                                break
                    except Exception as e:
                        logging.debug(f"セレクター {selector} での検索に失敗: {str(e)}")
                        continue

                # 結果が見つからない場合、ページ全体のテキストから判定
                if not result_text:
                    logging.info("個別の要素での結果テキスト取得に失敗。ページ全体から検索を試みます。")
                    page_text = driver.find_element(By.TAG_NAME, "body").text
                    if "提供エリアです" in page_text:
                        result_text = "提供エリアです"
                        logging.info("ページテキストから「提供エリアです」を検出")
                    elif "提供エリア外です" in page_text:
                        result_text = "提供エリア外です"
                        logging.info("ページテキストから「提供エリア外です」を検出")
                    elif "詳しい状況確認が必要" in page_text or "詳しい状況確認が必要" in (result_text or ""):
                        result_text = "要調査"
                        logging.info("ページテキストから「詳しい状況確認が必要」を検出")

                logging.info(f"最終的な結果テキスト: {result_text}")

                # スクリーンショットを保存（結果確認時のみ）
                timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                
                if result_text:
                    if "提供エリアです" in result_text or "の提供エリアです" in result_text:
                        screenshot_path = f"debug_available_confirmation_{timestamp}.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        return {
                            "status": "available",
                            "message": "提供可能",
                            "details": {
                                "判定結果": "OK",
                                "提供エリア": "提供可能エリアです",
                                "備考": "フレッツ光のサービスがご利用いただけます"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    elif "提供エリア外です" in result_text or "エリア外" in result_text:
                        screenshot_path = f"debug_not_provided_confirmation_{timestamp}.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        return {
                            "status": "unavailable",
                            "message": "未提供",
                            "details": {
                                "判定結果": "NG",
                                "提供エリア": "提供対象外エリアです",
                                "備考": "申し訳ございませんが、このエリアではサービスを提供しておりません"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    elif "要調査" in result_text or "詳しい状況確認が必要" in result_text:
                        screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        return {
                            "status": "investigation",
                            "message": "要調査",
                            "details": {
                                "判定結果": "要調査",
                                "提供エリア": "詳しい状況確認が必要です",
                                "備考": "ご指定の住所は『光アクセスサービス』の詳しい状況確認が必要です。"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                    else:
                        screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                        take_full_page_screenshot(driver, screenshot_path)
                        logging.warning(f"予期しない結果テキスト: {result_text}")
                        return {
                            "status": "failure",
                            "message": "判定失敗",
                            "details": {
                                "判定結果": "判定失敗",
                                "提供エリア": "調査が必要なエリアです",
                                "備考": "住所を特定できないため、担当者がお調べします"
                            },
                            "screenshot": screenshot_path,
                            "show_popup": show_popup
                        }
                else:

                    
                    screenshot_path = f"debug_error_confirmation_{timestamp}.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    logging.error("結果テキストが取得できませんでした")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "判定できませんでした",
                            "備考": "結果テキストが取得できませんでした"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }

            except Exception as e:
                screenshot_path = f"debug_error_confirmation_{timestamp}.png"
                take_full_page_screenshot(driver, screenshot_path)
                logging.error(f"結果テキストの取得中にエラー: {str(e)}")
                return {
                    "status": "failure",
                    "message": "判定失敗",
                    "details": {
                        "判定結果": "判定失敗",
                        "提供エリア": "判定できませんでした",
                        "備考": f"結果の判定に失敗しました: {str(e)}"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }

        except Exception as e:
            logging.error(f"結果の取得に失敗: {str(e)}")
            debug_page_state(driver, "結果取得_失敗")
            return {"status": "error", "message": "結果の取得に失敗しました"}

    except Exception as e:
        logging.error(f"番地入力画面の処理中にエラー: {str(e)}")
        debug_page_state(driver, "エラー発生時の状態")
        raise 

def handle_address_number_selection(driver, address_parts, progress_callback=None, show_popup=True):
    """
    番地選択画面（SelectAddressNum1/SelectAddressNum2）の処理を行う
    
    Args:
        driver: WebDriverインスタンス
        address_parts: 分割された住所情報
        progress_callback: 進捗コールバック関数
        show_popup: ポップアップ表示設定
        
    Returns:
        dict: 処理結果
    """
    try:
        logging.info("=== 番地選択画面の処理開始 ===")
        debug_page_state(driver, "番地選択画面_初期状態")
        
        # 進捗更新
        if progress_callback:
            progress_callback("番地を選択中...")
        
        # 番地情報を取得
        target_number = address_parts.get('number', '').split('-')[0] if address_parts.get('number') else None
        logging.info(f"選択対象の番地: {target_number}")
        
        if not target_number:
            logging.error("選択する番地が見つかりません")
            return {"status": "error", "message": "選択する番地が見つかりません"}
        
        # SelectAddressNum1での番地選択
        try:
            # 番地リストの読み込み完了を待機
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "li.addressNum1"))
            )
            logging.info("番地リストが表示されました")
            
            # 番地候補を取得
            address_candidates = driver.find_elements(By.CSS_SELECTOR, "li.addressNum1")
            logging.info(f"{len(address_candidates)} 件の番地候補が見つかりました")
            
            # 最適な番地を選択
            selected = False
            max_attempts = 3
            
            for attempt in range(max_attempts):
                try:
                    # 番地候補を取得（毎回新しく取得してStaleElementExceptionを回避）
                    address_candidates = driver.find_elements(By.CSS_SELECTOR, "li.addressNum1")
                    logging.info(f"{len(address_candidates)} 件の番地候補が見つかりました")
                    
                    for i, candidate in enumerate(address_candidates):
                        try:
                            candidate_text = candidate.get_attribute("data-addressnum1")
                            if not candidate_text:
                                candidate_text = candidate.text.strip()
                            
                            # 全角数字を半角数字に変換してから数字のみを抽出
                            zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
                            candidate_normalized = candidate_text.translate(zen_to_han)
                            candidate_number = re.sub(r'[^\d]', '', candidate_normalized)
                            target_number_clean = re.sub(r'[^\d]', '', target_number)
                            
                            logging.info(f"候補番地: {candidate_text} → 正規化: {candidate_normalized} → 数字: {candidate_number}")
                            
                            if candidate_number == target_number_clean:
                                # 番地ボタンをクリック
                                button = candidate.find_element(By.TAG_NAME, "button")
                                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                                time.sleep(0.5)  # スクロール完了を待つ
                                driver.execute_script("arguments[0].click();", button)
                                logging.info(f"番地を選択しました: {candidate_text}")
                                selected = True
                                break
                        except Exception as e:
                            logging.warning(f"候補 {i} の処理中にエラー: {str(e)}")
                            continue
                    
                    if selected:
                        break
                        
                except Exception as e:
                    logging.warning(f"番地選択の試行 {attempt + 1} でエラー: {str(e)}")
                    if attempt < max_attempts - 1:
                        time.sleep(1)  # 1秒待機してリトライ
                        continue
                    else:
                        raise
            
            if not selected:
                # 完全一致しない場合は最初の候補を選択
                logging.warning(f"完全一致する番地が見つからないため、最初の候補を選択します")
                try:
                    address_candidates = driver.find_elements(By.CSS_SELECTOR, "li.addressNum1")
                    if address_candidates:
                        first_candidate = address_candidates[0]
                        button = first_candidate.find_element(By.TAG_NAME, "button")
                        driver.execute_script("arguments[0].scrollIntoView(true);", button)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", button)
                        logging.info(f"最初の候補を選択しました: {first_candidate.text}")
                        selected = True
                except Exception as e:
                    logging.error(f"最初の候補選択でもエラー: {str(e)}")
                    raise
            
            # 次の画面への遷移を待機
            time.sleep(2)
            current_url = driver.current_url
            logging.info(f"番地選択後のURL: {current_url}")
            
            # 遷移先の判定
            if "SelectBuild1" in current_url:
                # 建物選択画面に遷移 → 集合住宅判定
                logging.info("建物選択画面に遷移しました")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_path = f"debug_apartment_confirmation_{timestamp}.png"
                take_full_page_screenshot(driver, screenshot_path)
                return {
                    "status": "apartment",
                    "message": "集合住宅",
                    "details": {
                        "判定結果": "NG",
                        "提供エリア": "集合住宅",
                        "備考": "集合住宅のため、判定を終了します"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }
                
            elif "SelectAddressNum2" in current_url:
                # 号選択画面に遷移
                logging.info("号選択画面(SelectAddressNum2)に遷移しました")
                return handle_go_selection(driver, address_parts, progress_callback, show_popup)
                
            elif "ProvideResult" in current_url:
                # 結果画面に直接遷移
                logging.info("結果画面に直接遷移しました")
                return handle_result_page(driver, show_popup)
                
            else:
                # 予期しない遷移
                logging.warning(f"予期しない遷移先: {current_url}")
                # 10秒待機して再度確認
                time.sleep(10)
                current_url = driver.current_url
                
                if "ProvideResult" in current_url:
                    return handle_result_page(driver, show_popup)
                else:
                    return {"status": "error", "message": "予期しない画面に遷移しました"}
                    
        except Exception as e:
            logging.error(f"番地選択処理中にエラー: {str(e)}")
            debug_page_state(driver, "番地選択_エラー")
            raise
            
    except Exception as e:
        logging.error(f"番地選択画面の処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"番地選択処理中にエラーが発生しました: {str(e)}"}

def handle_go_selection(driver, address_parts, progress_callback=None, show_popup=True):
    """
    号選択画面（SelectAddressNum2）の処理を行う
    
    Args:
        driver: WebDriverインスタンス
        address_parts: 分割された住所情報
        progress_callback: 進捗コールバック関数
        show_popup: ポップアップ表示設定
        
    Returns:
        dict: 処理結果
    """
    try:
        logging.info("=== 号選択画面の処理開始 ===")
        debug_page_state(driver, "号選択画面_初期状態")
        
        # 進捗更新
        if progress_callback:
            progress_callback("号を選択中...")
        
        # 号情報を取得（番地の2番目の部分）
        number_parts = address_parts.get('number', '').split('-')
        target_go = number_parts[1] if len(number_parts) > 1 else "1"
        logging.info(f"選択対象の号: {target_go}")
        
        # 号リストの読み込み完了を待機
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.addressNum2"))
        )
        logging.info("号リストが表示されました")
        
        # 号候補を取得
        go_candidates = driver.find_elements(By.CSS_SELECTOR, "li.addressNum2")
        logging.info(f"{len(go_candidates)} 件の号候補が見つかりました")
        
        # 最適な号を選択
        selected = False
        for candidate in go_candidates:
            candidate_text = candidate.get_attribute("data-addressnum2")
            if not candidate_text:
                candidate_text = candidate.text.strip()
            
            # 全角数字を半角数字に変換してから数字のみを抽出
            zen_to_han = str.maketrans('０１２３４５６７８９', '0123456789')
            candidate_normalized = candidate_text.translate(zen_to_han)
            candidate_go = re.sub(r'[^\d]', '', candidate_normalized)
            target_go_clean = re.sub(r'[^\d]', '', target_go)
            
            logging.info(f"候補号: {candidate_text} → 正規化: {candidate_normalized} → 数字: {candidate_go}")
            
            if candidate_go == target_go_clean:
                # 号ボタンをクリック
                button = candidate.find_element(By.TAG_NAME, "button")
                driver.execute_script("arguments[0].scrollIntoView(true);", button)
                time.sleep(0.5)  # スクロール完了を待つ
                driver.execute_script("arguments[0].click();", button)
                logging.info(f"号を選択しました: {candidate_text}")
                selected = True
                break
        
        if not selected:
            # 完全一致しない場合は最初の候補を選択
            logging.warning(f"完全一致する号が見つからないため、最初の候補を選択します")
            first_candidate = go_candidates[0]
            button = first_candidate.find_element(By.TAG_NAME, "button")
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", button)
            logging.info(f"最初の候補を選択しました: {first_candidate.text}")
        
        # 号選択後の遷移先を判定
        time.sleep(2)
        current_url = driver.current_url
        logging.info(f"号選択後のURL: {current_url}")
        
        # 建物選択画面に遷移した場合
        try:
            WebDriverWait(driver, 5).until(
                lambda d: "SelectBuild1" in d.current_url
            )
            logging.info("号選択後に建物選択画面に遷移しました")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = f"debug_apartment_confirmation_{timestamp}.png"
            take_full_page_screenshot(driver, screenshot_path)
            return {
                "status": "apartment",
                "message": "集合住宅",
                "details": {
                    "判定結果": "NG",
                    "提供エリア": "集合住宅",
                    "備考": "集合住宅のため、判定を終了します"
                },
                "screenshot": screenshot_path,
                "show_popup": show_popup
            }
        except TimeoutException:
            # 建物選択画面に遷移しない場合は結果画面への遷移を待機
            logging.info("号選択後に建物選択画面は表示されませんでした")
            pass
        
        # 結果画面への遷移を待機
        try:
            WebDriverWait(driver, 10).until(
                EC.url_contains("ProvideResult")
            )
            logging.info("結果画面へ遷移しました")
            
            # 結果画面の処理
            return handle_result_page(driver, show_popup)
        except TimeoutException:
            logging.error("結果画面への遷移がタイムアウトしました")
            current_url = driver.current_url
            logging.error(f"現在のURL: {current_url}")
            return {"status": "error", "message": "結果画面への遷移に失敗しました"}
        
    except Exception as e:
        logging.error(f"号選択画面の処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"号選択処理中にエラーが発生しました: {str(e)}"}

def handle_result_page(driver, show_popup=True):
    """
    結果画面の処理を行う
    
    Args:
        driver: WebDriverインスタンス
        show_popup: ポップアップ表示設定
        
    Returns:
        dict: 処理結果
    """
    try:
        logging.info("=== 結果画面の処理開始 ===")
        debug_page_state(driver, "結果画面_表示")

        # 結果テキストの取得
        try:
            # まず、ローディング表示が消えるのを待つ
            WebDriverWait(driver, 10).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "loading"))
            )
            
            # 結果テキストを取得（複数の方法で試行）
            result_text = None
            
            # 結果テキストの取得方法
            selectors = [
                (By.XPATH, "//div[contains(@class, 'main_wrap')]//h1/following-sibling::div"),
                (By.CLASS_NAME, "resultText"),
                (By.XPATH, "//div[contains(text(), 'フレッツ光') and contains(text(), 'エリア')]"),
                (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[contains(text(), 'エリア')]"),
                (By.XPATH, "//h1[contains(text(), 'エリア')]"),
                (By.XPATH, "//div[contains(@class, 'result')]"),
                (By.XPATH, "//div[contains(@class, 'main_wrap')]//div[not(@class)]")
            ]

            for selector_type, selector in selectors:
                try:
                    element = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((selector_type, selector))
                    )
                    if element and element.is_displayed():
                        result_text = element.text.strip()
                        logging.info(f"結果テキストを取得: {result_text} (セレクター: {selector})")
                        if result_text:
                            break
                except Exception as e:
                    logging.debug(f"セレクター {selector} での検索に失敗: {str(e)}")
                    continue

            # 結果が見つからない場合、ページ全体のテキストから判定
            if not result_text:
                logging.info("個別の要素での結果テキスト取得に失敗。ページ全体から検索を試みます。")
                page_text = driver.find_element(By.TAG_NAME, "body").text
                if "提供エリアです" in page_text:
                    result_text = "提供エリアです"
                    logging.info("ページテキストから「提供エリアです」を検出")
                elif "提供エリア外です" in page_text:
                    result_text = "提供エリア外です"
                    logging.info("ページテキストから「提供エリア外です」を検出")
                elif "詳しい状況確認が必要" in page_text:
                    result_text = "要調査"
                    logging.info("ページテキストから「詳しい状況確認が必要」を検出")

            logging.info(f"最終的な結果テキスト: {result_text}")

            # スクリーンショットを保存
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            
            if result_text:
                if "提供エリアです" in result_text or "の提供エリアです" in result_text:
                    screenshot_path = f"debug_available_confirmation_{timestamp}.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    return {
                        "status": "available",
                        "message": "提供可能",
                        "details": {
                            "判定結果": "OK",
                            "提供エリア": "提供可能エリアです",
                            "備考": "フレッツ光のサービスがご利用いただけます"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }
                elif "提供エリア外です" in result_text or "エリア外" in result_text:
                    screenshot_path = f"debug_not_provided_confirmation_{timestamp}.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    return {
                        "status": "unavailable",
                        "message": "未提供",
                        "details": {
                            "判定結果": "NG",
                            "提供エリア": "提供対象外エリアです",
                            "備考": "申し訳ございませんが、このエリアではサービスを提供しておりません"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }
                elif "要調査" in result_text or "詳しい状況確認が必要" in result_text:
                    screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    return {
                        "status": "investigation",
                        "message": "要調査",
                        "details": {
                            "判定結果": "要調査",
                            "提供エリア": "詳しい状況確認が必要です",
                            "備考": "ご指定の住所は『光アクセスサービス』の詳しい状況確認が必要です。"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }
                else:
                    screenshot_path = f"debug_investigation_confirmation_{timestamp}.png"
                    take_full_page_screenshot(driver, screenshot_path)
                    logging.warning(f"予期しない結果テキスト: {result_text}")
                    return {
                        "status": "failure",
                        "message": "判定失敗",
                        "details": {
                            "判定結果": "判定失敗",
                            "提供エリア": "調査が必要なエリアです",
                            "備考": "住所を特定できないため、担当者がお調べします"
                        },
                        "screenshot": screenshot_path,
                        "show_popup": show_popup
                    }
            else:
                screenshot_path = f"debug_error_confirmation_{timestamp}.png"
                take_full_page_screenshot(driver, screenshot_path)
                logging.error("結果テキストが取得できませんでした")
                return {
                    "status": "failure",
                    "message": "判定失敗",
                    "details": {
                        "判定結果": "判定失敗",
                        "提供エリア": "判定できませんでした",
                        "備考": "結果テキストが取得できませんでした"
                    },
                    "screenshot": screenshot_path,
                    "show_popup": show_popup
                }

        except Exception as e:
            timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            screenshot_path = f"debug_error_confirmation_{timestamp}.png"
            take_full_page_screenshot(driver, screenshot_path)
            logging.error(f"結果テキストの取得中にエラー: {str(e)}")
            return {
                "status": "failure",
                "message": "判定失敗",
                "details": {
                    "判定結果": "判定失敗",
                    "提供エリア": "判定できませんでした",
                    "備考": f"結果の判定に失敗しました: {str(e)}"
                },
                "screenshot": screenshot_path,
                "show_popup": show_popup
            }

    except Exception as e:
        logging.error(f"結果画面の処理中にエラー: {str(e)}")
        return {"status": "error", "message": f"結果画面の処理中にエラーが発生しました: {str(e)}"} 