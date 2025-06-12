"""
顧客情報取得サービス

CTIメインウィンドウから情報を取得し、
アプリケーションに反映する機能を提供します。
"""

import win32gui
import win32con
import win32api
import logging
from dataclasses import dataclass
from typing import Optional, Dict
import ctypes
import re
import os
import time
from enum import Enum
from datetime import datetime, timedelta

# ログレベルをINFOに設定
logging.getLogger().setLevel(logging.DEBUG)

class CTIStatus(Enum):
    """CTI状態の列挙型"""
    WAITING = "待ち受け中"
    DIALING = "発信中"
    TALKING = "通話中"
    UNKNOWN = "不明"

@dataclass
class CTIData:
    """CTIメインから取得するデータ構造"""
    customer_name: str = ""    # 顧客・漢字
    address: str = ""          # 住所
    phone: str = ""           # 優先電話番号1
    postal_code: str = ""      # 〒
    management_id: str = ""    # 管理番号
    list_name: str = ""        # リスト

class OneClickService:
    def __init__(self):
        self.window_handle = None
        self.field_info = {}
        # ラベルと実際のフィールド名のマッピング
        self.label_mappings = {
            "顧客・漢字": "customer_name",
            "住所": "address",
            "優先電話番号1": "phone",
            "〒": "postal_code",
            "管理番号": "management_id",
            "リスト": "list_name"
        }
        # CTI状態のラベルマッピング
        self.status_labels = {
            CTIStatus.WAITING.value: "waiting",
            CTIStatus.DIALING.value: "dialing",
            CTIStatus.TALKING.value: "talking"
        }
        # 前回のCTI状態を保持
        self._previous_status = ""
        # ログ出力の制御用
        self.last_log_time = datetime.now()
        self.log_interval = timedelta(minutes=1)  # ログ出力の間隔（1分）
        logging.debug("OneClickService initialized")

    def find_cti_window(self) -> bool:
        """
        CTIメインウィンドウを検索する
        
        Returns:
            bool: ウィンドウが見つかった場合True
        """
        try:
            def callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if "CTIメイン" in title:
                        logging.info(f"CTIメインウィンドウを検出: handle={hwnd}, title='{title}'")
                        self.window_handle = hwnd
                        return False
                return True
            
            win32gui.EnumWindows(callback, None)
            
            if self.window_handle:
                # ウィンドウが最小化されている場合は元に戻す
                if win32gui.IsIconic(self.window_handle):
                    win32gui.ShowWindow(self.window_handle, win32con.SW_RESTORE)
                
                # ウィンドウを前面に表示
                win32gui.SetForegroundWindow(self.window_handle)
                
                # 子ウィンドウを列挙してフィールド情報を収集
                def enum_child_windows(hwnd, param):
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            class_name = win32gui.GetClassName(hwnd)
                            if class_name.startswith("WindowsForms10."):
                                text = win32gui.GetWindowText(hwnd)
                                if text:
                                    self.field_info[text] = hwnd
                                    logging.debug(f"フィールド情報を追加: text='{text}', handle={hwnd}")
                    except Exception as e:
                        logging.warning(f"子ウィンドウの処理中にエラー: handle={hwnd}, error={str(e)}")
                    return True
                
                win32gui.EnumChildWindows(self.window_handle, enum_child_windows, None)
                logging.info(f"フィールド情報を {len(self.field_info)} 個収集")
                return True
            else:
                # 前回のログ出力から指定時間が経過している場合のみログを出力
                now = datetime.now()
                if now - self.last_log_time >= self.log_interval:
                    logging.warning("CTIメインウィンドウが見つかりません")
                    self.last_log_time = now
                return False
                
        except Exception as e:
            logging.error(f"CTIメインウィンドウの検索中にエラー: {str(e)}")
            return False

    def find_all_controls(self) -> list:
        """
        すべてのコントロールを検索
        
        Returns:
            list: コントロール情報のリスト
        """
        controls = []
        
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name.startswith("WindowsForms10."):
                        rect = win32gui.GetWindowRect(hwnd)
                        client_rect = list(rect)
                        client_rect[0], client_rect[1] = win32gui.ScreenToClient(self.window_handle, (rect[0], rect[1]))
                        client_rect[2], client_rect[3] = win32gui.ScreenToClient(self.window_handle, (rect[2], rect[3]))
                        
                        control = {
                            'hwnd': hwnd,
                            'class': class_name,
                            'text': self.get_control_text(hwnd),
                            'rect': rect,
                            'client_rect': client_rect
                        }
                        controls.append(control)
                except Exception as e:
                    logging.warning(f"コントロール情報の取得中にエラー: handle={hwnd}, error={str(e)}")
            return True
        
        if self.window_handle:
            win32gui.EnumChildWindows(self.window_handle, callback, None)
        
        return controls

    def find_edit_controls(self):
        """すべての編集可能なテキストボックスを取得"""
        edit_controls = []
        
        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name.startswith("WindowsForms10."):
                        # スタイルを取得
                        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                        
                        # 編集可能なテキストボックスの条件をチェック
                        if ("EDIT" in class_name or "TextBox" in class_name) and not (style & win32con.ES_READONLY):
                            rect = win32gui.GetWindowRect(hwnd)
                            client_rect = list(rect)
                            client_rect[0], client_rect[1] = win32gui.ScreenToClient(self.window_handle, (rect[0], rect[1]))
                            client_rect[2], client_rect[3] = win32gui.ScreenToClient(self.window_handle, (rect[2], rect[3]))
                            
                            # サイズを計算
                            size = (rect[2] - rect[0], rect[3] - rect[1])
                            
                            control = {
                                'hwnd': hwnd,
                                'class': class_name,
                                'text': self.get_control_text(hwnd),
                                'rect': rect,
                                'client_rect': client_rect,
                                'size': size
                            }
                            edit_controls.append(control)
                            logging.debug(f"編集可能なコントロールを検出: handle={hwnd}, "
                                        f"class='{class_name}', text='{control['text']}', "
                                        f"rect={rect}, size={size}")
                except Exception as e:
                    logging.warning(f"コントロール情報の取得中にエラー: handle={hwnd}, error={str(e)}")
            return True
        
        if self.window_handle:
            win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
            logging.info(f"編集可能なコントロールを {len(edit_controls)} 個検出")
        else:
            logging.warning("CTIメインウィンドウが見つかりません")
        
        return edit_controls

    def find_label_by_text(self, text: str) -> Optional[int]:
        """
        指定されたテキストを持つラベルを検索
        
        Args:
            text: 検索するテキスト
            
        Returns:
            Optional[int]: ラベルのハンドル
        """
        label_hwnd = None
        
        def callback(hwnd, _):
            nonlocal label_hwnd
            if win32gui.IsWindowVisible(hwnd):
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name.startswith("WindowsForms10.") and ("STATIC" in class_name or "Label" in class_name):
                        control_text = self.get_control_text(hwnd)
                        if control_text == text:
                            label_hwnd = hwnd
                            return False
                except Exception as e:
                    logging.warning(f"ラベル検索中にエラー: handle={hwnd}, error={str(e)}")
            return True
        
        if self.window_handle:
            win32gui.EnumChildWindows(self.window_handle, callback, None)
        
        return label_hwnd

    def find_field_near_label(self, label_hwnd: int, all_controls: bool = False) -> Optional[Dict]:
        """
        ラベルの近くにあるフィールドを検索
        
        Args:
            label_hwnd: ラベルのハンドル
            all_controls: すべてのコントロールを検索対象にするかどうか
            
        Returns:
            Optional[Dict]: フィールド情報
        """
        if not label_hwnd:
            return None
            
        try:
            # ラベルの位置を取得
            label_rect = win32gui.GetWindowRect(label_hwnd)
            label_client_rect = list(label_rect)
            label_client_rect[0], label_client_rect[1] = win32gui.ScreenToClient(self.window_handle, (label_rect[0], label_rect[1]))
            label_client_rect[2], label_client_rect[3] = win32gui.ScreenToClient(self.window_handle, (label_rect[2], label_rect[3]))
            
            # ラベルのテキストを取得
            label_text = self.get_control_text(label_hwnd)
            
            # コントロールを取得
            controls = self.find_all_controls() if all_controls else self.find_edit_controls()
            
            # 最も近いフィールドを探す
            closest_field = None
            min_distance = float('inf')
            
            for control in controls:
                control_rect = control['client_rect']
                
                # 距離を計算
                horizontal_distance = control_rect[0] - label_client_rect[2]
                vertical_distance = abs((control_rect[1] + control_rect[3]) // 2 - 
                                     (label_client_rect[1] + label_client_rect[3]) // 2)
                euclidean_distance = (horizontal_distance**2 + vertical_distance**2)**0.5
                
                # フィールドの条件をチェック
                if (horizontal_distance > 0 and  # ラベルの右側
                    horizontal_distance < 200 and  # 水平距離が200px以内
                    vertical_distance < 30 and  # 垂直距離が30px以内
                    euclidean_distance < min_distance):  # より近いフィールド
                    
                    min_distance = euclidean_distance
                    closest_field = control
            
            return closest_field
            
        except Exception as e:
            logging.error(f"フィールド検索中にエラー: {str(e)}")
            return None

    def get_control_text(self, hwnd: int) -> str:
        """
        コントロールのテキストを取得
        
        Args:
            hwnd: ウィンドウハンドル
            
        Returns:
            str: コントロールのテキスト
        """
        try:
            if not hwnd:
                return ""
            
            # GetWindowTextを試行
            text = win32gui.GetWindowText(hwnd)
            if text:
                return text
            
            # WM_GETTEXTを試行
            buffer_size = 4096
            buffer = ctypes.create_unicode_buffer(buffer_size)
            length = ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_GETTEXT, buffer_size, buffer)
            
            return buffer.value if length > 0 else ""
            
        except Exception as e:
            logging.error(f"テキスト取得エラー: handle={hwnd}, error={str(e)}")
            return ""

    def normalize_postal_code(self, postal_code: str) -> str:
        """
        郵便番号を正規化する
        
        Args:
            postal_code: 郵便番号
            
        Returns:
            str: 正規化された郵便番号
        """
        try:
            # 数字以外を削除
            postal_code = ''.join(c for c in postal_code if c.isdigit())
            
            # 7桁になるまで先頭に0を追加
            postal_code = postal_code.zfill(7)
            
            # XXX-XXXXの形式に変換
            return f"{postal_code[:3]}-{postal_code[3:]}"
            
        except Exception as e:
            logging.error(f"郵便番号の正規化中にエラー: {str(e)}")
            return postal_code

    def normalize_address(self, address: str) -> str:
        """
        住所を正規化する
        
        Args:
            address: 住所
            
        Returns:
            str: 正規化された住所
        """
        try:
            # 全角スペースを半角に変換
            address = address.replace('　', ' ')
            
            # 連続するスペースを1つに
            address = ' '.join(filter(None, address.split(' ')))
            
            # 全角数字とハイフンを半角に変換
            zen_to_han = str.maketrans('０１２３４５６７８９－', '0123456789-')
            address = address.translate(zen_to_han)
            
            # 都道府県名の正規化
            prefectures = [
                '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
                '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
                '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県',
                '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県',
                '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県',
                '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県',
                '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
            ]
            
            # 都道府県名が含まれているかチェック
            has_prefecture = False
            for pref in prefectures:
                if pref in address:
                    has_prefecture = True
                    break
            
            # 都道府県名が含まれていない場合は、先頭に「大阪府」を追加
            if not has_prefecture:
                address = f"大阪府 {address}"
            
            return address.strip()
            
        except Exception as e:
            logging.error(f"住所の正規化中にエラー: {str(e)}")
            return address

    def get_all_fields_data(self) -> Optional[CTIData]:
        """
        すべてのフィールドのデータを取得
        
        Returns:
            Optional[CTIData]: 取得したデータ
        """
        try:
            if not self.window_handle and not self.find_cti_window():
                logging.error("CTIメインウィンドウが見つかりません")
                return None
            
            data = CTIData()
            
            # すべてのコントロールを取得
            controls = self.find_edit_controls()
            
            # リスト名の検出（最優先）
            list_name = self.find_list_name(controls)
            if list_name:
                data.list_name = list_name
                logging.info(f"リスト名を検出: {list_name}")
            
            # 固定位置でのフィールド検出を試みる
            detected_data = self.detect_fields_by_position(controls)
            
            # 検出されたデータを統合（リスト名は上書きしない）
            if detected_data.customer_name:
                data.customer_name = detected_data.customer_name
            if detected_data.address:
                data.address = self.normalize_address(detected_data.address)
            if detected_data.phone:
                data.phone = detected_data.phone
            if detected_data.postal_code:
                data.postal_code = self.normalize_postal_code(detected_data.postal_code)
            if detected_data.management_id:
                data.management_id = detected_data.management_id
            # リスト名は既に設定されている場合は上書きしない
            if not data.list_name and detected_data.list_name:
                data.list_name = detected_data.list_name
            
            # 従来の方法でも検出を試みる
            if not any([data.customer_name, data.address, data.phone, data.postal_code, data.management_id]):
                logging.info("固定位置での検出に失敗したため、従来の方法で検出を試みます")
                for label_text, field_name in self.label_mappings.items():
                    try:
                        logging.debug(f"{label_text} の検索を開始")
                        label_hwnd = self.find_label_by_text(label_text)
                        if label_hwnd:
                            field = self.find_field_near_label(label_hwnd)
                            if field:
                                # リスト名は既に設定されている場合は上書きしない
                                if field_name == "list_name" and data.list_name:
                                    continue
                                value = field['text']
                                if field_name == "postal_code":
                                    value = self.normalize_postal_code(value)
                                elif field_name == "address":
                                    value = self.normalize_address(value)
                                setattr(data, field_name, value)
                                logging.info(f"{label_text}の値を取得: {value}")
                            else:
                                logging.warning(f"{label_text}のフィールドが見つかりません")
                    except Exception as e:
                        logging.error(f"{label_text}の処理中にエラー: {str(e)}")
            
            # 最終的な検出結果のサマリーを出力
            logging.info("=== 最終検出結果 ===")
            logging.info(f"顧客名: {data.customer_name}")
            logging.info(f"住所: {data.address}")
            logging.info(f"電話番号: {data.phone}")
            logging.info(f"郵便番号: {data.postal_code}")
            logging.info(f"管理番号: {data.management_id}")
            logging.info(f"リスト: {data.list_name}")
            
            # データが取得できたかチェック
            if any([data.customer_name, data.address, data.phone, data.postal_code, data.management_id, data.list_name]):
                return data
            else:
                logging.warning("フィールドデータを取得できませんでした")
                return None
                
        except Exception as e:
            logging.error(f"フィールドデータの取得中にエラー: {str(e)}")
            return None

    def parse_customer_info(self, text: str) -> Optional[Dict[str, str]]:
        """
        顧客情報を解析する
        
        Args:
            text (str): 解析する文字列
            
        Returns:
            dict: 解析された顧客情報
        """
        try:
            # 基本情報の抽出
            info = {
                'name': '',
                'postal_code': '',
                'address': '',
                'phone': ''
            }
            
            # 名前の抽出
            name_match = re.search(r'契約者名[:：](.+?)[\n\r]', text)
            if name_match:
                info['name'] = name_match.group(1).strip()
            
            # 郵便番号の抽出
            postal_match = re.search(r'郵便番号[:：](\d{3}-?\d{4})', text)
            if postal_match:
                info['postal_code'] = postal_match.group(1)
                if '-' not in info['postal_code']:
                    info['postal_code'] = f"{info['postal_code'][:3]}-{info['postal_code'][3:]}"
            
            # 住所の抽出
            address_match = re.search(r'住所[:：](.+?)[\n\r]', text)
            if address_match:
                info['address'] = address_match.group(1).strip()
            
            # 電話番号の抽出
            phone_match = re.search(r'電話番号[:：]([\d-]+)', text)
            if phone_match:
                info['phone'] = phone_match.group(1)
            
            return info if any(info.values()) else None
            
        except Exception as e:
            logging.error(f"顧客情報の解析中にエラー: {str(e)}")
            return None

    def find_list_name(self, controls):
        """
        リスト名を検索
        
        Args:
            controls: コントロールのリスト
            
        Returns:
            str: リスト名
        """
        try:
            # 上部のリストフィールドを直接検索（最優先）
            for control in controls:
                text = control['text']
                if not text:
                    continue
                
                # リストフィールドの条件：
                # 1. 上部に位置する（Y座標が100px以下）
                # 2. 適切な幅（150px以上）
                # 3. COMBOBOXまたはテキストボックスである
                client_rect = control['client_rect']
                width = client_rect[2] - client_rect[0]
                class_name = control['class'].lower()  # 小文字に変換して比較
                
                if (client_rect[1] < 100 and  # Y座標が100px以下
                    width >= 150 and  # 幅が150px以上
                    ("combobox" in class_name or "edit" in class_name or "textbox" in class_name)):  # コンボボックスまたはテキストボックス
                    
                    logging.info(f"リストフィールドを直接検出: '{text}', "
                               f"handle={control['hwnd']}, class='{class_name}', "
                               f"client_rect={client_rect}")
                    return text
            
            # リストラベルを検索
            def find_list_label(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        class_name = win32gui.GetClassName(hwnd)
                        if class_name.startswith("WindowsForms10.") and ("STATIC" in class_name or "Label" in class_name):
                            text = self.get_control_text(hwnd)
                            if text == "リスト":
                                field = self.find_field_near_label(hwnd)
                                if field and field['text']:
                                    logging.info(f"リストフィールドをラベルから検出: '{field['text']}', "
                                               f"handle={field['hwnd']}, class='{field['class']}', "
                                               f"client_rect={field['client_rect']}")
                                    return field['text']
                    except Exception as e:
                        logging.warning(f"リストラベル検索中にエラー: handle={hwnd}, error={str(e)}")
                return None
            
            if self.window_handle:
                list_name = win32gui.EnumChildWindows(self.window_handle, find_list_label, None)
                if list_name:
                    return list_name
            
            return None
            
        except Exception as e:
            logging.error(f"リスト名の検索中にエラー: {str(e)}")
            return None

    def detect_fields_by_position(self, controls) -> CTIData:
        """
        画面上の位置情報を使用してフィールドを検出する
        
        Args:
            controls (list): 検出対象のコントロールリスト
            
        Returns:
            CTIData: 検出されたデータ
        """
        data = CTIData()
        
        # 住所ラベルを探す（最優先）
        address_label_hwnd = self.find_label_by_text("住所")
        if address_label_hwnd:
            # ラベルの位置を取得
            label_rect = win32gui.GetWindowRect(address_label_hwnd)
            label_client_rect = list(label_rect)
            label_client_rect[0], label_client_rect[1] = win32gui.ScreenToClient(self.window_handle, (label_rect[0], label_rect[1]))
            label_client_rect[2], label_client_rect[3] = win32gui.ScreenToClient(self.window_handle, (label_rect[2], label_rect[3]))
            
            # 住所フィールドの検索条件
            best_field = None
            min_vertical_distance = float('inf')
            
            for control in controls:
                control_rect = control['client_rect']
                
                # 住所フィールドの条件：
                # 1. ラベルの右側にある
                # 2. 垂直方向の位置が近い
                # 3. 適切な幅を持つ
                horizontal_distance = control_rect[0] - label_client_rect[2]
                vertical_distance = abs((control_rect[1] + control_rect[3]) // 2 - 
                                     (label_client_rect[1] + label_client_rect[3]) // 2)
                width = control_rect[2] - control_rect[0]
                
                if (horizontal_distance > 0 and  # ラベルの右側
                    horizontal_distance < 100 and  # 水平距離が100px以内
                    vertical_distance < 20 and  # 垂直距離が20px以内
                    width >= 200 and  # 幅が200px以上
                    vertical_distance < min_vertical_distance):  # より近いフィールドを優先
                    
                    min_vertical_distance = vertical_distance
                    best_field = control
            
            # 最適な住所フィールドが見つかった場合
            if best_field:
                data.address = best_field['text']
                logging.info(f"住所ラベルの近くで住所フィールドを検出: '{best_field['text']}', "
                           f"handle={best_field['hwnd']}, client_rect={best_field['client_rect']}")
        
        # 住所ラベルから検出できなかった場合のバックアップ処理
        if not data.address:
            # 住所フィールドの特徴を持つコントロールを探す
            for control in controls:
                text = control['text']
                if not text:
                    continue
                
                # 住所フィールドの条件：
                # 1. 都道府県名を含む
                # 2. 市区町村名を含む
                # 3. 適切な長さ（極端に長い場合は営業メモの可能性）
                # 4. 適切な位置（画面上部のメイン情報エリア内）
                # 5. 適切な幅（住所フィールドは通常広い）
                client_rect = control['client_rect']
                width = client_rect[2] - client_rect[0]
                
                if (width >= 200 and  # 幅が200px以上
                    50 <= client_rect[1] <= 300 and  # Y座標が適切な範囲内
                    len(text) <= 100 and  # 長すぎないテキスト
                    re.search(r'[都道府県]', text) and  # 都道府県名を含む
                    re.search(r'[市区町村]', text) and  # 市区町村名を含む
                    not re.search(r'対応者|工事希望日|料金', text)):  # 営業メモの特徴的な文字列を含まない
                    
                    data.address = text
                    logging.info(f"住所フィールドの特徴から検出: '{text}', "
                               f"handle={control['hwnd']}, client_rect={client_rect}")
                    break

        # 電話番号の検出
        for control in controls:
            text = control['text']
            if not text:
                continue
            
            # 電話番号の条件：
            # 1. 数字とハイフンのみで構成
            # 2. 適切な長さ（10-13文字）
            # 3. 適切な位置（上部エリア内）
            if (all(c in '0123456789-' for c in text) and
                10 <= len(text) <= 13 and
                50 <= control['client_rect'][1] <= 200):
                
                data.phone = text
                logging.info(f"電話番号フィールドを検出: '{text}', "
                           f"handle={control['hwnd']}, client_rect={control['client_rect']}")
                break

        # 郵便番号の検出
        postal_label_hwnd = self.find_label_by_text("〒")
        if postal_label_hwnd:
            # ラベルの位置を取得
            label_rect = win32gui.GetWindowRect(postal_label_hwnd)
            label_client_rect = list(label_rect)
            label_client_rect[0], label_client_rect[1] = win32gui.ScreenToClient(self.window_handle, (label_rect[0], label_rect[1]))
            label_client_rect[2], label_client_rect[3] = win32gui.ScreenToClient(self.window_handle, (label_rect[2], label_rect[3]))
            
            # ラベルの中心座標を計算
            label_center_x = (label_client_rect[0] + label_client_rect[2]) // 2
            label_center_y = (label_client_rect[1] + label_client_rect[3]) // 2
            
            # すべてのコントロールの位置情報をログ出力（デバッグ用）
            for control in controls:
                text = control['text']
                if text and all(c in '0123456789-' for c in text) and 7 <= len(text) <= 8:
                    logging.info(f"郵便番号候補: text='{text}', rect={control['client_rect']}")
            
            # 郵便番号フィールドの検出（直接検出）
            for control in controls:
                text = control['text']
                if not text:
                    continue
                
                # 郵便番号の条件：
                # 1. 数字とハイフンのみで構成
                # 2. 7-8文字（ハイフンあり/なし）
                # 3. 適切な幅（30-100px）
                control_rect = control['client_rect']
                width = control_rect[2] - control_rect[0]
                
                if (all(c in '0123456789-' for c in text) and
                    7 <= len(text) <= 8 and
                    30 <= width <= 100):  # 幅が30-100px
                    
                    data.postal_code = text
                    logging.info(f"郵便番号フィールドを直接検出: '{text}', "
                               f"handle={control['hwnd']}, client_rect={control_rect}")
                    break

            # 直接検出で見つからなかった場合、ラベルからの相対位置で検索
            if not data.postal_code:
                for control in controls:
                    text = control['text']
                    if not text:
                        continue
                    
                    control_rect = control['client_rect']
                    control_center_x = (control_rect[0] + control_rect[2]) // 2
                    control_center_y = (control_rect[1] + control_rect[3]) // 2
                    
                    # 距離を計算（ラベルの中心からの距離）
                    horizontal_distance = control_center_x - label_center_x
                    vertical_distance = abs(control_center_y - label_center_y)
                    euclidean_distance = ((horizontal_distance**2 + vertical_distance**2)**0.5)
                    
                    # 郵便番号フィールドの検証をログ出力
                    if all(c in '0123456789-' for c in text):
                        logging.info(f"郵便番号フィールド検証: text='{text}', "
                                   f"rect={control_rect}, "
                                   f"distances=(h={horizontal_distance}, v={vertical_distance}, e={euclidean_distance})")
                    
                    # 郵便番号フィールドの条件チェック
                    if (all(c in '0123456789-' for c in text) and  # 数字とハイフンのみ
                        7 <= len(text) <= 8 and  # 適切な長さ
                        horizontal_distance > 0 and  # ラベルの右側にある
                        horizontal_distance < 200 and  # ラベルの右側200px以内
                        vertical_distance < 50):  # 垂直距離が50px以内
                        
                        data.postal_code = text
                        logging.info(f"郵便番号フィールドをラベルから検出: '{text}', "
                                   f"handle={control['hwnd']}, client_rect={control_rect}")
                        break

        # 管理番号の検出
        for control in controls:
            text = control['text']
            if not text:
                continue
            
            # 管理番号の条件：
            # 1. アンダースコアを含む
            # 2. 適切な長さ（10文字以上）
            # 3. 適切な位置（上部エリア内）
            if ('_' in text and
                len(text) > 10 and
                50 <= control['client_rect'][1] <= 200):
                
                data.management_id = text
                logging.info(f"管理番号フィールドを検出: '{text}', "
                           f"handle={control['hwnd']}, client_rect={control['client_rect']}")
                break

        # 顧客名が検出されていない場合、ラベルベースの方法で検出
        if not data.customer_name:
            customer_label_hwnd = self.find_label_by_text("顧客・漢字")
            if customer_label_hwnd:
                customer_field = self.find_field_near_label(customer_label_hwnd)
                if customer_field and customer_field['text']:
                    data.customer_name = customer_field['text']
                    logging.info(f"顧客名フィールドを直接検出: '{customer_field['text']}', "
                               f"handle={customer_field['hwnd']}, class='{customer_field['class']}', "
                               f"client_rect={customer_field['client_rect']}")

        # 最終的な検出結果のサマリーを出力
        logging.info("=== 検出結果 ===")
        logging.info(f"顧客名: {data.customer_name}")
        logging.info(f"住所: {data.address}")
        logging.info(f"電話番号: {data.phone}")
        logging.info(f"郵便番号: {data.postal_code}")
        logging.info(f"管理番号: {data.management_id}")
        logging.info(f"リスト: {data.list_name}")
        
        return data

    def get_status(self) -> str:
        """
        CTIの現在の状態を取得する
        
        Returns:
            str: CTIの状態（"waiting", "dialing", "talking"のいずれか）
                見つからない場合は空文字列を返す
        """
        try:
            if not self.window_handle:
                if not self.find_cti_window():
                    logging.debug("CTIウィンドウが見つかりません")
                    return ""
            
            # 状態表示コントロールを探す
            status_controls = []
            
            def enum_callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        text = win32gui.GetWindowText(hwnd)
                        if text and any(status.value in text for status in CTIStatus):
                            class_name = win32gui.GetClassName(hwnd)
                            rect = win32gui.GetWindowRect(hwnd)
                            status_controls.append({
                                'handle': hwnd,
                                'text': text,
                                'class': class_name,
                                'rect': rect
                            })
                    except Exception:
                        pass
                return True
            
            win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
            
            if status_controls:
                # 最初に見つかったコントロールを使用
                status_text = status_controls[0]['text'].strip()
                
                # 状態テキストを内部コードに変換
                status_code = self.status_labels.get(status_text, "")
                
                # 状態が変化した場合のみログを出力
                if status_code and status_code != self._previous_status:
                    logging.info(f"CTI状態が変化: {self._previous_status} → {status_code}")
                    self._previous_status = status_code
                
                return status_code
            else:
                logging.debug("状態表示コントロールが見つかりません")
                return ""
            
        except Exception as e:
            logging.error(f"CTI状態の取得中にエラー: {str(e)}")
            return "" 