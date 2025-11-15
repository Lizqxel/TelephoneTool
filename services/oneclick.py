"""
ワンクリック情報取得サービス

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

# ログレベルをINFOに設定
logging.getLogger().setLevel(logging.DEBUG)

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
        logging.debug("OneClickService initialized")

    def is_edit_control(self, hwnd):
        """
        コントロールが編集可能なテキストボックスかどうかを判定
        
        Args:
            hwnd: ウィンドウハンドル
            
        Returns:
            bool: 編集可能なテキストボックスの場合True
        """
        try:
            class_name = win32gui.GetClassName(hwnd)
            if not class_name.startswith("WindowsForms10."):
                return False
                
            # スタイルを取得
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            
            # 編集可能なテキストボックスの条件をチェック
            return ("EDIT" in class_name or "TextBox" in class_name) and not (style & win32con.ES_READONLY)
            
        except Exception as e:
            logging.warning(f"コントロールチェック中にエラー: {e}")
            return False

    def find_closest_textbox(self, label_hwnd, controls):
        """
        ラベルに最も近いテキストボックスを探す
        
        Args:
            label_hwnd: ラベルのウィンドウハンドル
            controls: 編集可能なコントロールのリスト
            
        Returns:
            tuple: (テキストボックスのハンドル, 距離)
        """
        try:
            label_rect = win32gui.GetWindowRect(label_hwnd)
            label_text = win32gui.GetWindowText(label_hwnd)
            logging.info(f"ラベルの検索開始: text='{label_text}', rect={label_rect}")
            
            # ラベルの中心座標を計算
            label_center_x = (label_rect[0] + label_rect[2]) // 2
            label_center_y = (label_rect[1] + label_rect[3]) // 2
            
            closest_hwnd = None
            min_distance = float('inf')
            
            # 郵便番号フィールド候補を格納するリスト
            postal_candidates = []
            
            # すべてのコントロールをログ出力（デバッグ用）
            logging.info(f"検索対象のコントロール数: {len(controls)}")
            
            for control in controls:
                control_rect = control['rect']
                control_text = control['text']
                control_class = control['class']
                control_size = control['size']
                
                # コントロールの中心座標を計算
                control_center_x = (control_rect[0] + control_rect[2]) // 2
                control_center_y = (control_rect[1] + control_rect[3]) // 2
                
                # 距離を計算
                horizontal_distance = control_center_x - label_center_x  # コントロールの中心 - ラベルの中心
                vertical_distance = abs(control_center_y - label_center_y)
                euclidean_distance = ((horizontal_distance**2 + vertical_distance**2)**0.5)
                
                # 郵便番号フィールドの特別処理
                if label_text == "〒" or "郵便" in label_text:
                    logging.info(f"郵便番号フィールドの検証: text='{control_text}', class='{control_class}', "
                               f"size={control_size}, rect={control_rect}, "
                               f"distances=(h={horizontal_distance}, v={vertical_distance}, e={euclidean_distance})")
                    
                    # 郵便番号フィールドの条件
                    # 特に「〒」ラベルの右側にあるフィールドを検出
                    if (horizontal_distance > 0 and  # ラベルの右側
                        horizontal_distance < 100 and  # 水平距離が100px以内
                        vertical_distance < 20 and  # 垂直距離が20px以内
                        control_text and  # テキストが存在する
                        all(c in '0123456789-' for c in control_text) and  # 数字とハイフンのみ
                        7 <= len(control_text) <= 8):  # 適切な長さ
                        
                        postal_candidates.append((control['hwnd'], vertical_distance, control_text))
                        logging.info(f"郵便番号フィールド候補を追加: handle={control['hwnd']}, "
                                   f"text='{control_text}', distance={vertical_distance}")
                    continue
                
                # 優先電話番号フィールドの特別処理
                if label_text == "優先電話番号1":
                    logging.info(f"優先電話番号フィールドの検証: text='{control_text}', class='{control_class}', "
                               f"size={control_size}, rect={control_rect}, "
                               f"distances=(h={horizontal_distance}, v={vertical_distance})")
                    
                    if (horizontal_distance > 0 and  # テキストボックスがラベルより右
                        horizontal_distance < 150 and  # 水平距離が150ピクセル以内
                        vertical_distance < 15 and  # 垂直距離が15ピクセル以内
                        80 <= control_size[0] <= 150):  # 幅が80-150px
                        
                        # 電話番号の形式チェック（数字とハイフンのみ）
                        if control_text and all(c in '0123456789-' for c in control_text):
                            return control['hwnd'], euclidean_distance
                    continue
                
                # リストフィールドの特別処理
                if label_text == "リスト":
                    logging.info(f"リストフィールドの検証: text='{control_text}', class='{control_class}', "
                               f"size={control_size}, rect={control_rect}, "
                               f"distances=(h={horizontal_distance}, v={vertical_distance})")
                    
                    if (control_size[0] >= 150 and  # 幅が150px以上
                        vertical_distance < 30):  # 垂直距離が30ピクセル以内
                        logging.info(f"リストフィールドを検出: handle={control['hwnd']}, text='{control_text}'")
                        return control['hwnd'], 0
                    continue
                
                # 住所フィールドの特別処理
                if label_text == "住所":
                    if (horizontal_distance > 0 and  # テキストボックスがラベルより右
                        vertical_distance < 20 and  # 垂直距離が20ピクセル以内
                        control_size[0] > 300):  # 幅が300px以上
                        logging.info(f"住所フィールドを検出: handle={control['hwnd']}, text='{control_text}'")
                        return control['hwnd'], 0
                    continue
                
                # 通常のフィールド
                if (horizontal_distance > 0 and  # テキストボックスがラベルより右
                    vertical_distance < 20):  # 垂直距離が20ピクセル以内
                    if euclidean_distance < min_distance:
                        min_distance = euclidean_distance
                        closest_hwnd = control['hwnd']
                        logging.debug(f"より近いフィールドを検出: distance={euclidean_distance}, text='{control_text}'")
            
            # 郵便番号フィールドの候補がある場合、最も近いものを返す
            if postal_candidates:
                # 垂直距離でソート
                postal_candidates.sort(key=lambda x: x[1])
                best_candidate = postal_candidates[0]
                logging.info(f"最適な郵便番号フィールドを選択: handle={best_candidate[0]}, "
                           f"text='{best_candidate[2]}', distance={best_candidate[1]}")
                return best_candidate[0], best_candidate[1]
            
            if closest_hwnd:
                logging.info(f"最も近いフィールドを検出: handle={closest_hwnd}, distance={min_distance}")
            else:
                logging.warning("適切なフィールドが見つかりません")
            return closest_hwnd, min_distance
            
        except Exception as e:
            logging.error(f"テキストボックス検索中にエラー: {str(e)}")
            return None, float('inf')

    def find_edit_controls(self):
        """すべての編集可能なテキストボックスを取得"""
        edit_controls = []
        
        def enum_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                class_name = win32gui.GetClassName(hwnd)
                if not class_name.startswith("WindowsForms10."):
                    return True
                
                # スタイルを取得
                style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                
                # 編集可能なテキストボックスの条件をチェック
                if "EDIT" in class_name or "TextBox" in class_name or "RICHEDIT" in class_name:
                    text = self.get_control_text(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    
                    # スクリーン座標をクライアント座標に変換
                    client_rect = list(rect)
                    client_rect[0], client_rect[1] = win32gui.ScreenToClient(self.window_handle, (rect[0], rect[1]))
                    client_rect[2], client_rect[3] = win32gui.ScreenToClient(self.window_handle, (rect[2], rect[3]))
                    
                    size = (rect[2] - rect[0], rect[3] - rect[1])  # (width, height)
                    
                    logging.debug(f"編集可能なコントロールを検出: text='{text}', class='{class_name}', size={size}, "
                                f"screen_rect={rect}, client_rect={client_rect}")
                    
                    edit_controls.append({
                        'hwnd': hwnd,
                        'text': text,
                        'rect': rect,
                        'client_rect': client_rect,
                        'class': class_name,
                        'style': style,
                        'size': size
                    })
            except Exception as e:
                logging.warning(f"コントロール列挙中にエラー: {e}")
            return True
        
        win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
        logging.info(f"編集可能なコントロールを {len(edit_controls)} 個検出")
        return edit_controls

    def find_cti_window(self) -> bool:
        """CTIメインウィンドウを検索"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    window_text = win32gui.GetWindowText(hwnd)
                    window_class = win32gui.GetClassName(hwnd)
                    logging.debug(f"検出されたウィンドウ: text='{window_text}', class='{window_class}', handle={hwnd}")
                    if "CTIメイン" in window_text:
                        self.window_handle = hwnd
                        logging.info(f"CTIメインウィンドウを検出: handle={hwnd}")
                        return False
                except Exception as e:
                    logging.error(f"ウィンドウ情報取得エラー: {str(e)}")
            return True
        
        try:
            logging.debug("CTIメインウィンドウの検索を開始")
            self.window_handle = None  # ハンドルをリセット
            win32gui.EnumWindows(callback, None)
            
            if self.window_handle:
                # ウィンドウの状態を確認
                window_text = win32gui.GetWindowText(self.window_handle)
                window_class = win32gui.GetClassName(self.window_handle)
                is_visible = win32gui.IsWindowVisible(self.window_handle)
                is_enabled = win32gui.IsWindowEnabled(self.window_handle)
                
                logging.info(f"CTIメインウィンドウの詳細:")
                logging.info(f"- テキスト: {window_text}")
                logging.info(f"- クラス名: {window_class}")
                logging.info(f"- 可視状態: {is_visible}")
                logging.info(f"- 有効状態: {is_enabled}")
                
                # ウィンドウの子要素を列挙（デバッグ用）
                def enum_child_windows(hwnd, param):
                    try:
                        text = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        logging.debug(f"子ウィンドウ: text='{text}', class='{class_name}', handle={hwnd}")
                    except Exception as e:
                        logging.error(f"子ウィンドウ情報取得エラー: {str(e)}")
                    return True
                
                win32gui.EnumChildWindows(self.window_handle, enum_child_windows, None)
                return True
            else:
                logging.error("CTIメインウィンドウが見つかりません")
                return False
                
        except Exception as e:
            logging.error(f"ウィンドウ検索中にエラー: {str(e)}")
            return False

    def find_field_and_value(self, field_name):
        """
        指定されたフィールド名に対応するラベルとテキストボックスを探す
        
        Args:
            field_name: フィールド名
            
        Returns:
            tuple: (ラベルのハンドル, テキストボックスのハンドル, テキストボックスの値)
        """
        label_hwnd = None
        edit_hwnd = None
        value = ""
        
        try:
            # ラベルを探す
            def find_label():
                nonlocal label_hwnd  # 関数の先頭でnonlocal宣言
                
                def enum_callback(hwnd, _):
                    if not win32gui.IsWindowVisible(hwnd):
                        return True
                    
                    class_name = win32gui.GetClassName(hwnd)
                    if not class_name.startswith("WindowsForms10."):
                        return True
                    
                    if "STATIC" in class_name or "Label" in class_name:
                        text = self.get_control_text(hwnd)
                        # 郵便番号フィールドの特別処理
                        if field_name == "〒":
                            if text == "〒" or "郵便" in text:
                                rect = win32gui.GetWindowRect(hwnd)
                                logging.info(f"郵便番号ラベルを検出: text='{text}', rect={rect}")
                                label_hwnd = hwnd
                                return False
                        else:
                            if text == field_name:
                                label_hwnd = hwnd
                                return False
                    return True
                
                win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
                return label_hwnd
            
            label_hwnd = find_label()
            if not label_hwnd:
                logging.warning(f"{field_name}のフィールドが見つかりません")
                return None, None, ""
                
            # 編集可能なテキストボックスを取得
            controls = self.find_edit_controls()
            
            # ラベルに最も近いテキストボックスを探す
            edit_hwnd, _ = self.find_closest_textbox(label_hwnd, controls)
            
            if not edit_hwnd:
                logging.warning(f"{field_name}のフィールドまたは値が見つかりません")
                return label_hwnd, None, ""
                
            # テキストボックスの値を取得
            value = self.get_control_text(edit_hwnd)
            logging.info(f"{field_name}の値を取得: {value}")
            
        except Exception as e:
            logging.error(f"フィールドの検索中にエラー: {str(e)}")
            return None, None, ""
            
        return label_hwnd, edit_hwnd, value

    def get_control_text(self, hwnd) -> str:
        """
        コントロールのテキストを取得
        
        Args:
            hwnd: ウィンドウハンドル
            
        Returns:
            str: コントロールのテキスト
        """
        try:
            if not hwnd:
                logging.error("無効なウィンドウハンドル")
                return ""
            
            # 複数のエンコーディングを試行
            encodings = ['shift-jis', 'cp932', 'utf-8']
            
            # GetWindowTextを試行
            text = win32gui.GetWindowText(hwnd)
            if text:
                return text
            
            # WM_GETTEXTを試行
            buffer_size = 4096
            for encoding in encodings:
                try:
                    # ANSI版
                    buffer = ctypes.create_string_buffer(buffer_size)
                    length = ctypes.windll.user32.SendMessageA(hwnd, win32con.WM_GETTEXT, buffer_size, buffer)
                    if length > 0:
                        text = buffer.raw[:length].decode(encoding)
                        if text:
                            return text
                except:
                    continue
            
            # Unicode版を試行
            try:
                buffer = ctypes.create_unicode_buffer(buffer_size)
                length = ctypes.windll.user32.SendMessageW(hwnd, win32con.WM_GETTEXT, buffer_size, buffer)
                if length > 0:
                    return buffer.value
            except:
                pass
            
            return ""
            
        except Exception as e:
            logging.error(f"テキスト取得エラー: handle={hwnd}, error={str(e)}")
            return ""

    def get_all_fields_data(self) -> Optional[CTIData]:
        """全フィールドのデータを取得"""
        logging.debug("データ取得を開始")
        
        if not self.window_handle:
            if not self.find_cti_window():
                logging.error("CTIメインウィンドウが見つかりません")
                return None

        data = CTIData()
        
        # すべてのコントロール（編集可能でないものも含む）を取得
        all_controls = self.find_all_controls()
        
        # 上部のリストフィールドを直接検索（最優先）
        # 画像から確認した位置情報に基づいて上部のリストフィールドを検索
        for control in all_controls:
            if control['text'] and "【NP光在宅】" in control['text']:
                # 上部にあるリストフィールドの位置を確認（Y座標が小さい＝上部にある）
                client_rect = control['client_rect']
                # 上部のリストフィールドは通常Y座標が100以下
                if client_rect[1] < 100:
                    data.list_name = control['text']
                    logging.info(f"上部のリストフィールドを直接検出: '{control['text']}', handle={control['hwnd']}, "
                               f"class='{control['class']}', client_rect={client_rect}")
                    break
        
        # 上部のリストフィールドが見つからない場合、リストラベルを検索
        if not data.list_name:
            list_label_hwnd = self.find_label_by_text("リスト")
            if list_label_hwnd:
                list_field = self.find_field_near_label(list_label_hwnd, all_controls=True)
                if list_field and list_field['text']:
                    data.list_name = list_field['text']
                    logging.info(f"リストフィールドを直接検出: '{list_field['text']}', "
                               f"handle={list_field['hwnd']}, class='{list_field['class']}', "
                               f"client_rect={list_field['client_rect']}")
        
        # 特定のキーワードを含むリスト名を検索（バックアップ）
        if not data.list_name:
            for control in all_controls:
                if control['text'] and "【NP光在宅】" in control['text']:
                    data.list_name = control['text']
                    logging.info(f"リスト名を直接検出: '{control['text']}', handle={control['hwnd']}, "
                               f"class='{control['class']}', client_rect={control['client_rect']}")
                    break
        
        # 編集可能なコントロールを取得
        controls = self.find_edit_controls()
        
        # 固定位置でのフィールド検出を試みる
        detected_data = self.detect_fields_by_position(controls)
        
        # 検出されたデータを統合（リスト名は上書きしない）
        if detected_data.customer_name:
            data.customer_name = detected_data.customer_name
        if detected_data.address:
            data.address = detected_data.address
        if detected_data.phone:
            data.phone = detected_data.phone
        if detected_data.postal_code:
            data.postal_code = detected_data.postal_code
        if detected_data.management_id:
            data.management_id = detected_data.management_id
        # リスト名は既に設定されている場合は上書きしない
        if not data.list_name and detected_data.list_name:
            data.list_name = detected_data.list_name
        
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
        
        # 従来の方法でも検出を試みる
        if not any([data.customer_name, data.address, data.phone, data.postal_code, data.management_id]):
            logging.info("固定位置での検出に失敗したため、従来の方法で検出を試みます")
            for label, attr_name in self.label_mappings.items():
                try:
                    logging.debug(f"{label} の検索を開始")
                    label_handle, value_handle, value = self.find_field_and_value(label)
                    if label_handle and value_handle:
                        # リスト名は既に設定されている場合は上書きしない
                        if attr_name == "list_name" and data.list_name:
                            continue
                        setattr(data, attr_name, value)
                        logging.info(f"{label}の値を取得: {value}")
                    else:
                        logging.warning(f"{label}のフィールドが見つかりません")
                except Exception as e:
                    logging.error(f"{label}の処理中にエラー: {str(e)}")
        
        # 最終的な検出結果のサマリーを出力
        logging.info("=== 最終検出結果 ===")
        logging.info(f"顧客名: {data.customer_name}")
        logging.info(f"住所: {data.address}")
        logging.info(f"電話番号: {data.phone}")
        logging.info(f"郵便番号: {data.postal_code}")
        logging.info(f"管理番号: {data.management_id}")
        logging.info(f"リスト: {data.list_name}")

        return data

    def find_all_controls(self):
        """すべてのコントロール（編集可能でないものも含む）を取得"""
        all_controls = []
        
        def enum_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                
                text = self.get_control_text(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                
                # スクリーン座標をクライアント座標に変換
                client_rect = list(rect)
                client_rect[0], client_rect[1] = win32gui.ScreenToClient(self.window_handle, (rect[0], rect[1]))
                client_rect[2], client_rect[3] = win32gui.ScreenToClient(self.window_handle, (rect[2], rect[3]))
                
                size = (rect[2] - rect[0], rect[3] - rect[1])  # (width, height)
                
                logging.debug(f"コントロールを検出: text='{text}', class='{class_name}', size={size}, "
                            f"screen_rect={rect}, client_rect={client_rect}")
                
                all_controls.append({
                    'hwnd': hwnd,
                    'text': text,
                    'rect': rect,
                    'client_rect': client_rect,
                    'class': class_name,
                    'size': size
                })
            except Exception as e:
                logging.warning(f"コントロール列挙中にエラー: {e}")
            return True
        
        win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
        logging.info(f"すべてのコントロールを {len(all_controls)} 個検出")
        return all_controls

    def find_list_name(self, controls):
        """
        リスト名を検出する特別な方法
        
        Args:
            controls: コントロールのリスト
            
        Returns:
            str: 検出されたリスト名
        """
        # リスト名の候補を格納する配列
        list_candidates = []
        
        # リスト名の特徴的なキーワード
        list_keywords = ["NP光", "在宅", "アナログ", "リスト", "西日本", "202410"]
        
        # すべてのコントロールを検索
        for control in controls:
            text = control['text']
            if text:
                # リスト名の特徴的なキーワードを含むテキストを検出
                for keyword in list_keywords:
                    if keyword in text:
                        # 特に「【NP光在宅】」を含むテキストを優先
                        if "【NP光在宅】" in text:
                            logging.info(f"最適なリスト名を直接検出: '{text}', handle={control['hwnd']}, "
                                       f"class='{control['class']}', client_rect={control['client_rect']}")
                            return text
                        
                        list_candidates.append(text)
                        logging.info(f"リスト名候補を検出: '{text}', handle={control['hwnd']}, "
                                   f"class='{control['class']}', client_rect={control['client_rect']}")
                        break
        
        # リスト名の候補から最適なものを選択
        if list_candidates:
            # 最も長いテキストを選択（より詳細な情報を含む可能性が高い）
            best_candidate = max(list_candidates, key=len)
            logging.info(f"最適なリスト名を選択: '{best_candidate}'")
            return best_candidate
        
        # リストラベルの近くにあるテキストを探す
        list_label_hwnd = None
        
        def find_list_label(hwnd, _):
            nonlocal list_label_hwnd
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                if "STATIC" in class_name or "Label" in class_name:
                    text = self.get_control_text(hwnd)
                    if text == "リスト":
                        list_label_hwnd = hwnd
                        return False
            return True
        
        win32gui.EnumChildWindows(self.window_handle, find_list_label, None)
        
        if list_label_hwnd:
            label_rect = win32gui.GetWindowRect(list_label_hwnd)
            
            # リストラベルの近くにあるテキストを持つコントロールを探す
            for control in controls:
                control_rect = control['rect']
                # リストラベルの右側にあるコントロール
                if (control_rect[0] > label_rect[2] and 
                    abs(control_rect[1] - label_rect[1]) < 50 and
                    control['text']):
                    logging.info(f"リストラベルの近くでテキストを検出: '{control['text']}'")
                    return control['text']
        
        return ""

    def detect_fields_by_position(self, controls):
        """
        画面上の位置情報を使用してフィールドを検出する
        
        Args:
            controls (list): 検出対象のコントロールリスト
            
        Returns:
            CTIData: 検出されたデータ
        """
        data = CTIData()

        # リスト名の検出（最優先）
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
                
                data.list_name = text
                logging.info(f"リストフィールドを検出: '{text}'")
                break

        # リストフィールドが見つからない場合、ラベルからの検出を試みる
        if not data.list_name:
            list_label_hwnd = self.find_label_by_text("リスト")
            if list_label_hwnd:
                # ラベルの位置を取得
                label_rect = win32gui.GetWindowRect(list_label_hwnd)
                label_client_rect = list(label_rect)
                label_client_rect[0], label_client_rect[1] = win32gui.ScreenToClient(self.window_handle, (label_rect[0], label_rect[1]))
                label_client_rect[2], label_client_rect[3] = win32gui.ScreenToClient(self.window_handle, (label_rect[2], label_rect[3]))
                
                # ラベルの中心座標を計算
                label_center_x = (label_client_rect[0] + label_client_rect[2]) // 2
                label_center_y = (label_client_rect[1] + label_client_rect[3]) // 2
                
                for control in controls:
                    text = control['text']
                    if not text:
                        continue
                    
                    control_rect = control['client_rect']
                    control_center_x = (control_rect[0] + control_rect[2]) // 2
                    control_center_y = (control_rect[1] + control_rect[3]) // 2
                    
                    # 距離を計算
                    horizontal_distance = control_center_x - label_center_x
                    vertical_distance = abs(control_center_y - label_center_y)
                    class_name = control['class'].lower()  # 小文字に変換して比較
                    
                    if (horizontal_distance > 0 and  # ラベルの右側
                        horizontal_distance < 100 and  # 水平距離が100px以内
                        vertical_distance < 20 and  # 垂直距離が20px以内
                        ("combobox" in class_name or "edit" in class_name or "textbox" in class_name)):  # コンボボックスまたはテキストボックス
                        
                        data.list_name = text
                        logging.info(f"リストフィールドをラベルから検出: '{text}'")
                        break

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
        logging.info("=== 最終検出結果 ===")
        logging.info(f"顧客名: {data.customer_name}")
        logging.info(f"住所: {data.address}")
        logging.info(f"電話番号: {data.phone}")
        logging.info(f"郵便番号: {data.postal_code}")
        logging.info(f"管理番号: {data.management_id}")
        logging.info(f"リスト: {data.list_name}")
        
        return data

    def find_label_by_text(self, label_text):
        """
        指定されたテキストを持つラベルを検索
        
        Args:
            label_text: ラベルのテキスト
            
        Returns:
            int: ラベルのウィンドウハンドル、見つからない場合はNone
        """
        label_hwnd = None
        
        def enum_callback(hwnd, _):
            nonlocal label_hwnd
            if not win32gui.IsWindowVisible(hwnd):
                return True
            
            class_name = win32gui.GetClassName(hwnd)
            if "STATIC" in class_name or "Label" in class_name:
                text = self.get_control_text(hwnd)
                if text == label_text:
                    label_hwnd = hwnd
                    return False
            return True
        
        win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
        
        if label_hwnd:
            logging.info(f"ラベルを検出: text='{label_text}', handle={label_hwnd}")
        else:
            logging.warning(f"ラベル '{label_text}' が見つかりません")
        
        return label_hwnd

    def find_field_near_label(self, label_hwnd, all_controls=False):
        """
        ラベルの近くにあるフィールドを検索
        
        Args:
            label_hwnd: ラベルのウィンドウハンドル
            all_controls: すべてのコントロールを検索対象にするかどうか
            
        Returns:
            dict: フィールドの情報、見つからない場合はNone
        """
        if not label_hwnd:
            return None
        
        label_rect = win32gui.GetWindowRect(label_hwnd)
        label_text = self.get_control_text(label_hwnd)
        logging.info(f"ラベルの近くのフィールドを検索: text='{label_text}', rect={label_rect}")
        
        # ラベルの中心座標を計算
        label_center_x = (label_rect[0] + label_rect[2]) // 2
        label_center_y = (label_rect[1] + label_rect[3]) // 2
        
        # 検索対象のコントロール
        if all_controls:
            controls = self.find_all_controls()
        else:
            controls = self.find_edit_controls()
        
        closest_field = None
        min_distance = float('inf')
        
        for control in controls:
            control_rect = control['rect']

            # サイズフィルタ: 幅や高さが極端に小さいコントロールは除外
            width = control_rect[2] - control_rect[0]
            height = control_rect[3] - control_rect[1]
            if width < 20 or height < 10:
                logging.debug(
                    f"フィールド候補を除外 (サイズが小さすぎます): "
                    f"text='{control['text']}', class='{control['class']}', "
                    f"rect={control_rect}, width={width}, height={height}"
                )
                continue

            # コントロールの中心座標を計算
            control_center_x = (control_rect[0] + control_rect[2]) // 2
            control_center_y = (control_rect[1] + control_rect[3]) // 2
            
            # 距離を計算
            horizontal_distance = control_center_x - label_center_x
            vertical_distance = abs(control_center_y - label_center_y)
            euclidean_distance = ((horizontal_distance**2 + vertical_distance**2)**0.5)
            
            # ラベルの右側にあるコントロール
            if (horizontal_distance > 0 and 
                horizontal_distance < 200 and 
                vertical_distance < 30):
                
                # リストラベルの場合、COMBOBOXを優先
                if label_text == "リスト" and "COMBOBOX" in control['class']:
                    logging.info(f"リストのコンボボックスを検出: text='{control['text']}', "
                               f"handle={control['hwnd']}, class='{control['class']}', "
                               f"client_rect={control['client_rect']}")
                    return control
                
                if euclidean_distance < min_distance:
                    min_distance = euclidean_distance
                    closest_field = control
                    logging.debug(f"より近いフィールドを検出: distance={euclidean_distance}, "
                                f"text='{control['text']}', class='{control['class']}'")
        
        if closest_field:
            logging.info(f"最も近いフィールドを検出: text='{closest_field['text']}', "
                       f"handle={closest_field['hwnd']}, class='{closest_field['class']}', "
                       f"client_rect={closest_field['client_rect']}")
        else:
            logging.warning(f"ラベル '{label_text}' の近くに適切なフィールドが見つかりません")
        
        return closest_field

    def get_richedit_text(self, hwnd) -> str:
        """
        RICHEDITコントロールのテキストを取得する特別な方法
        
        Args:
            hwnd: ウィンドウハンドル
            
        Returns:
            str: コントロールのテキスト
        """
        try:
            # 通常の方法でテキストを取得
            text = self.get_control_text(hwnd)
            if text:
                return text
            
            # RICHEDITコントロールの場合、EM_GETTEXT/EM_GETTEXTLENGTHメッセージを使用
            EM_GETTEXTLENGTH = 0x000E
            EM_GETTEXT = 0x000D
            
            # テキストの長さを取得
            length = ctypes.windll.user32.SendMessageW(hwnd, EM_GETTEXTLENGTH, 0, 0)
            if length > 0:
                # バッファを確保してテキストを取得
                buffer = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.SendMessageW(hwnd, EM_GETTEXT, length + 1, buffer)
                return buffer.value
            
            # 親ウィンドウを取得して子コントロールを探す
            parent = win32gui.GetParent(hwnd)
            if parent:
                # 親ウィンドウの子コントロールを列挙
                child_texts = []
                
                def enum_child_callback(child_hwnd, _):
                    if child_hwnd != hwnd and "RICHEDIT" in win32gui.GetClassName(child_hwnd):
                        child_text = self.get_control_text(child_hwnd)
                        if child_text:
                            child_texts.append(child_text)
                    return True
                
                win32gui.EnumChildWindows(parent, enum_child_callback, None)
                
                # 子コントロールのテキストを結合
                if child_texts:
                    return " ".join(child_texts)
            
            # リストラベルの近くにあるテキストを探す
            list_label_hwnd = None
            
            def find_list_label(hwnd, _):
                nonlocal list_label_hwnd
                if win32gui.IsWindowVisible(hwnd):
                    class_name = win32gui.GetClassName(hwnd)
                    if "STATIC" in class_name or "Label" in class_name:
                        text = self.get_control_text(hwnd)
                        if text == "リスト":
                            list_label_hwnd = hwnd
                            return False
                return True
            
            win32gui.EnumChildWindows(self.window_handle, find_list_label, None)
            
            if list_label_hwnd:
                label_rect = win32gui.GetWindowRect(list_label_hwnd)
                
                # リストラベルの近くにあるテキストを持つコントロールを探す
                for control in self.find_edit_controls():
                    control_rect = control['rect']
                    # リストラベルの右側にあるコントロール
                    if (control_rect[0] > label_rect[2] and 
                        abs(control_rect[1] - label_rect[1]) < 50 and
                        control['text']):
                        return control['text']
            
            # すべてのコントロールを検索して、リスト名らしきテキストを探す
            for control in self.find_edit_controls():
                if control['text'] and any(keyword in control['text'] for keyword in ["リスト", "NP光", "在宅", "アナログ"]):
                    return control['text']
            
            return ""
            
        except Exception as e:
            logging.error(f"RICHEDITテキスト取得エラー: handle={hwnd}, error={str(e)}")
            return ""
