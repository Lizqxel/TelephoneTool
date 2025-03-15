"""
電話ボタン監視サービス

CTIメインウィンドウの緑色電話ボタンをクリックした際に、
顧客情報の自動取得を行う機能を提供します。
"""

import win32gui
import win32con
import win32api
import logging
from typing import Optional, Callable, List, Tuple
import time
import threading
from ctypes import windll, CFUNCTYPE, POINTER, c_int, c_void_p, byref, Structure, c_long, c_ulong, c_uint, c_ulonglong

class PhoneButtonMonitor:
    """電話ボタン監視クラス"""
    
    def __init__(self, callback: Callable[[], None]):
        """
        初期化
        
        Args:
            callback: 電話ボタンクリック時に実行するコールバック関数
        """
        self.callback = callback
        self.window_handle = None
        self.button_handle = None
        self.is_monitoring = False
        self.monitor_thread = None
        self.button_rect = None
        self.last_redetect_time = 0  # 最後の再検出時間
        self.redetect_interval = 10  # 再検出間隔（秒）
        self.last_click_time = 0  # 最後のクリック時間
        self.click_interval = 0.5  # クリック間隔（秒）
        
    def find_cti_window(self) -> bool:
        """CTIメインウィンドウを検索"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if "CTIメイン" in title:
                        self.window_handle = hwnd
                        return False
                except Exception:
                    pass
            return True
            
        win32gui.EnumWindows(callback, None)
        return self.window_handle is not None

    def find_hold_button(self) -> Optional[Tuple[int, Tuple[int, int, int, int]]]:
        """
        保留ボタンを検索
        
        Returns:
            Optional[Tuple[int, Tuple[int, int, int, int]]]: 
            (ハンドル, (x1, y1, x2, y2))のタプル、見つからない場合はNone
        """
        hold_button = None
        
        def enum_callback(hwnd, _):
            nonlocal hold_button
            if not win32gui.IsWindowVisible(hwnd):
                return True
                
            try:
                text = win32gui.GetWindowText(hwnd)
                if "保留" in text:  # 保留ボタンのテキストで判定
                    rect = win32gui.GetWindowRect(hwnd)
                    hold_button = (hwnd, rect)
                    return False
            except Exception:
                pass
            return True
            
        if self.window_handle:
            win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
            
        return hold_button

    def find_all_buttons_in_target_area(self) -> List[Tuple[int, Tuple[int, int, int, int]]]:
        """
        目的のエリア内のボタンを検索
        - 画面中央より右側
        - 保留ボタンと同じY座標
        - 保留ボタンより左側
        
        Returns:
            List[Tuple[int, Tuple[int, int, int, int]]]: 
            (ハンドル, (x1, y1, x2, y2))のリスト
        """
        buttons = []
        
        # まず保留ボタンを検索
        hold_button = self.find_hold_button()
        if not hold_button:
            return buttons
            
        hold_handle, hold_rect = hold_button
        hold_y = (hold_rect[1] + hold_rect[3]) // 2  # 保留ボタンの中心Y座標
        
        def enum_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
                
            try:
                # ウィンドウの中心座標を取得
                window_rect = win32gui.GetWindowRect(self.window_handle)
                window_center_x = (window_rect[0] + window_rect[2]) // 2
                
                # ボタンの座標を取得
                rect = win32gui.GetWindowRect(hwnd)
                button_center_x = (rect[0] + rect[2]) // 2
                button_center_y = (rect[1] + rect[3]) // 2
                
                # 条件チェック（高速化のため、早期リターン）
                if button_center_x <= window_center_x:
                    return True
                if button_center_x >= hold_rect[0]:
                    return True
                if abs(button_center_y - hold_y) >= 20:
                    return True
                
                buttons.append((hwnd, rect))
            except Exception:
                pass
            return True
            
        if self.window_handle:
            win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
            
        return buttons

    def find_green_phone_button(self) -> bool:
        """緑色電話ボタンを検索"""
        # 目的のエリア内のボタンを検索
        buttons = self.find_all_buttons_in_target_area()
        
        if not buttons:
            return False
        
        # ボタンをX座標でソート（左から右）
        buttons.sort(key=lambda x: x[1][0])
        
        # 左側のボタンを選択
        if len(buttons) >= 1:
            target_button = buttons[0]  # インデックス0が左側のボタン
            self.button_handle = target_button[0]
            return True
        
        return False
        
    def start_monitoring(self):
        """ボタン監視を開始"""
        self.is_monitoring = True
        
        # ボタン再検出用のスレッドを開始
        self.redetect_thread = threading.Thread(target=self._redetect_loop)
        self.redetect_thread.daemon = True
        self.redetect_thread.start()
        
        # マウス監視用のスレッドを開始
        self.monitor_thread = threading.Thread(target=self._monitor_loop)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        
        logging.info("電話ボタン監視を開始しました")

    def _monitor_loop(self):
        """マウス監視ループ"""
        while self.is_monitoring:
            try:
                # マウスクリックをチェック（GetAsyncKeyStateは高速）
                if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                    current_time = time.time()
                    # 連続クリックを防ぐ
                    if current_time - self.last_click_time >= self.click_interval:
                        # マウス座標を取得（GetCursorPosは高速）
                        x, y = win32api.GetCursorPos()
                        
                        if self.button_rect:
                            if (self.button_rect[0] <= x <= self.button_rect[2] and 
                                self.button_rect[1] <= y <= self.button_rect[3]):
                                if self.callback:
                                    try:
                                        self.callback()
                                    except Exception as e:
                                        logging.error(f"コールバック実行中にエラー: {e}")
                                self.last_click_time = current_time
                
                # 最小限の待機時間
                time.sleep(0.001)
            except Exception as e:
                logging.error(f"マウス監視中にエラー: {e}")
                time.sleep(0.1)

    def stop_monitoring(self):
        """ボタン監視を停止"""
        self.is_monitoring = False
        if hasattr(self, 'redetect_thread'):
            self.redetect_thread.join()
        if hasattr(self, 'monitor_thread'):
            self.monitor_thread.join()
        logging.info("電話ボタン監視を停止しました")

    def _redetect_loop(self):
        """定期的にボタンを再検出するループ"""
        while self.is_monitoring:
            try:
                current_time = time.time()
                # 再検出間隔を制御
                if current_time - self.last_redetect_time >= self.redetect_interval:
                    if self.find_cti_window():
                        if self.find_green_phone_button():
                            try:
                                rect = win32gui.GetWindowRect(self.button_handle)
                                if all(isinstance(coord, int) and coord > -32000 for coord in rect):
                                    self.button_rect = rect
                            except Exception as e:
                                logging.error(f"ボタン座標の取得に失敗: {e}")
                                self.button_rect = None
                    self.last_redetect_time = current_time
                
                # 最小限の待機時間
                time.sleep(0.1)
            except Exception as e:
                logging.error(f"ボタン再検出中にエラー: {e}")
                time.sleep(0.5) 