"""
電話ボタン監視サービス

CTIメインウィンドウの緑色電話ボタンをクリックした際に、
顧客情報の自動取得を行う機能を提供します。

主な機能：
- 緑色電話ボタンの監視
- クリック検出と自動処理実行
- 通話終了後の一時停止機能（2秒間）
- エラーハンドリングとログ出力
"""

import win32gui
import win32con
import win32api
import logging
from typing import Optional, Callable, List, Tuple
import time
import threading
from ctypes import windll, CFUNCTYPE, POINTER, c_int, c_void_p, byref, Structure, c_long, c_ulong, c_uint, c_ulonglong
import json
import os

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
        self.delay_seconds = 0  # 遅延時間（秒）
        self.countdown_thread = None  # カウントダウン用スレッド
        self.is_counting_down = False  # カウントダウン中かどうか
        self.settings_file = "settings.json"  # 設定ファイルのパス
        
        # 一時停止関連
        self.is_paused = False
        self.pause_end_time = 0
        self.pause_duration = 2.0  # 一時停止時間（秒）
        
        # 設定の読み込み
        self.load_settings()
        
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.delay_seconds = settings.get('delay_seconds', 0)
            else:
                self.delay_seconds = 0
        except Exception as e:
            logging.error(f"設定の読み込みに失敗しました: {str(e)}")
            self.delay_seconds = 0
            
    def update_settings(self):
        """設定を更新する"""
        self.load_settings()
        
    def start_countdown(self):
        """カウントダウンを開始する"""
        if self.is_counting_down:
            # 既にカウントダウン中の場合は、カウントダウンをリセット
            self.is_counting_down = False
            if self.countdown_thread:
                self.countdown_thread.join(timeout=0.1)
            logging.info("カウントダウンをリセットしました")
            return
            
        if self.delay_seconds <= 0:
            # 遅延時間が0の場合は即座にコールバックを実行
            if self.callback:
                try:
                    self.callback()
                except Exception as e:
                    logging.error(f"コールバック実行中にエラー: {e}")
            return
            
        self.is_counting_down = True
        self.countdown_start_time = time.time()
        self.countdown_thread = threading.Thread(target=self._countdown_loop)
        self.countdown_thread.daemon = True
        self.countdown_thread.start()
        logging.info(f"カウントダウンを開始しました（{self.delay_seconds}秒）")
        
    def _countdown_loop(self):
        """カウントダウンループ"""
        try:
            remaining_time = self.delay_seconds
            while self.is_counting_down and remaining_time > 0:
                logging.info(f"残り時間: {remaining_time}秒")
                time.sleep(1)
                remaining_time -= 1
                
            if self.is_counting_down and remaining_time <= 0:
                # カウントダウンが正常に完了した場合
                if self.callback:
                    try:
                        self.callback()
                    except Exception as e:
                        logging.error(f"コールバック実行中にエラー: {e}")
                self.is_counting_down = False
                logging.info("カウントダウンが完了しました")
        except Exception as e:
            logging.error(f"カウントダウン中にエラーが発生しました: {e}")
            self.is_counting_down = False
        
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

    def pause_monitoring(self):
        """監視を一時停止（2秒間）"""
        self.is_paused = True
        self.pause_end_time = time.time() + self.pause_duration
        logging.info(f"電話ボタン監視を{self.pause_duration}秒間一時停止します")

    def _check_pause_status(self) -> bool:
        """
        一時停止状態をチェック
        
        Returns:
            bool: 一時停止中ならTrue
        """
        if self.is_paused:
            current_time = time.time()
            if current_time >= self.pause_end_time:
                self.is_paused = False
                logging.info("電話ボタン監視の一時停止が終了しました")
                return False
            return True
        return False

    def _monitor_loop(self):
        """監視ループ"""
        while self.is_monitoring:
            try:
                # 一時停止中はスキップ
                if self._check_pause_status():
                    time.sleep(0.1)
                    continue
                    
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
                                # カウントダウンを開始
                                self.start_countdown()
                                self.last_click_time = current_time
                
                # 最小限の待機時間
                time.sleep(0.001)
            except Exception as e:
                logging.error(f"マウス監視中にエラー: {e}")
                time.sleep(0.1)

    def stop_monitoring(self):
        """ボタン監視を停止"""
        self.is_monitoring = False
        self.is_counting_down = False
        if hasattr(self, 'redetect_thread'):
            self.redetect_thread.join()
        if hasattr(self, 'monitor_thread'):
            self.monitor_thread.join()
        if self.countdown_thread:
            self.countdown_thread.join()
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