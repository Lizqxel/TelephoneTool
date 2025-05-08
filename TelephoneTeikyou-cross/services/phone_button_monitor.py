"""
電話ボタン監視サービス

CTIメインウィンドウの緑色電話ボタンをクリックした際に、
顧客情報の自動取得を行う機能を提供します。
"""

import win32gui
import win32con
import win32api
import logging
import time
import threading
import json
import os
from typing import Optional, Callable

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

    def start_monitoring(self):
        """監視を開始"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self.monitor_thread.start()
            logging.info("電話ボタン監視を開始しました")

    def stop_monitoring(self):
        """監視を停止"""
        self.is_monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=1.0)
        logging.info("電話ボタン監視を停止しました")

    def _monitor_loop(self):
        """監視ループ"""
        while self.is_monitoring:
            try:
                # 一定間隔でウィンドウを再検出
                current_time = time.time()
                if current_time - self.last_redetect_time > self.redetect_interval:
                    self.find_cti_window()
                    self.last_redetect_time = current_time

                # ウィンドウが見つかっていない場合はスキップ
                if not self.window_handle:
                    time.sleep(1.0)
                    continue

                # 電話ボタンの状態を確認
                if self._check_phone_button():
                    # クリック間隔を確認
                    if current_time - self.last_click_time > self.click_interval:
                        self.last_click_time = current_time
                        if self.delay_seconds > 0:
                            self._start_countdown()
                        else:
                            self._execute_callback()

            except Exception as e:
                logging.error(f"監視ループでエラーが発生: {str(e)}")
            
            time.sleep(0.1)  # CPU負荷軽減のため

    def _check_phone_button(self) -> bool:
        """電話ボタンの状態を確認"""
        try:
            # ボタンの検索と状態確認の処理
            # この部分は実際のCTIアプリケーションの仕様に合わせて実装
            return False
        except Exception as e:
            logging.error(f"電話ボタンの状態確認中にエラー: {str(e)}")
            return False

    def _start_countdown(self):
        """カウントダウンを開始"""
        if not self.is_counting_down:
            self.is_counting_down = True
            self.countdown_thread = threading.Thread(target=self._countdown_loop, daemon=True)
            self.countdown_thread.start()

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
                self._execute_callback()
            self.is_counting_down = False
            
        except Exception as e:
            logging.error(f"カウントダウン中にエラー: {str(e)}")
            self.is_counting_down = False

    def _execute_callback(self):
        """コールバック関数を実行"""
        try:
            if self.callback:
                self.callback()
        except Exception as e:
            logging.error(f"コールバック実行中にエラー: {str(e)}") 