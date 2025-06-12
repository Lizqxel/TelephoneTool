"""
CTI状態監視モジュール

このモジュールは、CTIの状態を監視し、状態変化時にコールバック関数を呼び出します。
"""

import time
import logging
import threading
from typing import Callable, Optional
from services.oneclick import OneClickService
import win32gui
import win32con
from datetime import datetime, timedelta
from threading import Event


class CTIStatusMonitor:
    """CTI状態監視クラス"""
    
    def __init__(self,
                 on_dialing_to_talking_callback: Optional[Callable] = None,
                 on_call_ended_callback: Optional[Callable] = None,
                 on_talking_started_callback: Optional[Callable] = None):
        """
        CTI状態監視クラスを初期化
        
        Args:
            on_dialing_to_talking_callback (Callable, optional): 発信中→通話中の状態変化時のコールバック
            on_call_ended_callback (Callable, optional): 通話終了時のコールバック
            on_talking_started_callback (Callable, optional): 通話中状態開始時のコールバック
        """
        self.cti_service = OneClickService()
        self.on_dialing_to_talking_callback = on_dialing_to_talking_callback
        self.on_call_ended_callback = on_call_ended_callback
        self.on_talking_started_callback = on_talking_started_callback
        
        self.monitor_thread = None
        self.is_monitoring = False
        self.enable_auto_processing = True
        self.monitor_interval = 0.2
        
        # 前回の状態を保持
        self.previous_status = None
        
        # ログ出力の制御用
        self.last_log_time = datetime.now()
        self.log_interval = timedelta(minutes=1)  # ログ出力の間隔（1分）
    
    def start_monitoring(self):
        """CTI状態監視を開始"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self.monitor_thread.start()
            logging.info("CTI状態監視を開始しました")
    
    def stop_monitoring(self):
        """CTI状態監視を停止"""
        if self.is_monitoring:
            self.is_monitoring = False
            if self.monitor_thread:
                self.monitor_thread.join(timeout=1.0)
            logging.info("CTI状態監視を停止しました")
    
    def _monitor_loop(self):
        """CTI状態監視ループ"""
        while self.is_monitoring:
            try:
                # CTIの状態を取得
                current_status = self.cti_service.get_status()
                
                if current_status and self.previous_status:
                    # 発信中→通話中の状態変化を検出
                    if (self.previous_status == "dialing" and
                        current_status == "talking" and
                        self.on_dialing_to_talking_callback and
                        self.enable_auto_processing):
                        self.on_dialing_to_talking_callback()
                    
                    # 通話中→待ち受け中（通話終了）の状態変化を検出
                    elif (self.previous_status == "talking" and
                          current_status == "waiting" and
                          self.on_call_ended_callback):
                        self.on_call_ended_callback()
                
                # 通話中状態の開始を検出
                if (current_status == "talking" and
                    self.previous_status != "talking" and
                    self.on_talking_started_callback):
                    self.on_talking_started_callback()
                
                # 状態を更新
                self.previous_status = current_status
                
            except Exception as e:
                logging.error(f"CTI状態監視中にエラー: {str(e)}")
            
            # 監視間隔分スリープ
            time.sleep(self.monitor_interval)

    def find_cti_window(self) -> bool:
        """CTIメインウィンドウを検索"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if "CTIメイン" in title:
                        self.window_handle = hwnd
                        
                        # ウィンドウの詳細情報をログ出力
                        class_name = win32gui.GetClassName(hwnd)
                        is_visible = win32gui.IsWindowVisible(hwnd)
                        is_enabled = win32gui.IsWindowEnabled(hwnd)
                        
                        logging.info(f"CTIメインウィンドウを検出: handle={hwnd}")
                        logging.info(f"CTIメインウィンドウの詳細:")
                        logging.info(f"- テキスト: {title}")
                        logging.info(f"- クラス名: {class_name}")
                        logging.info(f"- 可視状態: {is_visible}")
                        logging.info(f"- 有効状態: {is_enabled}")
                        
                        return False
                except Exception:
                    pass
            return True
            
        try:
            # 既存のハンドルをクリア
            self.window_handle = None
            win32gui.EnumWindows(callback, None)
            
            if self.window_handle:
                logging.info(f"CTIメインウィンドウの検出に成功しました")
                return True
            else:
                # 前回のログ出力から指定時間が経過している場合のみログを出力
                now = datetime.now()
                if now - self.last_log_time >= self.log_interval:
                    logging.debug("CTIメインウィンドウが見つかりませんでした")
                    self.last_log_time = now
                return False
                
        except Exception as e:
            logging.error(f"CTIウィンドウの検索中にエラー: {str(e)}")
            return False

    def _check_status_change(self):
        """
        CTI状態の変化をチェック
        """
        try:
            # ウィンドウハンドルの確認と再検出
            if not self.window_handle or not win32gui.IsWindow(self.window_handle):
                # ウィンドウが無効になった場合、再検出を試行
                if self.find_cti_window():
                    logging.info("CTIウィンドウを再検出しました")
                else:
                    # 前回のログ出力から指定時間が経過している場合のみログを出力
                    now = datetime.now()
                    if now - self.last_log_time >= self.log_interval:
                        logging.debug("CTIウィンドウが見つかりません")
                        self.last_log_time = now
                    return
                    
            # 状態表示コントロールの確認と再検出
            if not self.status_text_handle:
                if self.find_status_text_control():
                    logging.info("CTI状態表示コントロールを再検出しました")
                else:
                    # 前回のログ出力から指定時間が経過している場合のみログを出力
                    now = datetime.now()
                    if now - self.last_log_time >= self.log_interval:
                        logging.debug("CTI状態表示コントロールが見つかりません")
                        self.last_log_time = now
                    return

            # 状態を更新
            self.previous_status = current_status
            
        except Exception as e:
            logging.error(f"CTI状態変化チェック中にエラー: {str(e)}") 