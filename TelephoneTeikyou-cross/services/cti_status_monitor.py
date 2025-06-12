"""
CTI状態監視モジュール

このモジュールは、CTIの状態を監視し、状態変化時にコールバック関数を呼び出します。
"""

import time
import logging
import threading
from typing import Callable, Optional
from services.oneclick import OneClickService


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