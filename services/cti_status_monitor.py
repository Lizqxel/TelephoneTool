"""
CTI状態監視サービス

CtiOutboundSysの状態変化（「発信中」→「通話中」）を監視し、
自動で顧客情報取得+提供判定を実行する機能を提供します。

主な機能：
- CTI状態の常時監視
- 状態変化の検出（待ち受け中→発信中→通話中）
- 発信中から通話中への変化時の自動処理実行
- エラーハンドリングとログ出力

制限事項：
- 常時監視のためCPU負荷を考慮
- 重複実行防止機能付き
- 監視間隔は0.2秒
"""

import win32gui
import win32con
import win32api
import logging
import time
import threading
from typing import Optional, Callable, Dict, Any
from dataclasses import dataclass
from enum import Enum
import json
import os

class CTIStatus(Enum):
    """CTI状態の列挙型"""
    WAITING = "待ち受け中"
    DIALING = "発信中"
    TALKING = "通話中"
    UNKNOWN = "不明"

@dataclass
class StatusChangeEvent:
    """状態変化イベント"""
    previous_status: CTIStatus
    current_status: CTIStatus
    timestamp: float
    
class CTIStatusMonitor:
    """CTI状態監視クラス"""
    
    def __init__(self, on_dialing_to_talking: Callable[[], None]):
        """
        初期化
        
        Args:
            on_dialing_to_talking: 発信中→通話中の変化時に実行するコールバック関数
        """
        self.on_dialing_to_talking = on_dialing_to_talking
        self.window_handle = None
        self.status_text_handle = None
        self.is_monitoring = False
        self.monitor_thread = None
        self.current_status = CTIStatus.UNKNOWN
        self.previous_status = CTIStatus.UNKNOWN
        self.last_detection_time = 0
        self.detection_interval = 0.2  # 監視間隔（秒）
        self.window_redetect_interval = 5  # ウィンドウ再検出間隔（秒）
        self.last_window_redetect_time = 0
        self.is_processing = False  # 処理中フラグ（重複実行防止）
        self.settings_file = "settings.json"
        self.enable_auto_processing = True  # 自動処理の有効/無効
        
        # 設定の読み込み
        self.load_settings()
        
        logging.info("CTI状態監視サービスを初期化しました")
        
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.enable_auto_processing = settings.get('enable_auto_cti_processing', True)
                    self.detection_interval = settings.get('cti_monitor_interval', 0.2)
            else:
                self.enable_auto_processing = True
                self.detection_interval = 0.2
        except Exception as e:
            logging.error(f"CTI監視設定の読み込みに失敗しました: {str(e)}")
            self.enable_auto_processing = True
            self.detection_interval = 0.2
            
    def update_settings(self):
        """設定を更新する"""
        self.load_settings()
        
    def find_cti_window(self) -> bool:
        """CTIメインウィンドウを検索"""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    title = win32gui.GetWindowText(hwnd)
                    if "CTIメイン" in title:
                        self.window_handle = hwnd
                        logging.debug(f"CTIメインウィンドウを検出: handle={hwnd}, title='{title}'")
                        return False
                except Exception:
                    pass
            return True
            
        try:
            win32gui.EnumWindows(callback, None)
            return self.window_handle is not None
        except Exception as e:
            logging.error(f"CTIウィンドウの検索中にエラー: {str(e)}")
            return False
            
    def find_status_text_control(self) -> bool:
        """状態表示テキストコントロールを検索"""
        if not self.window_handle:
            return False
            
        status_controls = []
        
        def enum_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                    
                # コントロールのテキストを取得
                text = win32gui.GetWindowText(hwnd)
                if text and any(status in text for status in ["待ち受け中", "発信中", "通話中"]):
                    class_name = win32gui.GetClassName(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    status_controls.append({
                        'handle': hwnd,
                        'text': text,
                        'class': class_name,
                        'rect': rect
                    })
                    logging.debug(f"状態コントロールを検出: text='{text}', class='{class_name}', handle={hwnd}")
                    
            except Exception:
                pass
            return True
            
        try:
            win32gui.EnumChildWindows(self.window_handle, enum_callback, None)
            
            if status_controls:
                # 最初に見つかったコントロールを使用
                selected_control = status_controls[0]
                self.status_text_handle = selected_control['handle']
                logging.info(f"状態表示コントロールを選択: text='{selected_control['text']}', handle={self.status_text_handle}")
                return True
            else:
                logging.warning("状態表示コントロールが見つかりませんでした")
                return False
                
        except Exception as e:
            logging.error(f"状態表示コントロールの検索中にエラー: {str(e)}")
            return False
            
    def get_current_status(self) -> CTIStatus:
        """現在のCTI状態を取得"""
        if not self.status_text_handle:
            return CTIStatus.UNKNOWN
            
        try:
            text = win32gui.GetWindowText(self.status_text_handle)
            
            if "通話中" in text:
                return CTIStatus.TALKING
            elif "発信中" in text:
                return CTIStatus.DIALING
            elif "待ち受け中" in text:
                return CTIStatus.WAITING
            else:
                return CTIStatus.UNKNOWN
                
        except Exception as e:
            logging.error(f"CTI状態の取得中にエラー: {str(e)}")
            return CTIStatus.UNKNOWN
            
    def start_monitoring(self):
        """状態監視を開始"""
        if self.is_monitoring:
            logging.warning("CTI状態監視は既に開始されています")
            return
            
        self.is_monitoring = True
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()
        logging.info("CTI状態監視を開始しました")
        
    def stop_monitoring(self):
        """状態監視を停止"""
        if not self.is_monitoring:
            return
            
        self.is_monitoring = False
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=1.0)
        logging.info("CTI状態監視を停止しました")
        
    def _monitor_loop(self):
        """監視ループ"""
        while self.is_monitoring:
            try:
                current_time = time.time()
                
                # 一定間隔でウィンドウとコントロールを再検出
                if current_time - self.last_window_redetect_time > self.window_redetect_interval:
                    if not self.window_handle or not win32gui.IsWindow(self.window_handle):
                        self.find_cti_window()
                        
                    if self.window_handle and not self.status_text_handle:
                        self.find_status_text_control()
                        
                    self.last_window_redetect_time = current_time
                
                # 状態チェック間隔を制御
                if current_time - self.last_detection_time >= self.detection_interval:
                    self._check_status_change()
                    self.last_detection_time = current_time
                    
            except Exception as e:
                logging.error(f"CTI状態監視ループでエラーが発生: {str(e)}")
                
            # CPU負荷軽減のため短いスリープ
            time.sleep(0.05)
            
    def _check_status_change(self):
        """状態変化をチェック"""
        if not self.enable_auto_processing:
            return
            
        new_status = self.get_current_status()
        
        # 状態が変化した場合
        if new_status != self.current_status:
            self.previous_status = self.current_status
            self.current_status = new_status
            
            logging.info(f"CTI状態が変化: {self.previous_status.value} → {self.current_status.value}")
            
            # 発信中→通話中の変化を検出
            if (self.previous_status == CTIStatus.DIALING and 
                self.current_status == CTIStatus.TALKING):
                self._handle_dialing_to_talking()
                
    def _handle_dialing_to_talking(self):
        """発信中→通話中の変化時の処理"""
        if self.is_processing:
            logging.info("既に自動処理が実行中のため、重複実行をスキップします")
            return
            
        self.is_processing = True
        
        try:
            logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
            logging.info("自動処理を開始します: 顧客情報取得 → 提供判定検索")
            
            # コールバック関数を実行
            if self.on_dialing_to_talking:
                self.on_dialing_to_talking()
                
        except Exception as e:
            logging.error(f"発信中→通話中の自動処理中にエラーが発生: {str(e)}")
        finally:
            # 一定時間後に処理フラグをリセット（重複実行防止の解除）
            threading.Timer(5.0, self._reset_processing_flag).start()
            
    def _reset_processing_flag(self):
        """処理中フラグをリセット"""
        self.is_processing = False
        logging.debug("CTI自動処理フラグをリセットしました")
        
    def get_status_info(self) -> Dict[str, Any]:
        """現在の監視状態情報を取得"""
        return {
            'is_monitoring': self.is_monitoring,
            'current_status': self.current_status.value,
            'previous_status': self.previous_status.value,
            'window_found': self.window_handle is not None,
            'status_control_found': self.status_text_handle is not None,
            'enable_auto_processing': self.enable_auto_processing,
            'is_processing': self.is_processing
        }
        
    def set_auto_processing(self, enabled: bool):
        """自動処理の有効/無効を設定"""
        self.enable_auto_processing = enabled
        logging.info(f"CTI自動処理を{'有効' if enabled else '無効'}にしました") 