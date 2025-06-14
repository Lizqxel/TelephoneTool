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
                 on_talking_started_callback: Optional[Callable] = None,
                 on_cancel_processing_callback: Optional[Callable] = None):
        """
        CTI状態監視クラスを初期化
        
        Args:
            on_dialing_to_talking_callback (Callable, optional): 発信中→通話中の状態変化時のコールバック
            on_call_ended_callback (Callable, optional): 通話終了時のコールバック
            on_talking_started_callback (Callable, optional): 通話中状態開始時のコールバック
            on_cancel_processing_callback (Callable, optional): 処理キャンセル時のコールバック
        """
        self.cti_service = OneClickService()
        self.on_dialing_to_talking_callback = on_dialing_to_talking_callback
        self.on_call_ended_callback = on_call_ended_callback
        self.on_talking_started_callback = on_talking_started_callback
        self.on_cancel_processing_callback = on_cancel_processing_callback
        
        self.monitor_thread = None
        self.is_monitoring = False
        self.enable_auto_processing = True
        self.monitor_interval = 0.5  # 0.2秒から0.5秒に変更（CPU負荷軽減）
        
        # 初期化ログ
        logging.info("★★★ CTIStatusMonitor 初期化完了 ★★★")
        logging.info(f"- 発信中→通話中コールバック: {self.on_dialing_to_talking_callback is not None}")
        logging.info(f"- 自動処理有効: {self.enable_auto_processing}")
        logging.info(f"- 監視間隔: {self.monitor_interval}秒")
        
        # 前回の状態を保持
        self.previous_status = None
        
        # ログ出力の制御用
        self.last_log_time = datetime.now()
        self.log_interval = timedelta(minutes=1)  # ログ出力の間隔（1分）
        
        # アクションボタン監視用
        self.action_buttons = ["NG", "次", "留守", "担当者不在"]
        self.previous_button_states = {}
        
        # 通話時間監視用
        self.call_start_time = None
        self.call_duration_threshold = 0  # デフォルト値（設定で変更可能）
    
    def set_call_duration_threshold(self, threshold_seconds):
        """
        通話時間閾値を設定
        
        Args:
            threshold_seconds (int): 閾値（秒）
        """
        self.call_duration_threshold = threshold_seconds
        logging.info(f"通話時間閾値を{threshold_seconds}秒に設定しました")
    
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
                    if (self.previous_status == "dialing" and current_status == "talking"):
                        logging.info(f"★★★ 発信中→通話中の状態変化を検出 ★★★")
                        logging.info(f"- コールバック設定: {self.on_dialing_to_talking_callback is not None}")
                        logging.info(f"- 自動処理有効: {self.enable_auto_processing}")
                        logging.info(f"- 通話時間閾値: {self.call_duration_threshold}秒")
                        
                        if (self.on_dialing_to_talking_callback and self.enable_auto_processing):
                            # 通話開始時間を記録
                            self.call_start_time = datetime.now()
                            
                            # 通話時間閾値をチェック
                            if self.call_duration_threshold > 0:
                                logging.info(f"★★★ {self.call_duration_threshold}秒後に自動処理を実行予定 ★★★")
                                # 指定秒数待機してから自動処理を実行
                                threading.Timer(
                                    self.call_duration_threshold,
                                    self._check_and_execute_auto_processing
                                ).start()
                            else:
                                logging.info("★★★ 即座に自動処理を実行します ★★★")
                                # 即座に自動処理を実行
                                self.on_dialing_to_talking_callback()
                        else:
                            if not self.on_dialing_to_talking_callback:
                                logging.warning("発信中→通話中コールバックが設定されていません")
                            if not self.enable_auto_processing:
                                logging.warning("自動処理が無効になっています")
                    
                    # 通話中→待ち受け中（通話終了）の状態変化を検出
                    elif (self.previous_status == "talking" and
                          current_status == "waiting" and
                          self.on_call_ended_callback):
                        self.call_start_time = None  # 通話時間をリセット
                        self.on_call_ended_callback()
                
                # 通話中状態の開始を検出
                if (current_status == "talking" and
                    self.previous_status != "talking" and
                    self.on_talking_started_callback):
                    self.on_talking_started_callback()
                
                # アクションボタンの状態をチェック
                self._check_action_buttons()
                
                # 状態を更新
                self.previous_status = current_status
                
            except Exception as e:
                logging.error(f"CTI状態監視中にエラー: {str(e)}")
            
            # 監視間隔分スリープ
            time.sleep(self.monitor_interval)
    
    def _check_and_execute_auto_processing(self):
        """
        通話時間閾値後の自動処理実行チェック
        """
        try:
            # 現在も通話中かチェック
            current_status = self.cti_service.get_status()
            if current_status == "talking" and self.call_start_time:
                # 通話時間を計算
                call_duration = (datetime.now() - self.call_start_time).total_seconds()
                if call_duration >= self.call_duration_threshold:
                    logging.info(f"★★★ 通話時間{call_duration:.1f}秒が閾値{self.call_duration_threshold}秒を超えたため自動処理を実行 ★★★")
                    if self.on_dialing_to_talking_callback:
                        self.on_dialing_to_talking_callback()
                else:
                    logging.info(f"通話時間{call_duration:.1f}秒が閾値{self.call_duration_threshold}秒未満のため自動処理をスキップ")
            else:
                logging.info("通話が既に終了しているため自動処理をスキップ")
        except Exception as e:
            logging.error(f"自動処理実行チェック中にエラー: {str(e)}")
    
    def _check_action_buttons(self):
        """
        アクションボタンの状態をチェックし、クリックを検出
        """
        try:
            for button_name in self.action_buttons:
                current_state = self._get_button_state(button_name)
                previous_state = self.previous_button_states.get(button_name, False)
                
                # ボタンがクリックされた（False→True）を検出
                if current_state and not previous_state:
                    logging.info(f"★★★ 「{button_name}」ボタンクリックを検出 ★★★")
                    if self.on_cancel_processing_callback:
                        logging.info(f"★★★ 「{button_name}」ボタンクリックによる処理キャンセル要求を受信 ★★★")
                        self.on_cancel_processing_callback()
                    else:
                        logging.warning("MainWindowのキャンセルコールバックが設定されていません")
                
                # 状態を更新
                self.previous_button_states[button_name] = current_state
                
        except Exception as e:
            logging.error(f"アクションボタン監視中にエラー: {str(e)}")
    
    def _get_button_state(self, button_name):
        """
        指定されたボタンの状態を取得
        
        Args:
            button_name (str): ボタン名
            
        Returns:
            bool: ボタンが押されている状態ならTrue
        """
        try:
            # CTIサービスからボタン状態を取得
            # 実際の実装では、CTIシステムのAPIやウィンドウ監視を使用
            return self.cti_service.is_button_pressed(button_name)
        except Exception as e:
            logging.debug(f"ボタン状態取得中にエラー ({button_name}): {str(e)}")
            return False

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

            # 現在の状態を取得
            current_status = self.get_current_status()
            
            # 状態が変化した場合の処理
            if current_status != self.previous_status:
                logging.info(f"CTI状態の変化を検出: {self.previous_status} → {current_status}")
                
                # 発信中→通話中の遷移を検出
                if self.previous_status == "dialing" and current_status == "talking":
                    if self.on_dialing_to_talking_callback:
                        self.on_dialing_to_talking_callback()
                
                # 通話開始を検出
                if current_status == "talking" and self.previous_status != "talking":
                    if self.on_talking_started_callback:
                        self.on_talking_started_callback()
                
                # 通話終了を検出
                if self.previous_status == "talking" and current_status != "talking":
                    if self.on_call_ended_callback:
                        self.on_call_ended_callback()
                
                # 状態を更新
                self.previous_status = current_status
            
        except Exception as e:
            logging.error(f"CTI状態変化チェック中にエラー: {str(e)}") 

    def get_current_status(self) -> str:
        """
        現在のCTI状態を取得
        
        Returns:
            str: CTIの状態（"waiting", "dialing", "talking"のいずれか）
        """
        try:
            if not self.status_text_handle:
                return ""
            
            # ウィンドウテキストを取得
            text = win32gui.GetWindowText(self.status_text_handle)
            
            # 状態を判定
            if "発信中" in text:
                return "dialing"
            elif "通話中" in text:
                return "talking"
            else:
                return "waiting"
                
        except Exception as e:
            logging.error(f"CTI状態の取得中にエラー: {str(e)}")
            return ""

    def find_status_text_control(self) -> bool:
        """
        状態表示コントロールを検索
        
        Returns:
            bool: コントロールが見つかった場合True
        """
        try:
            def callback(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        text = win32gui.GetWindowText(hwnd)
                        if any(status in text for status in ["待機中", "発信中", "通話中"]):
                            self.status_text_handle = hwnd
                            logging.info(f"状態表示コントロールを検出: handle={hwnd}, text='{text}'")
                            return False
                    except Exception:
                        pass
                return True
            
            if self.window_handle:
                self.status_text_handle = None
                win32gui.EnumChildWindows(self.window_handle, callback, None)
                
                if self.status_text_handle:
                    return True
                else:
                    # 前回のログ出力から指定時間が経過している場合のみログを出力
                    now = datetime.now()
                    if now - self.last_log_time >= self.log_interval:
                        logging.debug("状態表示コントロールが見つかりません")
                        self.last_log_time = now
                    return False
            else:
                return False
                
        except Exception as e:
            logging.error(f"状態表示コントロールの検索中にエラー: {str(e)}")
            return False 