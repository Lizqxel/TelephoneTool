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
import win32api
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
        
        # アクションボタン監視用（TelephoneToolと同じ詳細実装）
        self.next_button_handle = None  # 「次」ボタンのハンドル
        self.rusu_button_handle = None  # 「留守」ボタンのハンドル
        self.tantou_fuzai_button_handle = None  # 「担当者不在」ボタンのハンドル
        self.ng_button_handle = None  # 「NG」ボタンのハンドル
        self.last_button_click_time = 0
        self.button_click_interval = 0.5  # ボタンクリック間隔（秒）
        self.buttons_detected = False
        
        # 処理制御用
        import threading
        self.processing_lock = threading.Lock()
        self.processing_thread = None
        self.talking_start_time = 0
        
        # 通話時間監視用
        self.call_start_time = None
        self.call_duration_threshold = 0  # デフォルト値（設定で変更可能）
        
        # ウィンドウハンドル関連の初期化
        self.window_handle = None
        self.status_text_handle = None
    
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
                # CTI状態変化をチェック（_check_status_changeメソッドを使用）
                self._check_status_change()
                
                # アクションボタンのクリックをチェック（TelephoneToolと同じ実装）
                self._check_action_button_click()
                
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
            current_status = self.get_current_status()
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
    
    def _check_action_button_click(self):
        """
        アクションボタン（「次」「留守」「担当者不在」「NG」）のクリックをチェック
        TelephoneToolと同じ詳細実装
        """
        try:
            # ボタンの再検出が必要な場合
            if not any([
                self.next_button_handle,
                self.rusu_button_handle,
                self.tantou_fuzai_button_handle,
                self.ng_button_handle
            ]):
                if not self.find_action_buttons():
                    return
            
            # マウスクリックをチェック
            import win32api
            import win32con
            if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                current_time = time.time()
                # 連続クリックを防ぐ
                if current_time - self.last_button_click_time >= self.button_click_interval:
                    # マウス座標を取得
                    x, y = win32api.GetCursorPos()
                    
                    # 各ボタンの位置をチェック
                    clicked_button = None
                    button_rect = None
                    
                    import win32gui
                    if self.next_button_handle:
                        rect = win32gui.GetWindowRect(self.next_button_handle)
                        if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                            clicked_button = "次"
                            button_rect = rect
                            
                    if self.rusu_button_handle:
                        rect = win32gui.GetWindowRect(self.rusu_button_handle)
                        if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                            clicked_button = "留守"
                            button_rect = rect
                            
                    if self.tantou_fuzai_button_handle:
                        rect = win32gui.GetWindowRect(self.tantou_fuzai_button_handle)
                        if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                            clicked_button = "担当者不在"
                            button_rect = rect
                            
                    if self.ng_button_handle:
                        rect = win32gui.GetWindowRect(self.ng_button_handle)
                        if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                            clicked_button = "NG"
                            button_rect = rect
                    
                    if clicked_button:
                        # 提供判定の実行状態を確認
                        is_processing = False
                        with self.processing_lock:
                            is_processing = self.is_processing or (self.processing_thread and self.processing_thread.is_alive())
                        
                        logging.info(f"★★★ 「{clicked_button}」ボタンがクリックされました ★★★")
                        logging.info(f"- クリック時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                        logging.info(f"- マウス座標: x={x}, y={y}")
                        logging.info(f"- ボタン位置: left={button_rect[0]}, top={button_rect[1]}, right={button_rect[2]}, bottom={button_rect[3]}")
                        logging.info(f"- 提供判定状態: {'実行中' if is_processing else '未実行'}")
                        
                        self.last_button_click_time = current_time
                        
                        # 提供判定をキャンセル
                        self._cancel_processing(clicked_button)
                        
        except Exception as e:
            logging.error(f"アクションボタンクリックの検出中にエラー: {str(e)}")

    def _cancel_processing(self, button_name: str):
        """
        提供判定をキャンセルし、フラグをリセット
        TelephoneToolと同じ詳細実装
        
        Args:
            button_name: クリックされたボタンの名前
        """
        try:
            with self.processing_lock:
                logging.info(f"★★★ 「{button_name}」ボタンクリックによる処理キャンセルを開始 ★★★")
                logging.info(f"- キャンセル開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # 現在の処理状態を記録
                was_processing = self.is_processing
                if self.talking_start_time > 0:
                    elapsed_time = time.time() - self.talking_start_time
                    logging.info(f"- 処理開始からの経過時間: {elapsed_time:.1f}秒")
                
                # 1. MainWindowの検索処理をキャンセル（最優先で実行）
                if self.on_cancel_processing_callback:
                    try:
                        logging.info("★★★ MainWindowの検索処理キャンセルを実行中... ★★★")
                        logging.info(f"- コールバック関数: {self.on_cancel_processing_callback}")
                        self.on_cancel_processing_callback(button_name)
                        logging.info("★★★ MainWindowの検索処理キャンセルが完了しました ★★★")
                    except Exception as callback_error:
                        logging.error(f"検索処理キャンセルコールバックの実行中にエラー: {str(callback_error)}")
                        logging.error(f"エラー詳細: {type(callback_error).__name__}: {str(callback_error)}")
                        import traceback
                        logging.error(f"スタックトレース: {traceback.format_exc()}")
                else:
                    logging.warning("★★★ MainWindowのキャンセルコールバックが設定されていません ★★★")
                    logging.warning(f"- on_cancel_processing_callback: {self.on_cancel_processing_callback}")
                
                # 2. CTI監視の処理中フラグを確実にリセット
                self.is_processing = False
                self.talking_start_time = 0
                logging.info("- CTI監視の処理中フラグをリセットしました")
                
                # 結果をログ出力
                if was_processing:
                    logging.info(f"★★★ 「{button_name}」ボタンクリックによる処理キャンセルが完了しました ★★★")
                    logging.info("- 実行中の処理が正常にキャンセルされました")
                else:
                    logging.info(f"★★★ 「{button_name}」ボタンクリックによるフラグリセットが完了しました ★★★")
                    logging.info("- 処理は実行中ではありませんでしたが、フラグをリセットしました")
                    
        except Exception as e:
            logging.error(f"処理キャンセル中にエラーが発生: {str(e)}")
            logging.error(f"エラー詳細: {type(e).__name__}: {str(e)}")
            import traceback
            logging.error(f"スタックトレース: {traceback.format_exc()}")
            
            # エラー時も確実にフラグをリセット
            try:
                self.is_processing = False
                self.talking_start_time = 0
                logging.info("- エラー発生時もフラグをリセットしました")
                
                # エラー時でもMainWindowのキャンセル処理を試行
                if self.on_cancel_processing_callback:
                    logging.info("- エラー時にMainWindowキャンセル処理を再試行します")
                    self.on_cancel_processing_callback(button_name)
                    logging.info("- エラー時のMainWindowキャンセル処理が完了しました")
            except Exception as reset_error:
                logging.error(f"エラー時のフラグリセット処理でもエラー: {str(reset_error)}")
                logging.error(f"リセットエラー詳細: {type(reset_error).__name__}: {str(reset_error)}")

    def find_action_buttons(self) -> bool:
        """
        アクションボタン（「次」「留守」「担当者不在」「NG」）を検索
        TelephoneToolと同じ詳細実装
        
        Returns:
            bool: いずれかのボタンが見つかった場合True
        """
        try:
            import win32gui
            def callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        text = win32gui.GetWindowText(hwnd)
                        if text == "次":
                            self.next_button_handle = hwnd
                            rect = win32gui.GetWindowRect(hwnd)
                            if not self.buttons_detected:
                                logging.info(f"「次」ボタンを検出: handle={hwnd}, rect={rect}")
                        elif text == "留守":
                            self.rusu_button_handle = hwnd
                            rect = win32gui.GetWindowRect(hwnd)
                            if not self.buttons_detected:
                                logging.info(f"「留守」ボタンを検出: handle={hwnd}, rect={rect}")
                        elif text == "担当者不在":
                            self.tantou_fuzai_button_handle = hwnd
                            rect = win32gui.GetWindowRect(hwnd)
                            if not self.buttons_detected:
                                logging.info(f"「担当者不在」ボタンを検出: handle={hwnd}, rect={rect}")
                        elif text == "NG":
                            self.ng_button_handle = hwnd
                            rect = win32gui.GetWindowRect(hwnd)
                            if not self.buttons_detected:
                                logging.info(f"「NG」ボタンを検出: handle={hwnd}, rect={rect}")
                    except Exception:
                        pass
                return True
            
            # 既存のハンドルをクリア
            self.next_button_handle = None
            self.rusu_button_handle = None
            self.tantou_fuzai_button_handle = None
            self.ng_button_handle = None
            
            # CTIメインウィンドウ内の子ウィンドウを列挙
            if self.window_handle:
                win32gui.EnumChildWindows(self.window_handle, callback, None)
            
            # いずれかのボタンが見つかったかチェック
            found = any([
                self.next_button_handle,
                self.rusu_button_handle,
                self.tantou_fuzai_button_handle,
                self.ng_button_handle
            ])
            
            if found and not self.buttons_detected:
                logging.info("アクションボタンの検出に成功しました")
                self.buttons_detected = True
                return True
            elif not found:
                self.buttons_detected = False
                logging.debug("アクションボタンが見つかりませんでした")
                return False
            else:
                return True
                
        except Exception as e:
            logging.error(f"アクションボタンの検索中にエラー: {str(e)}")
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
                
                # 通話開始を検出
                if current_status == "talking" and self.previous_status != "talking":
                    if self.on_talking_started_callback:
                        self.on_talking_started_callback()
                
                # 通話終了を検出
                if self.previous_status == "talking" and current_status != "talking":
                    self.call_start_time = None  # 通話時間をリセット
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