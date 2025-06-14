"""
CTI状態監視サービス

CtiOutboundSysの状態変化（「発信中」→「通話中」）を監視し、
自動で顧客情報取得+提供判定を実行する機能を提供します。

主な機能：
- CTI状態の常時監視
- 状態変化の検出（待ち受け中→発信中→通話中）
- 発信中から通話中への変化時の自動処理実行
- 通話中状態への遷移時のイベント通知
- 通話終了時（通話中→待ち受け中）のイベント通知
- 「次」「留守」「担当者不在」「NG」ボタンクリックの検出と提供判定のキャンセル
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
    
    def __init__(self, on_dialing_to_talking_callback: Optional[Callable] = None,
                 on_call_ended_callback: Optional[Callable] = None,
                 on_talking_started_callback: Optional[Callable] = None):
        """
        初期化
        
        Args:
            on_dialing_to_talking_callback: 発信中→通話中の状態変化時のコールバック関数
            on_call_ended_callback: 通話終了時（通話中→待ち受け中）のコールバック関数
            on_talking_started_callback: 通話中状態開始時のコールバック関数
        """
        self.on_dialing_to_talking_callback = on_dialing_to_talking_callback
        self.on_call_ended_callback = on_call_ended_callback
        self.on_talking_started_callback = on_talking_started_callback
        
        # 状態管理
        self.current_status = CTIStatus.UNKNOWN
        self.previous_status = CTIStatus.UNKNOWN
        
        # ウィンドウハンドル
        self.window_handle = None  # CTIメインウィンドウのハンドル
        self.status_text_handle = None  # 状態表示コントロールのハンドル
        self.next_button_handle = None  # 「次」ボタンのハンドル
        self.rusu_button_handle = None  # 「留守」ボタンのハンドル
        self.tantou_fuzai_button_handle = None  # 「担当者不在」ボタンのハンドル
        self.ng_button_handle = None  # 「NG」ボタンのハンドル
        
        # 監視制御
        self.is_monitoring = False
        self.monitor_thread = None
        self.monitor_interval = 0.2  # 監視間隔（秒）
        
        # 重複実行防止用
        self.is_processing = False
        self.last_status_change_time = 0
        self.status_change_cooldown = 2.0  # 同じ状態変化の検出間隔（秒）
        self.last_dialing_to_talking_time = 0  # 発信中→通話中の最後の検出時刻
        
        # 監視制御
        self.last_detection_time = 0
        self.window_redetect_interval = 5  # ウィンドウ再検出間隔（秒）
        self.last_window_redetect_time = 0
        self.enable_auto_processing = True  # 自動処理の有効/無効
        
        # 通話時間設定
        self.call_duration_threshold = 0  # 通話時間の閾値（秒）
        self.talking_start_time = 0  # 通話開始時刻
        
        # ボタンクリック検出用
        self.last_button_click_time = 0  # 最後のボタンクリック時刻
        self.button_click_interval = 0.5  # ボタンクリック間隔（秒）
        
        # ボタン検出状態
        self.buttons_detected = False  # ボタンが検出済みかどうか
        
        # 提供判定実行用
        self.processing_thread = None  # 提供判定実行用スレッド
        self.processing_lock = threading.Lock()  # 提供判定実行用ロック
        
        # 設定の読み込み
        self.load_settings()
        
        logging.info("CTI状態監視サービスを初期化しました")
        
    def load_settings(self):
        """設定ファイルから設定を読み込む"""
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    self.enable_auto_processing = settings.get('enable_auto_cti_processing', True)
                    self.monitor_interval = settings.get('cti_monitor_interval', 0.2)
                    self.call_duration_threshold = settings.get('call_duration_threshold', 0)
            else:
                self.enable_auto_processing = True
                self.monitor_interval = 0.2
        except Exception as e:
            logging.error(f"CTI監視設定の読み込みに失敗しました: {str(e)}")
            self.enable_auto_processing = True
            self.monitor_interval = 0.2
            
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
                logging.debug("CTIメインウィンドウが見つかりませんでした")
                return False
                
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
        """監視を開始"""
        if not self.is_monitoring:
            self.is_monitoring = True
            self.monitor_thread = threading.Thread(target=self._monitor_loop, name="CTIMonitorThread")
            self.monitor_thread.daemon = True  # デーモンスレッドとして設定
            self.monitor_thread.start()
            logging.info("CTI状態監視を開始しました")

    def stop_monitoring(self):
        """監視を停止"""
        if self.is_monitoring:
            self.is_monitoring = False
            if self.monitor_thread:
                self.monitor_thread.join(timeout=1.0)  # 1秒待機
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
                        
                    if self.window_handle:
                        self.find_action_buttons()
                        
                    self.last_window_redetect_time = current_time
                
                # 状態チェック間隔を制御
                if current_time - self.last_detection_time >= self.monitor_interval:
                    self._check_status_change()
                    self._check_action_button_click()  # アクションボタンクリックをチェック
                    self.last_detection_time = current_time
                    
            except Exception as e:
                logging.error(f"CTI状態監視ループでエラーが発生: {str(e)}")
                
            # CPU負荷軽減のため短いスリープ
            time.sleep(0.05)
            
    def _check_status_change(self):
        """
        CTI状態の変化を検出し、適切なコールバックを呼び出す
        """
        try:
            # 現在のCTI状態を取得
            current_status = self.get_current_status()
            
            # 状態が変化した場合
            if current_status != self.current_status:
                logging.info(f"CTI状態が変化: {self.current_status} → {current_status}")
                
                # 発信中→通話中への変化を検出
                if self.current_status == CTIStatus.DIALING and current_status == CTIStatus.TALKING:
                    logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
                    self.talking_start_time = time.time()
                    if self.on_dialing_to_talking_callback:
                        self.on_dialing_to_talking_callback()
                
                # 通話中→待ち受け中への変化を検出（通話終了）
                elif self.current_status == CTIStatus.TALKING and current_status == CTIStatus.WAITING:
                    logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
                    # 提供判定が実行中でない場合のみフラグをリセット
                    if not self.is_processing:
                        logging.info("通話終了により処理中フラグをリセットしました")
                        self.is_processing = False
                    if self.on_call_ended_callback:
                        self.on_call_ended_callback()
                    logging.info("★★★ 通話終了を検出: 2秒後に電話ボタン監視を再開します ★★★")
                    time.sleep(2)  # 2秒待機
                    self.find_action_buttons()
                
                # 状態を更新
                self.current_status = current_status
                
        except Exception as e:
            logging.error(f"CTI状態変化検出中にエラー: {str(e)}")
    
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
            if self.on_dialing_to_talking_callback:
                self.on_dialing_to_talking_callback()
                
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
        
    def _detect_status_change(self):
        """
        CTI状態の変化を検出し、適切なコールバックを呼び出す
        """
        try:
            # 現在のCTI状態を取得
            current_status = self.get_current_status()
            
            # 状態が変化した場合
            if current_status != self.current_status:
                logging.info(f"CTI状態が変化: {self.current_status} → {current_status}")
                
                # 発信中→通話中への変化を検出
                if self.current_status == CTIStatus.DIALING and current_status == CTIStatus.TALKING:
                    logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
                    self.talking_start_time = time.time()
                    if self.on_dialing_to_talking_callback:
                        self.on_dialing_to_talking_callback()
                
                # 通話中→待ち受け中への変化を検出（通話終了）
                elif self.current_status == CTIStatus.TALKING and current_status == CTIStatus.WAITING:
                    logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
                    # 提供判定が実行中でない場合のみフラグをリセット
                    if not self.is_processing:
                        logging.info("通話終了により処理中フラグをリセットしました")
                        self.is_processing = False
                    if self.on_call_ended_callback:
                        self.on_call_ended_callback()
                    logging.info("★★★ 通話終了を検出: 2秒後に電話ボタン監視を再開します ★★★")
                    time.sleep(2)  # 2秒待機
                    self.find_action_buttons()
                
                # 状態を更新
                self.current_status = current_status
                
        except Exception as e:
            logging.error(f"CTI状態変化検出中にエラー: {str(e)}")
    
    def _should_trigger_auto_processing(self, current_time: float) -> bool:
        """
        自動処理を実行すべきかどうかを判定
        
        Args:
            current_time: 現在時刻
            
        Returns:
            bool: 実行すべきならTrue
        """
        # 自動処理が無効の場合
        if not self.enable_auto_processing:
            logging.debug("自動処理が無効に設定されています")
            return False
            
        # 現在処理中の場合
        if self.is_processing:
            logging.debug("既に自動処理が実行中です")
            return False
            
        # 前回の発信中→通話中検出から短時間の場合
        if (current_time - self.last_dialing_to_talking_time) < self.status_change_cooldown:
            time_since_last = current_time - self.last_dialing_to_talking_time
            logging.debug(f"前回の自動処理から{time_since_last:.2f}秒しか経過していません（最小間隔: {self.status_change_cooldown}秒）")
            return False
            
        return True 

    def _trigger_auto_processing(self):
        """
        自動処理を実行する
        """
        try:
            # 処理中フラグを設定
            self.is_processing = True
            
            if not self.enable_auto_processing:
                logging.info("自動処理が無効に設定されているため、処理をスキップします")
                return
                
            if self.on_dialing_to_talking_callback is None:
                logging.warning("コールバック関数が設定されていません")
                return
                
            # コールバック関数を実行
            if self.on_dialing_to_talking_callback:
                self.on_dialing_to_talking_callback()
                
        except Exception as e:
            logging.error(f"自動処理の実行中にエラーが発生: {str(e)}")
        finally:
            # 処理完了後、一定時間後にフラグをリセット（通話終了で自動リセットされるため、バックアップとして）
            def reset_processing_flag():
                try:
                    # 通話中でない場合のみリセット（通話終了で既にリセットされている可能性があるため）
                    if self.current_status != CTIStatus.TALKING:
                        self.is_processing = False
                        logging.debug("バックアップ処理: 処理中フラグをリセットしました")
                    else:
                        logging.debug("通話中のため、処理中フラグのバックアップリセットをスキップしました")
                except Exception as e:
                    logging.error(f"処理中フラグのリセット中にエラー: {str(e)}")
                    # エラーが発生した場合は強制的にリセット
                    self.is_processing = False
                    
            timer = threading.Timer(10.0, reset_processing_flag)  # 10秒後にバックアップリセット
            timer.daemon = True
            timer.start() 

    def _check_and_trigger_auto_processing(self):
        """
        通話時間をチェックし、条件を満たす場合に自動処理を実行
        """
        try:
            with self.processing_lock:
                if not self.is_processing and self.talking_start_time > 0:
                    elapsed_time = time.time() - self.talking_start_time
                    
                    # 通話時間が閾値を超えた場合
                    if elapsed_time >= self.call_duration_threshold:
                        logging.info(f"★★★ 通話時間が閾値を超えました（{elapsed_time:.1f}秒）★★★")
                        logging.info(f"- 閾値: {self.call_duration_threshold}秒")
                        logging.info(f"- 通話開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.talking_start_time))}")
                        
                        # 提供判定を別スレッドで実行
                        self._execute_auto_processing_in_thread()
                        
        except Exception as e:
            logging.error(f"通話時間チェック中にエラーが発生: {str(e)}")

    def _execute_auto_processing_in_thread(self):
        """
        提供判定を別スレッドで実行
        """
        try:
            with self.processing_lock:
                if self.processing_thread and self.processing_thread.is_alive():
                    logging.info("既に提供判定が実行中です")
                    return
                    
                self.is_processing = True
                logging.info("★★★ 提供判定を開始します ★★★")
                logging.info(f"- 開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                logging.info(f"- 現在のCTI状態: {self.current_status.value}")
                
                self.processing_thread = threading.Thread(
                    target=self._execute_auto_processing,
                    name="AutoProcessingThread"
                )
                self.processing_thread.daemon = True  # デーモンスレッドとして設定
                self.processing_thread.start()
                
        except Exception as e:
            logging.error(f"提供判定スレッドの起動中にエラー: {str(e)}")
            with self.processing_lock:
                self.is_processing = False

    def _execute_auto_processing(self):
        """
        提供判定を実行
        """
        try:
            # 提供判定の実行
            if self.on_dialing_to_talking_callback:
                self.on_dialing_to_talking_callback()
                
            logging.info("★★★ 提供判定が完了しました ★★★")
            logging.info(f"- 完了時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            
        except Exception as e:
            logging.error(f"提供判定の実行中にエラー: {str(e)}")
        finally:
            with self.processing_lock:
                self.is_processing = False
                logging.info("提供判定の実行状態をリセットしました")

    def _check_action_button_click(self):
        """
        アクションボタン（次、留守、担当者不在、NG）のクリックを検出
        """
        try:
            # マウスの現在位置を取得
            x, y = win32api.GetCursorPos()
            
            # 各ボタンの位置を確認
            for button_name, button_info in self.action_buttons.items():
                if button_info['handle']:
                    rect = button_info['rect']
                    # マウスがボタンの領域内にあるか確認
                    if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                        # クリックを検出
                        logging.info(f"★★★ 「{button_name}」ボタンがクリックされました ★★★")
                        logging.info(f"- クリック時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                        logging.info(f"- マウス座標: x={x}, y={y}")
                        logging.info(f"- ボタン位置: left={rect[0]}, top={rect[1]}, right={rect[2]}, bottom={rect[3]}")
                        logging.info(f"- 現在のCTI状態: {self.current_status}")
                        logging.info(f"- 提供判定状態: {'実行中' if self.is_processing else '未実行'}")
                        
                        # 提供判定が実行中の場合、キャンセル処理を実行
                        if self.is_processing:
                            self._cancel_processing(button_name)
                        return True
            return False
            
        except Exception as e:
            logging.error(f"アクションボタンクリック検出中にエラー: {str(e)}")
            return False

    def _cancel_processing(self, button_name):
        """
        提供判定をキャンセルする
        """
        try:
            with self.processing_lock:
                if self.is_processing:
                    logging.info(f"★★★ 提供判定をキャンセルします（{button_name}ボタンクリック）★★★")
                    logging.info(f"- キャンセル時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                    logging.info(f"- 現在のCTI状態: {self.current_status}")
                    logging.info(f"- 通話開始からの経過時間: {time.time() - self.talking_start_time:.2f}秒")
                    
                    # 提供判定の実行状態をリセット
                    self.is_processing = False
                    
                    # 通話開始時間をリセット
                    self.talking_start_time = None
                    
                    logging.info("提供判定のキャンセルが完了しました")
                    
        except Exception as e:
            logging.error(f"提供判定のキャンセル中にエラー: {str(e)}")
            # エラー時も確実にフラグをリセット
            self.is_processing = False

    def find_action_buttons(self) -> bool:
        """
        アクションボタン（次、留守、担当者不在、NG）を検索
        
        Returns:
            bool: いずれかのボタンが見つかった場合True
        """
        try:
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