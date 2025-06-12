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
        """CTI状態の監視を開始"""
        if self.is_monitoring:
            logging.warning("CTI状態監視は既に開始されています")
            return
            
        self.is_monitoring = True
        logging.info("CTI状態監視を開始します")
        
        # 現在のCTI状態を取得して初期化
        self._initialize_current_status()
        
        # 監視スレッドを開始
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()
        
    def _initialize_current_status(self):
        """現在のCTI状態を取得して初期化"""
        try:
            # CTIウィンドウとコントロールを検出
            if self.find_cti_window() and self.find_status_text_control():
                # 現在の状態テキストを取得
                try:
                    status_text = win32gui.GetWindowText(self.status_text_handle).strip()
                    
                    # テキストから状態を判定
                    if status_text == "待ち受け中":
                        self.current_status = CTIStatus.WAITING
                        logging.info(f"CTI初期状態を設定: {self.current_status.value}")
                    elif status_text == "発信中":
                        self.current_status = CTIStatus.DIALING
                        logging.info(f"CTI初期状態を設定: {self.current_status.value}")
                    elif status_text == "通話中":
                        self.current_status = CTIStatus.TALKING
                        logging.info(f"CTI初期状態を設定: {self.current_status.value}")
                        # 通話中の場合は処理中フラグをFalseに（新規電話に備える）
                        self.is_processing = False
                    else:
                        self.current_status = CTIStatus.UNKNOWN
                        logging.info(f"CTI初期状態を設定: {self.current_status.value}")
                        
                    # previous_statusも同じ値に設定（初期化時の誤検出を防ぐ）
                    self.previous_status = self.current_status
                    
                except Exception as e:
                    logging.warning(f"CTI初期状態の取得に失敗: {str(e)}")
                    self.current_status = CTIStatus.UNKNOWN
                    self.previous_status = CTIStatus.UNKNOWN
            else:
                logging.warning("CTIウィンドウまたはコントロールが見つからないため、状態をUNKNOWNに設定")
                self.current_status = CTIStatus.UNKNOWN
                self.previous_status = CTIStatus.UNKNOWN
                
        except Exception as e:
            logging.error(f"CTI状態の初期化中にエラー: {str(e)}")
            self.current_status = CTIStatus.UNKNOWN
            self.previous_status = CTIStatus.UNKNOWN
        
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
                if current_time - self.last_detection_time >= self.monitor_interval:
                    self._check_status_change()
                    self.last_detection_time = current_time
                    
            except Exception as e:
                logging.error(f"CTI状態監視ループでエラーが発生: {str(e)}")
                
            # CPU負荷軽減のため短いスリープ
            time.sleep(0.05)
            
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
                    # デバッグレベルでログ出力（頻繁な出力を避ける）
                    logging.debug("CTIウィンドウが見つかりません")
                    return
                    
            # 状態表示コントロールの確認と再検出
            if not self.status_text_handle:
                if self.find_status_text_control():
                    logging.info("CTI状態表示コントロールを再検出しました")
                else:
                    logging.debug("CTI状態表示コントロールが見つかりません")
                    return
                    
            # 状態テキストを取得して状態判定
            try:
                status_text = win32gui.GetWindowText(self.status_text_handle).strip()
                
                if not status_text:
                    return
                
                # テキストから状態を判定
                if status_text == "待ち受け中":
                    new_status = CTIStatus.WAITING
                elif status_text == "発信中":
                    new_status = CTIStatus.DIALING
                elif status_text == "通話中":
                    new_status = CTIStatus.TALKING
                else:
                    new_status = CTIStatus.UNKNOWN
                    
                # 状態変化を検出
                self._detect_status_change(new_status)
                    
            except win32gui.error as e:
                # Win32 APIエラーの場合、コントロールハンドルをクリア
                logging.debug(f"状態テキスト取得でWin32エラー: {str(e)}")
                self.status_text_handle = None
            except Exception as e:
                logging.debug(f"状態テキスト取得中にエラー: {str(e)}")
                
        except Exception as e:
            # 重要なエラーのみINFOレベルで出力
            logging.info(f"CTI状態チェック中にエラーが発生: {str(e)}")
    
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
        
    def _detect_status_change(self, new_status: CTIStatus):
        """
        状態変化を検出し、必要に応じてコールバックを実行
        
        Args:
            new_status: 新しい状態
        """
        try:
            current_time = time.time()
            
            # 状態が変化した場合のみ処理
            if new_status != self.current_status:
                previous_status = self.current_status
                self.current_status = new_status
                
                logging.info(f"CTI状態が変化: {previous_status.value} → {new_status.value}")
                
                # 通話中状態への遷移を検出
                if new_status == CTIStatus.TALKING:
                    logging.info("★★★ 通話中状態を検出 ★★★")
                    if self.on_talking_started_callback:
                        try:
                            self.on_talking_started_callback()
                        except Exception as e:
                            logging.error(f"通話中状態開始コールバックの実行中にエラー: {str(e)}")
                
                # 通話終了の検出（通話中 → 待ち受け中）
                if (previous_status == CTIStatus.TALKING and 
                    new_status == CTIStatus.WAITING):
                    logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
                    # 処理中フラグをリセット（通話終了により新しい電話に備える）
                    self.is_processing = False
                    logging.info("通話終了により処理中フラグをリセットしました")
                    
                    # 通話終了コールバックを実行
                    if self.on_call_ended_callback:
                        try:
                            self.on_call_ended_callback()
                        except Exception as e:
                            logging.error(f"通話終了コールバックの実行中にエラー: {str(e)}")
                
                # 発信中→通話中の変化を検出（厳密な条件チェック付き）
                elif (previous_status == CTIStatus.DIALING and 
                      new_status == CTIStatus.TALKING):
                    
                    # 重複実行防止チェック
                    if self._should_trigger_auto_processing(current_time):
                        logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
                        logging.info("自動処理を開始します: 顧客情報取得 → 提供判定検索")
                        
                        # 最後の実行時刻を更新
                        self.last_dialing_to_talking_time = current_time
                        
                        # 自動処理を実行
                        self._trigger_auto_processing()
                    else:
                        logging.info("★★★ CTI状態変化検出: 発信中 → 通話中 ★★★")
                        logging.info("重複実行防止により自動処理をスキップしました")
                
                # 不正な状態変化パターンの検出と警告
                elif (previous_status == CTIStatus.UNKNOWN and 
                      new_status == CTIStatus.TALKING):
                    logging.warning("不正な状態変化を検出: 不明 → 通話中（処理をスキップ）")
                elif (previous_status == CTIStatus.WAITING and 
                      new_status == CTIStatus.TALKING):
                    logging.warning("待ち受け中から直接通話中への変化を検出（発信中を経由していない可能性）")
                    # この場合は自動処理を実行しない
                        
        except Exception as e:
            logging.error(f"状態変化検出中にエラーが発生: {str(e)}")
    
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