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
import traceback

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
                 on_talking_started_callback: Optional[Callable] = None,
                 on_cancel_processing_callback: Optional[Callable] = None):
        """
        初期化
        
        Args:
            on_dialing_to_talking_callback: 発信中→通話中の状態変化時のコールバック関数
            on_call_ended_callback: 通話終了時（通話中→待ち受け中）のコールバック関数
            on_talking_started_callback: 通話中状態開始時のコールバック関数
            on_cancel_processing_callback: アクションボタンクリック時の処理キャンセルコールバック関数
        """
        self.on_dialing_to_talking_callback = on_dialing_to_talking_callback
        self.on_call_ended_callback = on_call_ended_callback
        self.on_talking_started_callback = on_talking_started_callback
        self.on_cancel_processing_callback = on_cancel_processing_callback
        
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
        self.button_click_interval = 0.2  # ボタンクリック間隔（秒） - 0.5秒から0.2秒に短縮
        
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
                    self.monitor_interval = settings.get('cti_monitor_interval', 0.5)
                    self.call_duration_threshold = settings.get('call_duration_threshold', 0)
            else:
                self.enable_auto_processing = True
                self.monitor_interval = 0.5
                self.call_duration_threshold = 0
        except Exception as e:
            logging.error(f"CTI監視設定の読み込みに失敗しました: {str(e)}")
            self.enable_auto_processing = True
            self.monitor_interval = 0.5
            self.call_duration_threshold = 0
            
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
                    # 通話開始時刻を記録（発信中→通話中の場合のみ）
                    if previous_status == CTIStatus.DIALING:
                        self.talking_start_time = current_time
                        logging.info(f"通話開始時刻を記録: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(current_time))}")
                    elif self.talking_start_time == 0:
                        # 何らかの理由で通話開始時刻が記録されていない場合
                        self.talking_start_time = current_time
                        logging.info(f"通話開始時刻を補完記録: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(current_time))}")
                    
                    if self.on_talking_started_callback:
                        try:
                            self.on_talking_started_callback()
                        except Exception as e:
                            logging.error(f"通話中状態開始コールバックの実行中にエラー: {str(e)}")
                
                # 通話終了の検出（通話中 → 待ち受け中）
                if (previous_status == CTIStatus.TALKING and 
                    new_status == CTIStatus.WAITING):
                    logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
                    
                    # 処理中フラグはリセットしない（実際の処理状態を維持）
                    # アクションボタンによるキャンセルを可能にするため
                    if self.is_processing:
                        logging.info("処理中フラグは維持します（アクションボタンによるキャンセルを可能にするため）")
                        logging.info(f"- 現在の処理状態: 実行中")
                        if self.talking_start_time > 0:
                            elapsed_time = time.time() - self.talking_start_time
                            logging.info(f"- 処理開始からの経過時間: {elapsed_time:.1f}秒")
                    else:
                        logging.info("処理中フラグは既にリセット済みです")
                    
                    # 通話終了コールバックを実行
                    if self.on_call_ended_callback:
                        try:
                            self.on_call_ended_callback()
                        except Exception as e:
                            logging.error(f"通話終了コールバックの実行中にエラー: {str(e)}")
                
                # 発信中→通話中の変化を検出（TelephoneTeikyou-crossと同じ実装）
                elif (previous_status == CTIStatus.DIALING and 
                      new_status == CTIStatus.TALKING):
                    
                    logging.info("★★★ 発信中→通話中の状態変化を検出 ★★★")
                    logging.info(f"- コールバック設定: {self.on_dialing_to_talking_callback is not None}")
                    logging.info(f"- 自動処理有効: {self.enable_auto_processing}")
                    logging.info(f"- 通話時間閾値: {self.call_duration_threshold}秒")
                    
                    if (self.on_dialing_to_talking_callback and self.enable_auto_processing):
                        # 通話開始時間を記録
                        self.talking_start_time = current_time
                        
                        # 通話時間閾値をチェック
                        if self.call_duration_threshold > 0:
                            logging.info(f"★★★ {self.call_duration_threshold}秒後に自動処理を実行予定 ★★★")
                            # 指定秒数待機してから自動処理を実行
                            threading.Timer(
                                self.call_duration_threshold,
                                self._check_and_trigger_auto_processing
                            ).start()
                        else:
                            logging.info("★★★ 即座に自動処理を実行します ★★★")
                            
                            # 前回の処理フラグをチェック（新しい電話での処理開始前）
                            if self.is_processing:
                                logging.warning("前回の処理が完了していません。フラグをリセットして新しい処理を開始します")
                                logging.info(f"- 前回処理の通話開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.talking_start_time)) if self.talking_start_time > 0 else '不明'}")
                                if self.talking_start_time > 0:
                                    elapsed_time = time.time() - self.talking_start_time
                                    logging.info(f"- 前回処理からの経過時間: {elapsed_time:.1f}秒")
                                # 前回のフラグをリセット
                                self.is_processing = False
                                self.talking_start_time = 0
                                logging.info("- 前回の処理フラグをリセットしました")
                            
                            # 新しい処理を開始
                            self.is_processing = True
                            logging.info("- 新しい自動処理を開始します")
                            
                            # 即座に自動処理を実行
                            try:
                                self.on_dialing_to_talking_callback()
                            except Exception as e:
                                logging.error(f"自動処理の実行中にエラーが発生: {str(e)}")
                                # エラー時はフラグをリセット
                                self.is_processing = False
                                self.talking_start_time = 0
                            finally:
                                # 処理完了後、バックアップとして一定時間後にフラグをリセット（念のため）
                                def backup_reset_processing_flag():
                                    try:
                                        # 長時間経過した場合のバックアップリセット（異常終了対策）
                                        if self.is_processing and self.talking_start_time > 0:
                                            elapsed_time = time.time() - self.talking_start_time
                                            if elapsed_time > 300:  # 5分以上経過した場合
                                                logging.warning(f"処理開始から{elapsed_time:.0f}秒経過：バックアップ処理でフラグをリセットします")
                                                self.is_processing = False
                                                self.talking_start_time = 0
                                            else:
                                                logging.debug(f"処理継続中（経過時間: {elapsed_time:.1f}秒）")
                                        else:
                                            logging.debug("バックアップチェック：処理は既に完了済みです")
                                    except Exception as e:
                                        logging.error(f"バックアップフラグリセット中にエラー: {str(e)}")
                                        # エラーが発生した場合は強制的にリセット
                                        self.is_processing = False
                                        self.talking_start_time = 0
                                        
                                timer = threading.Timer(300.0, backup_reset_processing_flag)  # 5分後にバックアップチェック
                                timer.daemon = True
                                timer.start()
                    else:
                        if not self.on_dialing_to_talking_callback:
                            logging.warning("発信中→通話中コールバックが設定されていません")
                        if not self.enable_auto_processing:
                            logging.warning("自動処理が無効になっています")
                
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
    
    

    def _check_and_trigger_auto_processing(self):
        """
        通話時間閾値後の自動処理実行チェック（TelephoneTeikyou-crossと同じ実装）
        """
        try:
            # 現在も通話中かチェック
            current_status = self.get_current_status()
            if current_status == CTIStatus.TALKING and self.talking_start_time > 0:
                # 通話時間を計算
                elapsed_time = time.time() - self.talking_start_time
                if elapsed_time >= self.call_duration_threshold:
                    logging.info(f"★★★ 通話時間{elapsed_time:.1f}秒が閾値{self.call_duration_threshold}秒を超えたため自動処理を実行 ★★★")
                    if self.on_dialing_to_talking_callback:
                        # 前回の処理フラグをチェック（新しい電話での処理開始前）
                        if self.is_processing:
                            logging.warning("前回の処理が完了していません。フラグをリセットして新しい処理を開始します")
                            logging.info(f"- 前回処理の通話開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.talking_start_time)) if self.talking_start_time > 0 else '不明'}")
                            if self.talking_start_time > 0:
                                elapsed_time_prev = time.time() - self.talking_start_time
                                logging.info(f"- 前回処理からの経過時間: {elapsed_time_prev:.1f}秒")
                            # 前回のフラグをリセット
                            self.is_processing = False
                            self.talking_start_time = 0
                            logging.info("- 前回の処理フラグをリセットしました")
                        
                        # 新しい処理を開始
                        self.is_processing = True
                        logging.info("- 新しい自動処理を開始します")
                        
                        try:
                            self.on_dialing_to_talking_callback()
                        except Exception as e:
                            logging.error(f"自動処理の実行中にエラーが発生: {str(e)}")
                            # エラー時はフラグをリセット
                            self.is_processing = False
                            self.talking_start_time = 0
                        finally:
                            # 処理完了後、バックアップとして一定時間後にフラグをリセット（念のため）
                            def backup_reset_processing_flag():
                                try:
                                    # 長時間経過した場合のバックアップリセット（異常終了対策）
                                    if self.is_processing and self.talking_start_time > 0:
                                        elapsed_time_backup = time.time() - self.talking_start_time
                                        if elapsed_time_backup > 300:  # 5分以上経過した場合
                                            logging.warning(f"処理開始から{elapsed_time_backup:.0f}秒経過：バックアップ処理でフラグをリセットします")
                                            self.is_processing = False
                                            self.talking_start_time = 0
                                        else:
                                            logging.debug(f"処理継続中（経過時間: {elapsed_time_backup:.1f}秒）")
                                    else:
                                        logging.debug("バックアップチェック：処理は既に完了済みです")
                                except Exception as e:
                                    logging.error(f"バックアップフラグリセット中にエラー: {str(e)}")
                                    # エラーが発生した場合は強制的にリセット
                                    self.is_processing = False
                                    self.talking_start_time = 0
                                    
                            timer = threading.Timer(300.0, backup_reset_processing_flag)  # 5分後にバックアップチェック
                            timer.daemon = True
                            timer.start()
                else:
                    logging.info(f"通話時間{elapsed_time:.1f}秒が閾値{self.call_duration_threshold}秒未満のため自動処理をスキップ")
            else:
                logging.info("通話が既に終了しているため自動処理をスキップ")
        except Exception as e:
            logging.error(f"自動処理実行チェック中にエラー: {str(e)}")

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
        アクションボタン（「次」「留守」「担当者不在」「NG」）のクリックをチェック
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
            
            # マウスクリックをチェック（より精密な検出）
            # 左クリックの状態を確認（押下中と離した瞬間の両方をチェック）
            left_button_state = win32api.GetAsyncKeyState(win32con.VK_LBUTTON)
            
            # ボタンが押されている状態（0x8000）または押されて離された状態（0x0001）をチェック
            if (left_button_state & 0x8000) or (left_button_state & 0x0001):
                current_time = time.time()
                # 連続クリックを防ぐ（間隔を短縮してより敏感に検出）
                if current_time - self.last_button_click_time >= 0.2:  # 0.5秒から0.2秒に短縮
                    # マウス座標を取得
                    x, y = win32api.GetCursorPos()
                    
                    # 各ボタンの位置をチェック
                    clicked_button = None
                    button_rect = None
                    
                    # ボタンハンドルの有効性を確認してからチェック
                    if self.next_button_handle and win32gui.IsWindow(self.next_button_handle):
                        try:
                            rect = win32gui.GetWindowRect(self.next_button_handle)
                            if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                                clicked_button = "次"
                                button_rect = rect
                        except Exception:
                            self.next_button_handle = None
                            
                    if self.rusu_button_handle and win32gui.IsWindow(self.rusu_button_handle):
                        try:
                            rect = win32gui.GetWindowRect(self.rusu_button_handle)
                            if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                                clicked_button = "留守"
                                button_rect = rect
                        except Exception:
                            self.rusu_button_handle = None
                            
                    if self.tantou_fuzai_button_handle and win32gui.IsWindow(self.tantou_fuzai_button_handle):
                        try:
                            rect = win32gui.GetWindowRect(self.tantou_fuzai_button_handle)
                            if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                                clicked_button = "担当者不在"
                                button_rect = rect
                        except Exception:
                            self.tantou_fuzai_button_handle = None
                            
                    if self.ng_button_handle and win32gui.IsWindow(self.ng_button_handle):
                        try:
                            rect = win32gui.GetWindowRect(self.ng_button_handle)
                            if (rect[0] <= x <= rect[2] and rect[1] <= y <= rect[3]):
                                clicked_button = "NG"
                                button_rect = rect
                        except Exception:
                            self.ng_button_handle = None
                    
                    if clicked_button:
                        # 提供判定の実行状態を確認
                        is_processing = False
                        with self.processing_lock:
                            is_processing = self.is_processing or (self.processing_thread and self.processing_thread.is_alive())
                        
                        logging.info(f"★★★ 「{clicked_button}」ボタンがクリックされました ★★★")
                        logging.info(f"- クリック時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                        logging.info(f"- マウス座標: x={x}, y={y}")
                        logging.info(f"- ボタン位置: left={button_rect[0]}, top={button_rect[1]}, right={button_rect[2]}, bottom={button_rect[3]}")
                        logging.info(f"- 現在のCTI状態: {self.current_status.value}")
                        logging.info(f"- 提供判定状態: {'実行中' if is_processing else '未実行'}")
                        
                        self.last_button_click_time = current_time
                        
                        # 提供判定をキャンセル
                        self._cancel_processing(clicked_button)
                        
        except Exception as e:
            logging.error(f"アクションボタンクリックの検出中にエラー: {str(e)}")

    def _cancel_processing(self, button_name: str):
        """
        提供判定をキャンセルし、フラグをリセット
        
        Args:
            button_name: クリックされたボタンの名前
        """
        try:
            with self.processing_lock:
                logging.info(f"★★★ 「{button_name}」ボタンクリックによる処理キャンセルを開始 ★★★")
                logging.info(f"- キャンセル開始時刻: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                logging.info(f"- 現在のCTI状態: {self.current_status.value}")
                
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

    def _check_cti_status(self):
        """CTIの状態を監視し、状態変化を検出する"""
        try:
            # CTIの状態を取得
            current_status = self.get_current_status()
            
            # 状態が変化した場合
            if current_status != self.current_status:
                logging.info(f"CTI状態が変化: {self.current_status} → {current_status}")
                
                # 通話終了を検出
                if self.current_status == "通話中" and current_status == "待ち受け中":
                    logging.info("★★★ 通話終了を検出: 通話中 → 待ち受け中 ★★★")
                    
                    # 提供判定が実行中でない場合のみフラグをリセット
                    if not self.is_processing:
                        logging.info("通話終了により処理中フラグをリセットしました")
                        self.talking_start_time = 0
                    
                    # 電話ボタン監視を再開
                    logging.info("★★★ 通話終了を検出: 2秒後に電話ボタン監視を再開します ★★★")
                    threading.Timer(2.0, self._start_phone_button_monitoring).start()
                
                # 状態を更新
                self.current_status = current_status
                
                # 状態変化時のコールバックを実行
                if current_status == "待ち受け中" and self.on_call_ended_callback:
                    self.on_call_ended_callback()
                elif current_status == "通話中" and self.on_talking_started_callback:
                    self.on_talking_started_callback()
                elif current_status == "発信中" and self.on_dialing_to_talking_callback:
                    self.on_dialing_to_talking_callback()
                
                # 通話開始を検出
                if current_status == "通話中" and self.current_status == "発信中":
                    logging.info("★★★ 通話開始を検出: 発信中 → 通話中 ★★★")
                    self._handle_dialing_to_talking()
                
        except Exception as e:
            logging.error(f"CTI状態監視中にエラーが発生: {str(e)}")
            logging.error(traceback.format_exc()) 