"""
CTI（Computer Telephony Integration）サービスを提供するモジュール

このモジュールは、電話システムとの連携を行うサービスを提供します。
主な機能：
- 電話番号のダイヤル
- 通話状態の監視
- 通話履歴の管理
- CTIサーバーとの通信

制限事項：
- 通話履歴は最大1000件まで保存
- 接続タイムアウトは30秒
- ダイヤル間隔は最小1秒
"""

import os
import sys
import json
import logging
import socket
import time
from typing import Dict, List, Optional, Tuple
from datetime import datetime

class CTIService:
    """CTIサービスクラス"""
    
    def __init__(self, host: str = "localhost", port: int = 5000):
        """
        サービスの初期化
        
        Args:
            host (str): CTIサーバーのホスト名
            port (int): CTIサーバーのポート番号
        """
        self.host = host
        self.port = port
        self.socket = None
        self.connected = False
        self.call_history: List[Dict] = []
        
    def connect(self) -> bool:
        """
        CTIサーバーに接続します。
        
        Returns:
            bool: 接続に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            if self.connected:
                return True
                
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(30)
            self.socket.connect((self.host, self.port))
            self.connected = True
            logging.info("CTIサーバーに接続しました。")
            return True
        except Exception as e:
            logging.error(f"CTIサーバーへの接続中にエラーが発生しました: {str(e)}")
            self.connected = False
            return False
            
    def disconnect(self) -> None:
        """CTIサーバーから切断します。"""
        try:
            if self.socket:
                self.socket.close()
            self.connected = False
            logging.info("CTIサーバーから切断しました。")
        except Exception as e:
            logging.error(f"CTIサーバーからの切断中にエラーが発生しました: {str(e)}")
            
    def dial(self, phone_number: str) -> bool:
        """
        電話番号にダイヤルします。
        
        Args:
            phone_number (str): ダイヤルする電話番号
            
        Returns:
            bool: ダイヤルに成功した場合はTrue、失敗した場合はFalse
        """
        try:
            if not self.connected and not self.connect():
                return False
                
            # ダイヤルコマンドの送信
            command = {
                "type": "dial",
                "number": phone_number,
                "timestamp": datetime.now().isoformat()
            }
            self._send_command(command)
            
            # 通話履歴に追加
            self._add_to_call_history({
                "number": phone_number,
                "timestamp": datetime.now(),
                "status": "dialing"
            })
            
            return True
        except Exception as e:
            logging.error(f"ダイヤル中にエラーが発生しました: {str(e)}")
            return False
            
    def hangup(self) -> bool:
        """
        通話を切断します。
        
        Returns:
            bool: 切断に成功した場合はTrue、失敗した場合はFalse
        """
        try:
            if not self.connected:
                return False
                
            # 切断コマンドの送信
            command = {
                "type": "hangup",
                "timestamp": datetime.now().isoformat()
            }
            self._send_command(command)
            
            # 通話履歴を更新
            if self.call_history:
                self.call_history[-1]["status"] = "ended"
                self.call_history[-1]["end_time"] = datetime.now()
                
            return True
        except Exception as e:
            logging.error(f"通話切断中にエラーが発生しました: {str(e)}")
            return False
            
    def get_call_status(self) -> Dict[str, str]:
        """
        現在の通話状態を取得します。
        
        Returns:
            Dict[str, str]: 通話状態の情報
        """
        try:
            if not self.connected:
                return {"status": "disconnected"}
                
            # 状態取得コマンドの送信
            command = {
                "type": "get_status",
                "timestamp": datetime.now().isoformat()
            }
            response = self._send_command(command)
            
            return response.get("status", {"status": "unknown"})
        except Exception as e:
            logging.error(f"通話状態の取得中にエラーが発生しました: {str(e)}")
            return {"status": "error"}
            
    def get_call_history(
        self,
        limit: int = 100,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None
    ) -> List[Dict]:
        """
        通話履歴を取得します。
        
        Args:
            limit (int): 取得件数の制限
            start_date (Optional[datetime]): 開始日時
            end_date (Optional[datetime]): 終了日時
            
        Returns:
            List[Dict]: 通話履歴のリスト
        """
        try:
            history = self.call_history.copy()
            
            # 日時でフィルタリング
            if start_date:
                history = [h for h in history if h["timestamp"] >= start_date]
            if end_date:
                history = [h for h in history if h["timestamp"] <= end_date]
                
            # 件数制限
            return history[-limit:]
        except Exception as e:
            logging.error(f"通話履歴の取得中にエラーが発生しました: {str(e)}")
            return []
            
    def _send_command(self, command: Dict) -> Dict:
        """
        CTIサーバーにコマンドを送信します。
        
        Args:
            command (Dict): 送信するコマンド
            
        Returns:
            Dict: サーバーからの応答
        """
        try:
            if not self.connected:
                raise ConnectionError("CTIサーバーに接続されていません。")
                
            # コマンドの送信
            self.socket.send(json.dumps(command).encode())
            
            # 応答の受信
            response = self.socket.recv(4096)
            return json.loads(response.decode())
        except Exception as e:
            logging.error(f"コマンドの送信中にエラーが発生しました: {str(e)}")
            raise
            
    def _add_to_call_history(self, call: Dict) -> None:
        """
        通話履歴に追加します。
        
        Args:
            call (Dict): 追加する通話情報
        """
        if len(self.call_history) >= 1000:
            self.call_history.pop(0)
        self.call_history.append(call) 