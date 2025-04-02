"""
CTIサービス

このモジュールは、CTIシステムとの連携機能を提供します。
"""

import logging
from dataclasses import dataclass
from typing import Optional

@dataclass
class CTIData:
    """CTIから取得したデータを保持するクラス"""
    customer_name: Optional[str] = None
    phone: Optional[str] = None
    address: Optional[str] = None
    postal_code: Optional[str] = None

class CTIService:
    """CTIシステムとの連携を管理するクラス"""
    
    def __init__(self):
        """CTIServiceの初期化"""
        self.logger = logging.getLogger(__name__)
    
    def get_all_fields_data(self) -> Optional[CTIData]:
        """
        CTIから全てのフィールドデータを取得する
        
        Returns:
            Optional[CTIData]: CTIから取得したデータ。取得できない場合はNone
        """
        try:
            # TODO: 実際のCTIシステムとの連携処理を実装
            # 現時点ではダミーデータを返す
            return CTIData(
                customer_name="テスト顧客",
                phone="090-1234-5678",
                address="東京都渋谷区",
                postal_code="150-0002"
            )
        except Exception as e:
            self.logger.error(f"CTIデータの取得中にエラーが発生: {e}")
            return None 