"""
地域検索サービスを提供するモジュール

このモジュールは、地域情報の検索と管理を行うサービスを提供します。
主な機能：
- 都道府県と市区町村のデータ管理
- 地域コードの検索
- 地域情報の取得
- 検索結果のキャッシュ

制限事項：
- キャッシュサイズは最大1000件
- 検索結果は最大100件まで
- データベース接続のタイムアウトは30秒
"""

import os
import sys
import json
import logging
import sqlite3
from typing import Dict, List, Optional, Tuple
from datetime import datetime, timedelta

class AreaSearchService:
    """地域検索サービスクラス"""
    
    def __init__(self, db_path: str = "data/area.db"):
        """
        サービスの初期化
        
        Args:
            db_path (str): データベースファイルのパス
        """
        self.db_path = db_path
        self.cache: Dict[str, Dict] = {}
        self.cache_expiry = timedelta(hours=1)
        self.setup_database()
        
    def setup_database(self) -> None:
        """データベースのセットアップ"""
        try:
            os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 都道府県テーブルの作成
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS prefectures (
                        code TEXT PRIMARY KEY,
                        name TEXT NOT NULL,
                        kana TEXT NOT NULL
                    )
                """)
                
                # 市区町村テーブルの作成
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS cities (
                        code TEXT PRIMARY KEY,
                        prefecture_code TEXT NOT NULL,
                        name TEXT NOT NULL,
                        kana TEXT NOT NULL,
                        FOREIGN KEY (prefecture_code) REFERENCES prefectures(code)
                    )
                """)
                
                conn.commit()
        except Exception as e:
            logging.error(f"データベースのセットアップ中にエラーが発生しました: {str(e)}")
            raise
            
    def get_prefectures(self) -> List[Dict[str, str]]:
        """
        都道府県一覧を取得します。
        
        Returns:
            List[Dict[str, str]]: 都道府県のリスト
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT code, name FROM prefectures ORDER BY kana")
                return [{"code": row[0], "name": row[1]} for row in cursor.fetchall()]
        except Exception as e:
            logging.error(f"都道府県一覧の取得中にエラーが発生しました: {str(e)}")
            return []
            
    def get_cities(self, prefecture_code: str) -> List[Dict[str, str]]:
        """
        市区町村一覧を取得します。
        
        Args:
            prefecture_code (str): 都道府県コード
            
        Returns:
            List[Dict[str, str]]: 市区町村のリスト
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT code, name FROM cities WHERE prefecture_code = ? ORDER BY kana",
                    (prefecture_code,)
                )
                return [{"code": row[0], "name": row[1]} for row in cursor.fetchall()]
        except Exception as e:
            logging.error(f"市区町村一覧の取得中にエラーが発生しました: {str(e)}")
            return []
            
    def search_by_area(
        self,
        prefecture_code: str,
        city_code: str,
        limit: int = 100
    ) -> List[Dict[str, str]]:
        """
        地域コードで検索します。
        
        Args:
            prefecture_code (str): 都道府県コード
            city_code (str): 市区町村コード
            limit (int): 取得件数の制限
            
        Returns:
            List[Dict[str, str]]: 検索結果のリスト
        """
        try:
            cache_key = f"{prefecture_code}:{city_code}"
            cached_result = self._get_from_cache(cache_key)
            if cached_result:
                return cached_result
                
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # 地域情報の取得
                cursor.execute("""
                    SELECT p.name, c.name
                    FROM prefectures p
                    JOIN cities c ON p.code = c.prefecture_code
                    WHERE p.code = ? AND c.code = ?
                """, (prefecture_code, city_code))
                
                result = cursor.fetchone()
                if not result:
                    return []
                    
                prefecture_name, city_name = result
                
                # TODO: 電話番号データの検索
                # ここで電話番号データベースから該当する地域の電話番号を取得
                
                # キャッシュに保存
                self._add_to_cache(cache_key, [])
                return []
        except Exception as e:
            logging.error(f"地域検索中にエラーが発生しました: {str(e)}")
            return []
            
    def _get_from_cache(self, key: str) -> Optional[List[Dict[str, str]]]:
        """
        キャッシュからデータを取得します。
        
        Args:
            key (str): キャッシュキー
            
        Returns:
            Optional[List[Dict[str, str]]]: キャッシュされたデータ、存在しない場合はNone
        """
        if key in self.cache:
            cached_data = self.cache[key]
            if datetime.now() - cached_data["timestamp"] < self.cache_expiry:
                return cached_data["data"]
            else:
                del self.cache[key]
        return None
        
    def _add_to_cache(self, key: str, data: List[Dict[str, str]]) -> None:
        """
        データをキャッシュに追加します。
        
        Args:
            key (str): キャッシュキー
            data (List[Dict[str, str]]): キャッシュするデータ
        """
        if len(self.cache) >= 1000:
            # 最も古いキャッシュを削除
            oldest_key = min(
                self.cache.keys(),
                key=lambda k: self.cache[k]["timestamp"]
            )
            del self.cache[oldest_key]
            
        self.cache[key] = {
            "data": data,
            "timestamp": datetime.now()
        } 