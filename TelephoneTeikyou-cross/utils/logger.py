"""
ロギング機能を提供するモジュール

このモジュールは、アプリケーション全体で使用する
ロギング機能を提供します。
"""

import logging
import os
from datetime import datetime

# ログディレクトリの作成
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# ログファイルのパス
log_file = os.path.join(log_dir, f"app_{datetime.now().strftime('%Y%m%d')}.log")

# ロガーの設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# ファイルハンドラの設定
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# コンソールハンドラの設定
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# フォーマッタの設定
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# ハンドラの追加
logger.addHandler(file_handler)
logger.addHandler(console_handler) 