"""
バージョン情報管理モジュール

このモジュールは、アプリケーションのバージョン情報を一元管理します。
"""

# アプリケーションのバージョン
VERSION = "4.0.5"

# アプリケーション名
APP_NAME = "TelephoneTool"

# アプリケーションの説明
APP_DESCRIPTION = "電話番号管理ツール"

# 開発者情報
DEVELOPER = "Lizqxel"
DEVELOPER_EMAIL = "your.email@example.com"

# ライセンス情報
LICENSE = "MIT License"

# GitHub関連の情報
# アップデートは exe だけを置く公開の配布用リポジトリから取る
# （ソースのリポジトリを非公開にしても、ツール内のアップデートが止まらないように）
GITHUB_OWNER = "Lizqxel"
GITHUB_REPO = "TelephoneTool-release"

# アップデートチェックの間隔（秒）
UPDATE_CHECK_INTERVAL = 86400  # 24時間


def version_key(text):
    """'v3.10.0' や '3.9.5-beta' を (3, 10, 0) のような数値の並びにする。"""
    import re
    return tuple(int(number) for number in re.findall(r"\d+", str(text).split("-")[0]))


def is_newer_version(latest, current):
    """latest が current より新しいか。文字の並びではなく数値で比べる（3.10.0 は 3.9.9 より新しい）。

    注意：3.9.4 以前の配布版は文字の並びで比べるため、3.10.0 などを新しいと判定できない。
    そのため次は 4.0.0 とし（先頭の「4」と「3」の比較なので古い版からも新しいと判定される）、
    以降はこの数値の比較で判定する。
    """
    a, b = version_key(latest), version_key(current)
    width = max(len(a), len(b))
    return a + (0,) * (width - len(a)) > b + (0,) * (width - len(b)) 