"""
リリースビルドスクリプト

このスクリプトは、アプリケーションのリリースビルドを実行します。
PyInstallerを使用して単一のexeファイルを生成します。
"""

import os
import sys
import shutil
from pathlib import Path

def build_release():
    """リリースビルドを実行する"""
    # 一時ファイルとキャッシュの削除
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("*.spec"):
        os.remove("*.spec")
    
    # PyInstallerでビルド
    os.system("pyinstaller build_release.spec")
    
    print("\nビルドが完了しました。")
    print(f"生成されたファイル: dist/TelephoneTool-{VERSION}.exe")

if __name__ == "__main__":
    from version import VERSION
    build_release() 