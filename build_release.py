
"""
リリースビルドスクリプト

このスクリプトは、アプリケーションのリリースビルドを実行します。
PyInstallerを使用して単一のexeファイルを生成します。
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def build_release():
    """リリースビルドを実行する"""
    try:
        # 一時ファイルとキャッシュの削除
        if os.path.exists("build"):
            shutil.rmtree("build")
        if os.path.exists("dist"):
            shutil.rmtree("dist")
        if os.path.exists("*.spec"):
            os.remove("*.spec")
        
        # PyInstallerでビルド
        result = subprocess.run(["pyinstaller", "build_release.spec"], 
                              capture_output=True, 
                              text=True)
        
        if result.returncode != 0:
            print("ビルドエラー:")
            print(result.stdout)
            print(result.stderr)
            return False
        
        print("\nビルドが完了しました。")
        print(f"生成されたファイル: dist/TelephoneTool-{VERSION}.exe")
        return True
        
    except Exception as e:
        print(f"ビルド中にエラーが発生しました: {e}")
        return False

if __name__ == "__main__":
    from version import VERSION
    success = build_release()
    sys.exit(0 if success else 1) 