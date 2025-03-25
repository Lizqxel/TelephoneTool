"""
アプリケーション起動ラッパー

メインアプリケーションを起動する前に必要なパッチと環境設定を適用するラッパーモジュールです。
"""

import os
import sys
import logging
import traceback

# 環境変数の設定
os.environ['PYTHONIOENCODING'] = 'utf-8'
os.environ['PYTHONUTF8'] = '1'
os.environ['QT_AUTO_SCREEN_SCALE_FACTOR'] = '1'
os.environ['QT_ENABLE_HIGHDPI_SCALING'] = '1'

# WebGLレンダリングの問題回避
os.environ['QTWEBENGINE_CHROMIUM_FLAGS'] = '--disable-gpu --enable-unsafe-swiftshader'

def main():
    """アプリケーション起動ラッパー関数"""
    try:
        # PySide6パッチの適用
        from pyside6_fix import apply_patches
        apply_patches()
        
        # メインアプリケーションをインポートして実行
        import main as app_main
        return app_main.main()
    except Exception as e:
        logging.error(f"アプリケーション起動中にエラーが発生しました: {e}")
        logging.error(traceback.format_exc())
        print(f"エラーが発生しました: {e}")
        print("詳細はログファイル (app.log) を確認してください。")
        input("Enterキーを押して終了...")
        return 1

if __name__ == "__main__":
    sys.exit(main()) 