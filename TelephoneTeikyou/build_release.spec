"""
リリース用のPyInstaller設定ファイル

このファイルは、実行可能ファイル（exe）を作成するための設定を定義します。
以下の機能を含みます：
- 必要なファイルとリソースの同梱
- アイコンの設定
- バージョン情報の設定
- 依存関係の管理
"""

import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# 現在のディレクトリをPythonパスに追加
sys.path.append(os.path.abspath(SPECPATH))
from version import VERSION  # バージョン情報を動的に読み込む

# 作業ディレクトリの取得
work_dir = os.path.abspath(SPECPATH)

# 実行ファイル名の設定
exe_name = f'TelephoneTeikyou-{VERSION}'  # バージョン番号を含める

block_cipher = None

# 追加するファイル
added_files = [
    ('qt.conf', '.'),
    ('version.py', '.'),
    ('requirements.txt', '.'),
    ('icon.ico', '.'),
    ('utils', 'utils'),
    ('ui', 'ui'),
    ('services', 'services')
]

# データファイルの収集
datas = []
datas += collect_data_files('selenium')
datas += collect_data_files('webdriver_manager')
datas += collect_data_files('pykakasi')  # pykakasi のデータファイルを追加
for src, dst in added_files:
    if os.path.exists(os.path.join(work_dir, src)):
        datas.append((os.path.join(work_dir, src), dst))

# 必要なサブモジュールを収集
hiddenimports = []
hiddenimports += collect_submodules('selenium')
hiddenimports += collect_submodules('webdriver_manager')
hiddenimports += collect_submodules('pykakasi')  # pykakasi のサブモジュールを追加
hiddenimports += ['jaconv']  # ひらがな→カタカナ変換に使用
hiddenimports += ['bs4', 'soupsieve']  # BeautifulSoup と依存関係を明示的に同梱

# 実行ファイルの設定
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 実行ファイルの作成
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=exe_name,  # バージョン番号を含む名前を使用
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'
) 