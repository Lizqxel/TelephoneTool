# -*- mode: python ; coding: utf-8 -*-

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

# バージョン情報の取得
VERSION = "1.0.1"

# 作業ディレクトリの取得
work_dir = os.path.abspath(SPECPATH)

block_cipher = None

# 追加するファイル
added_files = [
    ('settings.json', '.'),
    ('map.png', '.'),
    ('qt.conf', '.'),
    ('version.py', '.'),
    ('requirements.txt', '.'),
    ('icon.ico', '.'),
    ('utils', 'utils'),
    ('ui', 'ui'),
    ('services', 'services'),
    ('drivers', 'drivers')
]

# データファイルの収集
datas = []
for src, dst in added_files:
    if os.path.exists(os.path.join(work_dir, src)):
        datas.append((os.path.join(work_dir, src), dst))

# 実行ファイルの設定
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=[],
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
    name=f'TelephoneTool-{VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # アイコンを設定
) 