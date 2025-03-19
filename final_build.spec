# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules, collect_data_files

# 現在の作業ディレクトリを取得
current_dir = os.path.abspath(os.getcwd())

# すべての依存関係を明示的に収集
all_hiddenimports = []
all_hiddenimports += collect_submodules('PySide6')
all_hiddenimports += collect_submodules('selenium')
all_hiddenimports += collect_submodules('webdriver_manager')
all_hiddenimports += [
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.common.keys',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
    'selenium.webdriver.chrome.options',
    'webdriver_manager.chrome',
    'webdriver_manager.core.driver',
    'webdriver_manager.core.download_manager',
    'pyside6_fix',
    'browser_options',
]

# スクリプトの依存関係も収集
ui_modules = collect_submodules('ui')
services_modules = collect_submodules('services')
utils_modules = collect_submodules('utils')

all_hiddenimports += ui_modules + services_modules + utils_modules

# データファイルも明示的に収集
datas = [
    ('settings.json', '.'),
    ('map.png', '.'),
    ('ui/*.py', 'ui'),
    ('services/*.py', 'services'),
    ('utils/*.py', 'utils'),
    ('chrome_data/', 'chrome_data'),
    ('*.json', '.'),
    ('debug_*.png', '.'),
    ('pyside6_fix.py', '.'),
    ('browser_options.py', '.'),
]

# PySide6のデータファイルも収集
datas += collect_data_files('PySide6')

# 基本的な依存関係のスキャン - スタートアップラッパーを使用
a = Analysis(
    ['startup.py'],  # メインスクリプトをstartup.pyに変更
    pathex=[current_dir],
    binaries=[],
    datas=datas,
    hiddenimports=all_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# PyQtとPySide6のスタイル設定をバイナリとして含める
a.binaries += [
    ('qt.conf', os.path.join(current_dir, 'qt.conf'), 'DATA')
]

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# GUI設定のマージを回避するディレクトリ形式でのビルド
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TelephoneTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPXは無効化
    console=True,  # デバッグ表示のためにコンソールは有効に
    icon='map.png',
)

# すべてのバイナリとデータを含めた最終パッケージ
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='TelephoneTool',
) 