"""
アプリ本体の全モジュールが import できることを確認するスモークテスト

リファクタリング（不要ファイル・不要コードの削除）で
必要なモジュールを壊していないことを検出するためのもの。
"""

import importlib
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

PACKAGES = ["ui", "services", "utils"]


def _collect_modules():
    modules = ["main", "first_run_setup", "version"]
    for pkg in PACKAGES:
        for name in sorted(os.listdir(os.path.join(ROOT, pkg))):
            if name.endswith(".py") and name != "__init__.py":
                modules.append(f"{pkg}.{name[:-3]}")
    return modules


@pytest.mark.parametrize("module_name", _collect_modules())
def test_module_imports(module_name):
    importlib.import_module(module_name)
