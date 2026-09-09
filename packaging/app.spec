# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

# ------------------------------------------------------------------------------
# パス・基本設定
# ------------------------------------------------------------------------------
spec_dir = os.path.abspath(SPECPATH)
project_root = os.path.abspath(os.path.join(spec_dir, '..')) if os.path.basename(spec_dir) == 'packaging' else spec_dir

# エントリーポイント
main_script = os.path.join(project_root, 'main.py')

# アプリ名・アイコン・バージョン情報定義
app_name = 'Random-Presenter-Selector'
icon_path = os.path.join(project_root, 'packaging', 'app_icon.ico')
version_file = os.path.join(project_root, 'packaging', 'version_info.txt')

datas = []
binaries = []
hiddenimports = ['openpyxl', 'core', 'ui']

# ------------------------------------------------------------------------------
# CustomTkinter のテーマ・アセット一括収集
# ------------------------------------------------------------------------------
ctk_ret = collect_all('customtkinter')
datas += ctk_ret[0]
binaries += ctk_ret[1]
hiddenimports += ctk_ret[2]

# assets ディレクトリが存在する場合は同梱
assets_dir = os.path.join(project_root, 'assets')
if os.path.exists(assets_dir):
    datas.append((assets_dir, 'assets'))

# ------------------------------------------------------------------------------
# Analysis（依存関係解析と不要モジュールの除外）
# ------------------------------------------------------------------------------
a = Analysis(
    [main_script],
    pathex=[project_root],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 配布パッケージ軽量化のため不要な大型ライブラリを明示的に除外
    excludes=[
        'matplotlib',
        'pandas',
        'scipy',
        'torch',
        'PIL.ImageQt',
        'tkinter.test',
        'unittest',
        'pydoc',
        'doctest',
        'test',
        'pytest',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# ------------------------------------------------------------------------------
# EXE 定義（GUI アプリ: console=False）
# ------------------------------------------------------------------------------
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if os.path.exists(icon_path) else None,
    version=version_file if os.path.exists(version_file) else None,
)

# ------------------------------------------------------------------------------
# COLLECT（高速起動のディレクトリ形式 --onedir 出力）
# ------------------------------------------------------------------------------
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)
