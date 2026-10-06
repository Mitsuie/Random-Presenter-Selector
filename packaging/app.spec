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

# アプリ名・バージョンは core/constants.py を唯一の定義元とする
sys.path.insert(0, project_root)
from core.constants import APP_NAME, APP_DISPLAY_NAME, APP_VERSION, APP_DESCRIPTION, COMPANY_NAME

app_name = APP_NAME
icon_path = os.path.join(project_root, 'packaging', 'app_icon.ico')


def _write_version_file():
    """Windows 実行ファイルのバージョン情報を constants.py から生成する"""
    nums = [int(p) for p in APP_VERSION.split('.')[:3] if p.isdigit()]
    nums = (nums + [0, 0, 0])[:3] + [0]
    ver_tuple = tuple(nums)
    ver_str = '.'.join(str(n) for n in nums)
    strings = [
        ('CompanyName', COMPANY_NAME),
        ('FileDescription', APP_DESCRIPTION),
        ('FileVersion', ver_str),
        ('InternalName', APP_NAME),
        ('LegalCopyright', f'Copyright (C) 2026 {COMPANY_NAME}. All rights reserved.'),
        ('OriginalFilename', f'{APP_NAME}.exe'),
        ('ProductName', APP_DISPLAY_NAME),
        ('ProductVersion', ver_str),
    ]
    structs = ', '.join(f'StringStruct({k!r}, {v!r})' for k, v in strings)
    content = (
        "VSVersionInfo(\n"
        f"  ffi=FixedFileInfo(filevers={ver_tuple}, prodvers={ver_tuple}, mask=0x3f, flags=0x0,\n"
        "                    OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),\n"
        "  kids=[\n"
        f"    StringFileInfo([StringTable('041104b0', [{structs}]), StringTable('040904b0', [{structs}])]),\n"
        "    VarFileInfo([VarStruct('Translation', [1041, 1200, 1033, 1200])])\n"
        "  ]\n"
        ")\n"
    )
    os.makedirs(workpath, exist_ok=True)
    path = os.path.join(workpath, 'version_info.txt')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    return path


version_file = _write_version_file()

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

# ウィンドウアイコン（実行時に get_resource_path("packaging/app_icon.ico") で参照）
if os.path.exists(icon_path):
    datas.append((icon_path, 'packaging'))

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
    version=version_file,
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
