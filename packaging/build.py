"""
Random-Presenter-Selector - Full-Auto Clean Build Engine
Supports:
- Clean temporary venv build (prevents bundling unused packages)
- Cloud-sync-safe local %TEMP% workspace (avoids WinError 5 / WinError 32 on Google Drive)
- Running process auto-kill
- PyInstaller standalone binary compilation (--onedir)
- Inno Setup installer compilation (robust auto-discovery across PATH, User AppData, Program Files, and Registry)
- Safe cleanup with read-only flag removal & retry logic
"""
import os
import sys
import shutil
import stat
import subprocess
import time
import tempfile
from pathlib import Path

# ==============================================================================
# プロジェクト基本設定
# ==============================================================================
APP_NAME = "Random-Presenter-Selector"       # アプリケーション識別名
APP_EXE_NAME = f"{APP_NAME}.exe"             # 出力される実行ファイル名
SPEC_FILENAME = "app.spec"                   # packaging/ 内の spec ファイル名
ISS_FILENAME = "installer.iss"               # packaging/ 内の iss ファイル名
REQUIREMENTS_FILENAME = "requirements.txt"   # 最小依存関係ファイル名

# 出力先ディレクトリ設定
DIST_DIR_NAME = "dist"                       # exe実行ファイル群の出力先ディレクトリ
INSTALLER_DIR_NAME = "dist_installer"        # インストーラーexeの出力先ディレクトリ
# ==============================================================================

def remove_readonly(func, path, excinfo):
    """Windowsの読み取り専用属性を強制解除して削除を再試行するハンドラ"""
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass

def safe_rmtree(path, retries=5, delay=0.5):
    """
    Windows環境のファイルロックや読み取り専用属性に対応した堅牢なディレクトリ削除
    """
    p = Path(path)
    if not p.exists():
        return

    for attempt in range(retries):
        try:
            if sys.version_info >= (3, 12):
                def _onexc(func, filepath, err):
                    try:
                        os.chmod(filepath, stat.S_IWRITE)
                        func(filepath)
                    except Exception:
                        pass
                shutil.rmtree(p, onexc=_onexc)
            else:
                shutil.rmtree(p, onerror=remove_readonly)

            if not p.exists():
                return
        except Exception:
            pass
        time.sleep(delay)

    if p.exists():
        try:
            shutil.rmtree(p, ignore_errors=True)
        except Exception:
            pass

def kill_existing_process():
    """実行中の対象プロセスがあれば安全に終了してファイルロックを解除"""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", APP_EXE_NAME, "/T"],
            capture_output=True,
            check=False
        )
    except Exception:
        pass

def find_iscc_path():
    """
    Inno Setup コンパイラ (ISCC.exe) をシステム全体から自動探索
    1. PATH 環境変数
    2. %LOCALAPPDATA%\\Programs\\Inno Setup 6 (一般ユーザー権限インストール)
    3. %ProgramFiles%\\Inno Setup 6, %ProgramFiles(x86)%\\Inno Setup 6 (管理者権限インストール)
    4. Windows レジストリ (HKCU / HKLM)
    """
    # 1. PATH からの探索
    which_iscc = shutil.which("ISCC.exe") or shutil.which("iscc")
    if which_iscc:
        return Path(which_iscc)

    # 2. 代表的なインストール先候補
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    program_files = os.environ.get("ProgramFiles", "")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "")

    candidates = [
        Path(local_app_data) / "Programs" / "Inno Setup 6" / "ISCC.exe" if local_app_data else None,
        Path(program_files_x86) / "Inno Setup 6" / "ISCC.exe" if program_files_x86 else None,
        Path(program_files) / "Inno Setup 6" / "ISCC.exe" if program_files else None,
        Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
        Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
    ]
    for cand in candidates:
        if cand and cand.exists():
            return cand

    # 3. Windows レジストリからの検出（HKCU, HKLM）
    try:
        import winreg
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for subkey in (
                r"Software\Microsoft\Windows\CurrentVersion\Uninstall",
                r"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            ):
                try:
                    with winreg.OpenKey(root, subkey) as k:
                        num_subkeys = winreg.QueryInfoKey(k)[0]
                        for i in range(num_subkeys):
                            try:
                                child_name = winreg.EnumKey(k, i)
                                with winreg.OpenKey(k, child_name) as ck:
                                    try:
                                        disp, _ = winreg.QueryValueEx(ck, "DisplayName")
                                        if "Inno Setup" in disp:
                                            loc, _ = winreg.QueryValueEx(ck, "InstallLocation")
                                            p = Path(loc) / "ISCC.exe"
                                            if p.exists():
                                                return p
                                    except OSError:
                                        pass
                            except OSError:
                                pass
                except OSError:
                    pass
    except Exception:
        pass

    return None

def main():
    # packaging フォルダの親ディレクトリ（プロジェクトルート）を基準にする
    root_dir = Path(__file__).resolve().parent.parent
    os.chdir(root_dir)

    print("=" * 65)
    print(f"  {APP_NAME} - Full-Auto Clean Build Engine")
    print("=" * 65)

    # クラウドストレージ同期（Google Drive, OneDrive等）のロックを回避するため
    # ビルド用仮想環境および作業中間ファイルはローカル Temp に配置
    temp_base = Path(tempfile.gettempdir())
    sanitized_name = "".join(c if c.isalnum() else "_" for c in APP_NAME).lower()
    venv_dir = temp_base / f"{sanitized_name}_build_venv"
    build_dir = temp_base / f"{sanitized_name}_build_work"

    dist_dir = root_dir / DIST_DIR_NAME
    installer_output_dir = root_dir / INSTALLER_DIR_NAME
    requirements_file = root_dir / REQUIREMENTS_FILENAME
    spec_file = root_dir / "packaging" / SPEC_FILENAME
    iss_file = root_dir / "packaging" / ISS_FILENAME

    # 1. 過去のビルド残骸・ロックのクリーンアップ
    print(f"[1/5] Cleaning previous build artifacts and creating clean virtual environment...")
    kill_existing_process()
    safe_rmtree(venv_dir)
    safe_rmtree(build_dir)
    safe_rmtree(dist_dir)

    # 仮想環境の作成
    venv_created = False
    for attempt in range(3):
        try:
            subprocess.run([sys.executable, "-m", "venv", "--clear", str(venv_dir)], check=True)
            venv_created = True
            break
        except Exception as e:
            print(f"[WARNING] Virtualenv creation attempt {attempt + 1} failed: {e}. Retrying...")
            safe_rmtree(venv_dir)
            time.sleep(1)

    if not venv_created:
        print("[ERROR] Failed to create clean virtual environment.")
        return 1

    venv_python = venv_dir / "Scripts" / "python.exe"
    venv_pip = venv_dir / "Scripts" / "pip.exe"
    venv_pyinstaller = venv_dir / "Scripts" / "pyinstaller.exe"

    # 2. 最小限の依存関係をインストール
    print(f"[2/5] Installing minimal dependencies from {REQUIREMENTS_FILENAME}...")
    try:
        subprocess.run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip"], check=False)
        if requirements_file.exists():
            subprocess.run([str(venv_pip), "install", "-r", str(requirements_file)], check=True)
        # PyInstaller を確実にインストール
        subprocess.run([str(venv_pip), "install", "pyinstaller"], check=True)
    except Exception as e:
        print(f"[ERROR] Failed to install dependencies: {e}")
        safe_rmtree(venv_dir)
        return 1

    # 3. PyInstaller によるディレクトリ形式バイナリ生成
    print("[3/5] Building standalone application with PyInstaller...")
    try:
        subprocess.run(
            [
                str(venv_pyinstaller),
                "--noconfirm",
                "--workpath", str(build_dir),
                "--distpath", str(dist_dir),
                str(spec_file)
            ],
            check=True
        )
        print("[INFO] PyInstaller build completed successfully.")
    except Exception as e:
        print(f"[ERROR] PyInstaller build failed: {e}")
        safe_rmtree(venv_dir)
        safe_rmtree(build_dir)
        return 1

    # 4. Inno Setup の自動検出とインストーラーコンパイル
    print("[4/5] Compiling Windows Setup Installer with Inno Setup...")
    installer_output_dir.mkdir(parents=True, exist_ok=True)
    iscc_path = find_iscc_path()

    if iscc_path and iss_file.exists():
        print(f"[INFO] Found Inno Setup compiler: {iscc_path}")
        res = subprocess.run([str(iscc_path), str(iss_file)])
        if res.returncode == 0:
            print("[INFO] Inno Setup installer compiled successfully!")
        else:
            print("[ERROR] Inno Setup compilation failed.")
    else:
        if not iscc_path:
            print("[WARNING] Inno Setup compiler (ISCC.exe) not found.")
            print("[INFO] To create an installer exe, download Inno Setup 6 from: https://jrsoftware.org/isdl.php")
        if not iss_file.exists():
            print(f"[WARNING] Inno Setup script not found: {iss_file}")

    # 5. 一時ビルド環境の完全クリーンアップ
    print("[5/5] Cleaning up temporary build virtualenv and cache...")
    safe_rmtree(venv_dir)
    safe_rmtree(build_dir)

    print("=" * 65)
    print("  BUILD PROCESS COMPLETED SUCCESSFULLY!")

    print(f"  - Standalone Application Folder : {dist_dir / APP_NAME}")
    if installer_output_dir.exists():
        print(f"  - Windows Setup Installer       : {installer_output_dir}")
    print("=" * 65)
    return 0

if __name__ == "__main__":
    sys.exit(main())
