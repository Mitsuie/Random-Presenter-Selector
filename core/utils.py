"""
共通ユーティリティモジュール
リソースパスの動的解決など、環境依存を吸収する関数を提供します。
"""
import sys
from pathlib import Path

def get_resource_path(relative_path: str) -> Path:
    """
    開発環境（通常スクリプト実行時）と PyInstaller パッケージ化環境（exe実行時）の
    両方で正しいアセットファイルへの絶対パスを返します。
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller の一時展開ディレクトリ（--onefile）またはパッケージルート（--onedir）
        base_path = Path(sys._MEIPASS)
    else:
        # 通常実行時（プロジェクトルート基準: core/ の親ディレクトリ）
        base_path = Path(__file__).resolve().parent.parent

    return (base_path / relative_path).resolve()
