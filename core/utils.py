"""
共通ユーティリティモジュール
リソースパスの動的解決など、環境依存を吸収する関数を提供します。
"""
import os
import sys
import tempfile
from pathlib import Path
from typing import Callable, IO


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


def atomic_write_text(path: str, write_func: Callable[[IO[str]], None], encoding: str = 'utf-8-sig') -> None:
    """
    同じディレクトリの一時ファイルに書き込み、成功した場合のみ os.replace で置き換える。
    書き込み途中で失敗しても元のファイルは壊れない。

    Raises:
        PermissionError: 置き換え先がExcel等で開かれている場合など
    """
    dir_name = os.path.dirname(os.path.abspath(path))
    fd, tmp_path = tempfile.mkstemp(dir=dir_name, prefix='.~', suffix='.tmp')
    try:
        with os.fdopen(fd, mode='w', encoding=encoding, newline='') as f:
            write_func(f)
        os.replace(tmp_path, path)
    except BaseException:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise
