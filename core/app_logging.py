"""
ログ出力の設定
保存先: %LOCALAPPDATA%\\Random-Presenter-Selector\\app.log（1MB × 3世代）
※ 学生の氏名など名簿の内容はログに書かないこと（ファイルパスと例外情報のみ）
"""
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional
from core.utils import get_app_data_dir

LOG_FILENAME = "app.log"


def setup_logging() -> Optional[Path]:
    """
    ファイルへのログ出力を設定する。

    Returns:
        Optional[Path]: ログファイルのパス。ディレクトリを作れない場合は None（ログなしで続行）
    """
    log_path = get_app_data_dir() / LOG_FILENAME
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(log_path, maxBytes=1024 * 1024, backupCount=3, encoding="utf-8")
    except OSError:
        return None

    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    return log_path
