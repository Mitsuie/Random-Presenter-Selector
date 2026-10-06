"""
利用者設定（外観モード・直近のCSV・欠席者の扱い）の保存と読み込み
保存先: %LOCALAPPDATA%\\Random-Presenter-Selector\\settings.json
"""
import json
from pathlib import Path
from typing import Any, Optional
from core.utils import get_app_data_dir


class Settings:
    DEFAULTS = {
        "theme": "Light",
        "last_csv_path": "",
        "include_absent": False,
    }

    def __init__(self, path: Optional[Path] = None):
        self.path = Path(path) if path else get_app_data_dir() / "settings.json"
        self.values = dict(self.DEFAULTS)
        try:
            with open(self.path, encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                for key, default in self.DEFAULTS.items():
                    if isinstance(loaded.get(key), type(default)):
                        self.values[key] = loaded[key]
        except (OSError, ValueError):
            # 初回起動・ファイル破損時は既定値を使う
            pass

    def get(self, key: str) -> Any:
        return self.values[key]

    def set(self, key: str, value: Any) -> None:
        self.values[key] = value
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.values, f, ensure_ascii=False, indent=2)
        except OSError:
            # 設定が保存できなくても授業の進行は妨げない
            pass
