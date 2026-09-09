import tkinter.font as tkfont

_cached_font_family = None

def get_font_family() -> str:
    """
    環境で利用可能な最も視認性の高い日本語フォントを返す。
    優先順位: Meiryo UI -> Yu Gothic UI -> Segoe UI -> MS UI Gothic
    """
    global _cached_font_family
    if _cached_font_family is not None:
        return _cached_font_family

    try:
        available = set(tkfont.families())
        for candidate in ["Meiryo UI", "Yu Gothic UI", "Segoe UI", "MS UI Gothic"]:
            if candidate in available:
                _cached_font_family = candidate
                return _cached_font_family
    except Exception:
        pass

    _cached_font_family = "Meiryo UI"
    return _cached_font_family
