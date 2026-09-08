"""
アプリケーションアイコン自動生成スクリプト
案A（プレゼンタースポットライト＆ランダムダイスターゲット）から、
背景を美しく透過・トリミングし、高品質なマルチサイズ .ico ファイルを生成します。
"""
import os
from pathlib import Path
from PIL import Image, ImageDraw

def generate_app_icon(src_img_path: str, output_ico_path: str, output_png_path: str = None):
    img = Image.open(src_img_path).convert("RGBA")

    # アイコン本体の角丸四角形領域を正確にトリミング (171, 171) 〜 (897, 897)
    left, top, right, bottom = 171, 171, 897, 897
    cropped = img.crop((left, top, right, bottom))
    w, h = cropped.size

    # 4倍スーパーサンプリングで滑らかなアンチエイリアス角丸マスクを生成
    ss = 4
    mask = Image.new("L", (w * ss, h * ss), 0)
    draw = ImageDraw.Draw(mask)
    # モダンアプリアイコン標準の角丸比率（約22%）
    radius = int(w * ss * 0.22)
    draw.rounded_rectangle([0, 0, w * ss - 1, h * ss - 1], radius=radius, fill=255)
    mask = mask.resize((w, h), Image.Resampling.LANCZOS)

    # 角丸透過マスクを適用
    cropped.putalpha(mask)

    # 1. 透過PNGとして保存（高解像度 726x726）
    if output_png_path:
        out_png = Path(output_png_path)
        out_png.parent.mkdir(parents=True, exist_ok=True)
        cropped.save(str(out_png), format="PNG")
        print(f"[INFO] Transparent PNG saved at: {out_png}")

    # 2. マルチサイズ (16x16 〜 256x256) の Windows .ico ファイルとして保存
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    out_ico = Path(output_ico_path)
    out_ico.parent.mkdir(parents=True, exist_ok=True)
    cropped.save(str(out_ico), format="ICO", sizes=icon_sizes)
    print(f"[INFO] Multi-size ICO saved at: {out_ico}")

if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    src_mockup = r"C:\Users\mituz\.gemini\antigravity-ide\brain\7f0f2906-371a-47b0-90ad-193eb1a39229\presenter_selector_icon_mockup_1788888138372.jpg"
    out_ico = str(script_dir / "app_icon.ico")
    out_png = str(script_dir / "app_icon.png")

    generate_app_icon(src_mockup, out_ico, out_png)
