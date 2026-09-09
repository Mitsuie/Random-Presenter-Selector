import tkinter as tk
from tkinter import messagebox
from typing import Optional
import customtkinter as ctk
from core.constants import APP_VERSION, APP_DISPLAY_NAME, GITHUB_OWNER, GITHUB_REPO
from core.app_updater import check_for_updates_async, UpdateInfo
from core.utils import get_resource_path
from ui.update_dialog import UpdateDialog
from ui.font_config import get_font_family
from ui.selector_view import SelectorView
from ui.converter_view import ConverterView

class MainWindow(ctk.CTk):
    """演習投影統合管理システムのメインウィンドウ"""

    MODE_SELECTOR = "🎯 学生指名"
    MODE_CONVERTER = "📋 名簿変換"

    def __init__(self, initial_mode=None):
        super().__init__()

        # 基本設定
        self.title(APP_DISPLAY_NAME)
        self.geometry("720x620")
        self.resizable(False, False)

        # アプリアイコン設定
        icon_path = get_resource_path("packaging/app_icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        # 初期外観モード
        self.current_theme = "Light"
        ctk.set_appearance_mode(self.current_theme)
        ctk.set_default_color_theme("blue")

        # フォント設定
        self.font_family = get_font_family()

        # 画面構築
        self.create_header()
        self.create_content_area()

        # 初期表示モードの設定
        if initial_mode == self.MODE_CONVERTER:
            self.show_converter()
        else:
            self.show_selector()

        # ウィンドウを画面中央に配置
        self.center_window()

        # バックグラウンドで非同期更新確認（起動遅延ゼロ）
        self._start_update_check()

    def center_window(self):
        self.update_idletasks()
        w = 720
        h = 620
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def create_header(self):
        """最上部のヘッダーバー（モード切替 ＋ 外観切替）"""
        self.header_frame = ctk.CTkFrame(self, height=56, corner_radius=0, fg_color=("gray90", "gray17"))
        self.header_frame.pack(fill="x", side=tk.TOP)
        self.header_frame.pack_propagate(False)

        # モード切替セグメントボタン
        self.mode_segment = ctk.CTkSegmentedButton(
            self.header_frame,
            values=[self.MODE_SELECTOR, self.MODE_CONVERTER],
            command=self.on_mode_change,
            font=ctk.CTkFont(family=self.font_family, size=16, weight="bold"),
            height=38
        )
        self.mode_segment.set(self.MODE_SELECTOR)
        self.mode_segment.pack(side=tk.LEFT, padx=16, pady=9)

        # 終了ボタン（最右端）
        self.exit_btn = ctk.CTkButton(
            self.header_frame,
            text="🚪 終了",
            command=self.confirm_exit,
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            width=80,
            height=36,
            fg_color=("#F08080", "#CD5C5C"),
            hover_color=("#CD5C5C", "#8B0000"),
            text_color=("black", "white")
        )
        self.exit_btn.pack(side=tk.RIGHT, padx=(6, 16), pady=10)

        # 外観切替ボタン（終了ボタンの左隣）
        self.theme_btn = ctk.CTkButton(
            self.header_frame,
            text="🌓 外観切替",
            command=self.toggle_theme,
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            width=100,
            height=36,
            fg_color=("gray80", "gray28"),
            hover_color=("gray70", "gray38"),
            text_color=("black", "white")
        )
        self.theme_btn.pack(side=tk.RIGHT, padx=(6, 6), pady=10)

        # 更新案内ボタン（新バージョン検出時のみ動的に表示）
        self.update_btn = None

    def _start_update_check(self):
        """起動時にバックグラウンドで最新バージョンを確認"""
        check_for_updates_async(
            repo_owner=GITHUB_OWNER,
            repo_name=GITHUB_REPO,
            current_version=APP_VERSION,
            on_complete=self._on_update_checked
        )

    def _on_update_checked(self, info: Optional[UpdateInfo]):
        """ワーカースレッドからの完了通知をメインスレッドへ安全にディスパッチ"""
        self.after(0, lambda: self._apply_update_ui(info))

    def _apply_update_ui(self, info: Optional[UpdateInfo]):
        """新バージョンが存在する場合、ヘッダーに更新可能ボタンを表示"""
        if info and info.is_update_available:
            if not self.update_btn:
                self.update_btn = ctk.CTkButton(
                    self.header_frame,
                    text=f"🚀 v{info.version} 更新可能",
                    command=lambda: UpdateDialog(self, info, APP_VERSION),
                    font=ctk.CTkFont(family=self.font_family, size=13, weight="bold"),
                    height=36,
                    fg_color=("#0969da", "#1f6feb"),
                    hover_color=("#054da7", "#1158c7"),
                    text_color="white"
                )
                # 外観切替ボタンの左隣（RIGHTパック順）に配置
                self.update_btn.pack(side=tk.RIGHT, padx=(6, 6), pady=10)

    def confirm_exit(self):
        """プログラム終了の確認ダイアログ"""
        if messagebox.askyesno("確認", "プログラムを終了しますか？"):
            self.destroy()

    def create_content_area(self):
        """各機能画面を配置するコンテナエリア"""
        self.content_container = ctk.CTkFrame(self, fg_color="transparent")
        self.content_container.pack(fill="both", expand=True)

        # 学生指名ビュー
        self.selector_view = SelectorView(self.content_container)

        # 名簿変換ビュー
        self.converter_view = ConverterView(
            self.content_container,
            on_start_lottery=self.handle_start_lottery_from_converter
        )

    def on_mode_change(self, selected_mode):
        if selected_mode == self.MODE_SELECTOR:
            self.show_selector()
        elif selected_mode == self.MODE_CONVERTER:
            self.show_converter()

    def show_selector(self):
        self.converter_view.pack_forget()
        self.selector_view.pack(fill="both", expand=True)
        self.mode_segment.set(self.MODE_SELECTOR)

    def show_converter(self):
        self.selector_view.pack_forget()
        self.converter_view.pack(fill="both", expand=True)
        self.mode_segment.set(self.MODE_CONVERTER)

    def handle_start_lottery_from_converter(self, csv_file_path: str):
        """名簿変換完了後に、そのCSVを学生指名画面にセットして自動遷移する"""
        self.selector_view.load_csv_file(csv_file_path)
        self.show_selector()

    def toggle_theme(self):
        if self.current_theme == "Light":
            self.current_theme = "Dark"
            ctk.set_appearance_mode("Dark")
        else:
            self.current_theme = "Light"
            ctk.set_appearance_mode("Light")
