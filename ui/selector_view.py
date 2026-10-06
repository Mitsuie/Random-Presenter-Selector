import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from core.lottery_manager import LotteryManager, LotteryError
from core.settings import Settings
from ui.font_config import get_font_family

class SelectorView(ctk.CTkFrame):
    """学生指名機能のビュー（待機画面・結果画面を内包）"""

    def __init__(self, master, settings: Settings, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.settings = settings
        self.manager = LotteryManager()
        self.font_family = get_font_family()

        self.create_wait_screen()
        self.create_result_screen()
        self.show_wait_screen()

    def create_wait_screen(self):
        self.wait_frame = ctk.CTkFrame(self, fg_color="transparent")

        title_label = ctk.CTkLabel(
            self.wait_frame,
            text="学生指名",
            font=ctk.CTkFont(family=self.font_family, size=28, weight="bold")
        )
        title_label.pack(pady=(24, 14))

        # ファイル情報・ダッシュボードカード
        self.info_card = ctk.CTkFrame(self.wait_frame, corner_radius=12)
        self.info_card.pack(pady=8, padx=30, fill="x")

        self.file_name_label = ctk.CTkLabel(
            self.info_card,
            text="対象CSV: 未選択",
            font=ctk.CTkFont(family=self.font_family, size=18, weight="bold"),
            wraplength=640
        )
        self.file_name_label.pack(pady=(12, 6), padx=16)

        # 3連ステータスバッジコンテナ（コンパクト化）
        stats_container = ctk.CTkFrame(self.info_card, fg_color="transparent")
        stats_container.pack(fill="x", padx=16, pady=(0, 12))
        stats_container.grid_columnconfigure((0, 1, 2), weight=1, uniform="stats")

        # 1. 総人数ボックス
        box_total = ctk.CTkFrame(stats_container, corner_radius=8, fg_color=("#F3E8FF", "#2E1F47"))
        box_total.grid(row=0, column=0, padx=6, sticky="ew")

        lbl_total_title = ctk.CTkLabel(
            box_total,
            text="総学生数",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            text_color=("#6B21A8", "#C084FC")
        )
        lbl_total_title.pack(pady=(6, 1))

        self.lbl_total_val = ctk.CTkLabel(
            box_total,
            text="- 名",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold"),
            text_color=("#581C87", "#E9D5FF")
        )
        self.lbl_total_val.pack(pady=(0, 6))

        # 2. 未投影（残）ボックス
        box_pending = ctk.CTkFrame(stats_container, corner_radius=8, fg_color=("#E0F2FE", "#133E60"))
        box_pending.grid(row=0, column=1, padx=6, sticky="ew")

        lbl_pending_title = ctk.CTkLabel(
            box_pending,
            text="未投影 (残り)",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            text_color=("#0369A1", "#7DD3FC")
        )
        lbl_pending_title.pack(pady=(6, 1))

        self.lbl_pending_val = ctk.CTkLabel(
            box_pending,
            text="- 名",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold"),
            text_color=("#0284C7", "#38BDF8")
        )
        self.lbl_pending_val.pack(pady=(0, 6))

        # 3. 実施済ボックス
        box_done = ctk.CTkFrame(stats_container, corner_radius=8, fg_color=("#E8F8F5", "#144837"))
        box_done.grid(row=0, column=2, padx=6, sticky="ew")

        lbl_done_title = ctk.CTkLabel(
            box_done,
            text="実施済",
            font=ctk.CTkFont(family=self.font_family, size=12, weight="bold"),
            text_color=("#0F766E", "#5EEAD4")
        )
        lbl_done_title.pack(pady=(6, 1))

        self.lbl_done_val = ctk.CTkLabel(
            box_done,
            text="- 名",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold"),
            text_color=("#059669", "#34D399")
        )
        self.lbl_done_val.pack(pady=(0, 6))

        # ファイル選択ボタン（大型）
        select_file_btn = ctk.CTkButton(
            self.wait_frame,
            text="📂 CSVファイルを選択",
            command=self.choose_file,
            font=ctk.CTkFont(family=self.font_family, size=22, weight="bold"),
            width=360,
            height=65,
            fg_color=("#98FB98", "#2E8B57"),
            hover_color=("#90EE90", "#228B22"),
            text_color=("black", "white")
        )
        select_file_btn.pack(pady=14)

        # 抽選開始ボタン（大型）
        self.draw_btn = ctk.CTkButton(
            self.wait_frame,
            text="🎲 抽選開始",
            command=self.start_draw,
            font=ctk.CTkFont(family=self.font_family, size=24, weight="bold"),
            width=360,
            height=65,
            fg_color=("#87CEFA", "#2563EB"),
            hover_color=("#00BFFF", "#1D4ED8"),
            text_color=("black", "white")
        )
        self.draw_btn.pack(pady=(14, 2))

        hint_font = ctk.CTkFont(family=self.font_family, size=13)
        hint_color = ("gray40", "gray65")
        ctk.CTkLabel(
            self.wait_frame,
            text="キー操作: Space / Enter / PageDown で抽選　F11 で全画面",
            font=hint_font,
            text_color=hint_color
        ).pack()

        # 取り消し・表示設定
        tools_frame = ctk.CTkFrame(self.wait_frame, fg_color="transparent")
        tools_frame.pack(pady=(10, 0))

        self.undo_btn = ctk.CTkButton(
            tools_frame,
            text="↩ 直前の記録を取り消す",
            command=self.undo_last_record,
            font=ctk.CTkFont(family=self.font_family, size=14, weight="bold"),
            width=200,
            height=34,
            state="disabled",
            fg_color=("gray80", "gray28"),
            hover_color=("gray70", "gray38"),
            text_color=("black", "white"),
            text_color_disabled=("gray60", "gray45")
        )
        self.undo_btn.pack()

        self.include_absent_var = tk.BooleanVar(value=self.settings.get("include_absent"))
        self.manager.include_absent = bool(self.include_absent_var.get())
        ctk.CTkCheckBox(
            tools_frame,
            text="欠席者も抽選対象にする",
            variable=self.include_absent_var,
            command=self.on_toggle_include_absent,
            font=ctk.CTkFont(family=self.font_family, size=14)
        ).pack(pady=(10, 0))

        # 保存結果などの一時的なお知らせ
        self.status_label = ctk.CTkLabel(
            self.wait_frame,
            text="",
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            text_color=("#117864", "#4ADE80")
        )
        self.status_label.pack(pady=(8, 0))
        self._status_after_id = None

    def create_result_screen(self):
        self.result_frame = ctk.CTkFrame(self, fg_color="transparent")

        # 当選学生カード（大画面）
        self.student_card = ctk.CTkFrame(self.result_frame, corner_radius=14)
        self.student_card.pack(pady=(28, 18), padx=30, fill="x")

        # 全画面・リサイズ時にウィンドウの大きさに合わせて拡大するフォント
        self.result_fonts = [
            (ctk.CTkFont(family=self.font_family, size=24, weight="bold"), 24),
            (ctk.CTkFont(family=self.font_family, size=40, weight="bold"), 40),
            (ctk.CTkFont(family=self.font_family, size=22), 22),
        ]
        id_font, name_font, kana_font = (font for font, _ in self.result_fonts)

        self.id_label = ctk.CTkLabel(
            self.student_card,
            text="学籍番号: --------",
            font=id_font,
            text_color=("gray20", "gray85")
        )
        self.id_label.pack(pady=(18, 6))

        self.name_label = ctk.CTkLabel(
            self.student_card,
            text="氏名",
            font=name_font
        )
        self.name_label.pack(pady=6)

        self.kana_label = ctk.CTkLabel(
            self.student_card,
            text="(フリガナ)",
            font=kana_font,
            text_color=("gray30", "gray80")
        )
        self.kana_label.pack(pady=(4, 20))

        prompt_label = ctk.CTkLabel(
            self.result_frame,
            text="出席状況を選択してください",
            font=ctk.CTkFont(family=self.font_family, size=18, weight="bold")
        )
        prompt_label.pack(pady=10)

        # ボタン群（横並び）
        btn_frame = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        btn_frame.pack(pady=16)

        btn_font = ctk.CTkFont(family=self.font_family, size=20, weight="bold")
        w, h = 160, 70

        # 出席ボタン
        ok_btn = ctk.CTkButton(
            btn_frame,
            text="出席 (○)",
            command=lambda: self.record_result("○"),
            font=btn_font,
            width=w,
            height=h,
            fg_color=("#90EE90", "#2E8B57"),
            hover_color=("#3CB371", "#228B22"),
            text_color=("black", "white")
        )
        ok_btn.pack(side=tk.LEFT, padx=12)

        # 欠席ボタン
        ng_btn = ctk.CTkButton(
            btn_frame,
            text="欠席 (×)",
            command=lambda: self.record_result("×"),
            font=btn_font,
            width=w,
            height=h,
            fg_color=("#F08080", "#B22222"),
            hover_color=("#CD5C5C", "#8B0000"),
            text_color=("black", "white")
        )
        ng_btn.pack(side=tk.LEFT, padx=12)

        # 戻るボタン
        back_btn = ctk.CTkButton(
            btn_frame,
            text="記入なし\n(戻る)",
            command=self.cancel_selection,
            font=btn_font,
            width=w,
            height=h,
            fg_color=("#D3D3D3", "#696969"),
            hover_color=("#A9A9A9", "#505050"),
            text_color=("black", "white")
        )
        back_btn.pack(side=tk.LEFT, padx=12)

        ctk.CTkLabel(
            self.result_frame,
            text="キー操作: 1 / O = 出席　2 / X = 欠席　Esc = 戻る",
            font=ctk.CTkFont(family=self.font_family, size=13),
            text_color=("gray40", "gray65")
        ).pack()

        self._base_size = None
        self._font_scale = 1.0
        self.result_frame.bind("<Configure>", self._on_result_resize)

    def _on_result_resize(self, event):
        """ウィンドウの拡大に合わせて氏名などのフォントを大きくする（縮小はしない）"""
        if self._base_size is None:
            self._base_size = (event.width, event.height)
            return
        base_w, base_h = self._base_size
        scale = max(1.0, min(event.width / base_w, event.height / base_h))
        if abs(scale - self._font_scale) < 0.05:
            return
        self._font_scale = scale
        for font, base in self.result_fonts:
            font.configure(size=int(base * scale))

    def show_wait_screen(self):
        self.result_frame.pack_forget()
        self.wait_frame.pack(fill="both", expand=True)

    def cancel_selection(self):
        """記録せずに待機画面へ戻る"""
        self.manager.clear_selection()
        self.show_wait_screen()

    def show_result_screen(self):
        self.wait_frame.pack_forget()
        self.result_frame.pack(fill="both", expand=True)

    def handle_key(self, keysym: str) -> bool:
        """
        キーボード・プレゼン用リモコンでの操作。処理した場合は True を返す。
        """
        if self.result_frame.winfo_ismapped():
            if keysym in ("1", "o", "O"):
                self.record_result("○")
            elif keysym in ("2", "x", "X"):
                self.record_result("×")
            elif keysym == "Escape":
                self.cancel_selection()
            else:
                return False
            return True

        # PageDown / → はプレゼン用リモコンの「次へ」
        if keysym in ("space", "Return", "KP_Enter", "Next", "Right"):
            self.start_draw()
            return True
        return False

    def show_status(self, text: str, is_error: bool = False):
        """待機画面に数秒間だけお知らせを表示する"""
        if self._status_after_id:
            self.after_cancel(self._status_after_id)
        color = ("#B91C1C", "#F87171") if is_error else ("#117864", "#4ADE80")
        self.status_label.configure(text=text, text_color=color)
        self._status_after_id = self.after(6000, lambda: self.status_label.configure(text=""))

    def choose_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.load_csv_file(filename)

    def load_csv_file(self, filename: str, silent: bool = False) -> bool:
        """
        CSVファイルを読み込み、情報をUIに反映する。
        silent=True の場合は失敗してもダイアログを出さない（起動時の自動読み込み用）。
        """
        try:
            stats = self.manager.load_csv(filename)
        except Exception as e:
            if silent:
                self.show_status("前回のCSVを読み込めませんでした。CSVファイルを選択してください。", is_error=True)
            else:
                messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました:\n{e}")
            return False

        self.update_stats_display(filename, stats)
        self.update_undo_button()
        self.settings.set("last_csv_path", os.path.abspath(filename))
        # 抽選結果画面のまま別の名簿を読み込んだ場合に、前の学生が表示され続けないようにする
        self.show_wait_screen()
        return True

    def update_stats_display(self, filename: str, stats: dict):
        # 同名のCSVが別フォルダにあっても区別できるよう、親フォルダ名も表示する
        abs_path = os.path.abspath(filename)
        display = os.path.join(os.path.basename(os.path.dirname(abs_path)), os.path.basename(abs_path))
        self.file_name_label.configure(text=f"対象CSV: {display}")
        self.lbl_total_val.configure(text=f"{stats['total']} 名")
        self.lbl_pending_val.configure(text=f"{stats['pending']} 名")
        done_text = f"{stats['done']} 名"
        if stats.get("absent") and not self.manager.include_absent:
            done_text += f"（うち欠席 {stats['absent']}）"
        self.lbl_done_val.configure(text=done_text)

    def refresh_stats(self):
        if self.manager.filename:
            self.update_stats_display(self.manager.filename, self.manager.get_statistics())

    def update_undo_button(self):
        self.undo_btn.configure(state="normal" if self.manager.last_record else "disabled")

    def on_toggle_include_absent(self):
        self.manager.include_absent = bool(self.include_absent_var.get())
        self.settings.set("include_absent", self.manager.include_absent)
        self.refresh_stats()

    def start_draw(self):
        if not self.manager.filename:
            messagebox.showwarning("警告", "CSVファイルを選択してください。")
            return

        student = self.manager.draw_student()
        if not student:
            if self.manager.include_absent:
                msg = "抽選対象の学生はいません（全員が出席として記録済みです）。"
            else:
                msg = "投影実施可否が空欄の学生はいません（全員実施済みです）。"
            messagebox.showinfo("情報", msg)
            return

        self.id_label.configure(text=f"学籍番号: {student['student_id']}")
        self.name_label.configure(text=student['name'])
        self.kana_label.configure(text=f"({student['kana']})")
        self.show_result_screen()

    def record_result(self, result: str):
        try:
            reloaded = self.manager.save_result(result)
        except PermissionError as pe:
            messagebox.showerror("ファイルロックエラー", str(pe))
            return
        except LotteryError as le:
            # 外部変更で対象学生に記録できなかった場合など。最新の状態を表示して待機画面へ戻る
            self.update_stats_display(self.manager.filename, self.manager.get_statistics())
            messagebox.showerror("エラー", str(le))
            if self.manager.current_selected_index is None:
                self.show_wait_screen()
            return
        except Exception as e:
            messagebox.showerror("エラー", f"保存中にエラーが発生しました:\n{e}")
            return

        self.update_stats_display(self.manager.filename, self.manager.get_statistics())
        self.update_undo_button()
        msg = f"✓ {self.manager.last_record['name']} さん: {result} を保存しました"
        if reloaded:
            msg += "（外部で変更されたCSVを読み込み直しました）"
        self.show_status(msg)
        self.show_wait_screen()

    def undo_last_record(self):
        record = self.manager.last_record
        if not record:
            return
        if not messagebox.askyesno(
            "記録の取り消し",
            f"{record['name']} さん（学籍番号 {record['student_id']}）の記録「{record['value']}」を取り消しますか？\n"
            "取り消すと未投影に戻ります。"
        ):
            return
        try:
            self.manager.undo_last()
        except PermissionError as pe:
            messagebox.showerror("ファイルロックエラー", str(pe))
            return
        except Exception as e:
            messagebox.showerror("エラー", f"記録を取り消せませんでした:\n{e}")
            self.update_stats_display(self.manager.filename, self.manager.get_statistics())
            self.update_undo_button()
            return

        self.update_stats_display(self.manager.filename, self.manager.get_statistics())
        self.update_undo_button()
        self.show_status(f"↩ {record['name']} さんの記録を取り消しました")
