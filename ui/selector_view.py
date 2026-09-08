import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from core.lottery_manager import LotteryManager, LotteryError
from ui.font_config import get_font_family

class SelectorView(ctk.CTkFrame):
    """学生指名機能のビュー（待機画面・結果画面を内包）"""

    def __init__(self, master, on_exit_request=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.on_exit_request = on_exit_request
        self.manager = LotteryManager()
        self.font_family = get_font_family()

        self.create_wait_screen()
        self.create_result_screen()
        self.show_wait_screen()

    def create_wait_screen(self):
        self.wait_frame = ctk.CTkFrame(self, fg_color="transparent")

        title_label = ctk.CTkLabel(
            self.wait_frame,
            text="演習投影学生指名プログラム",
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
        self.draw_btn.pack(pady=14)

    def create_result_screen(self):
        self.result_frame = ctk.CTkFrame(self, fg_color="transparent")

        # 当選学生カード（大画面）
        self.student_card = ctk.CTkFrame(self.result_frame, corner_radius=14)
        self.student_card.pack(pady=(28, 18), padx=30, fill="x")

        self.id_label = ctk.CTkLabel(
            self.student_card,
            text="学籍番号: --------",
            font=ctk.CTkFont(family=self.font_family, size=24, weight="bold"),
            text_color=("gray20", "gray85")
        )
        self.id_label.pack(pady=(18, 6))

        self.name_label = ctk.CTkLabel(
            self.student_card,
            text="氏名",
            font=ctk.CTkFont(family=self.font_family, size=40, weight="bold")
        )
        self.name_label.pack(pady=6)

        self.kana_label = ctk.CTkLabel(
            self.student_card,
            text="(フリガナ)",
            font=ctk.CTkFont(family=self.font_family, size=22),
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
            command=self.show_wait_screen,
            font=btn_font,
            width=w,
            height=h,
            fg_color=("#D3D3D3", "#696969"),
            hover_color=("#A9A9A9", "#505050"),
            text_color=("black", "white")
        )
        back_btn.pack(side=tk.LEFT, padx=12)

    def show_wait_screen(self):
        self.result_frame.pack_forget()
        self.wait_frame.pack(fill="both", expand=True)

    def show_result_screen(self):
        self.wait_frame.pack_forget()
        self.result_frame.pack(fill="both", expand=True)

    def choose_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.load_csv_file(filename)

    def load_csv_file(self, filename: str):
        """CSVファイルを読み込み、情報をUIに反映する"""
        try:
            stats = self.manager.load_csv(filename)
            self.update_stats_display(filename, stats)
        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました:\n{e}")

    def update_stats_display(self, filename: str, stats: dict):
        base_name = os.path.basename(filename)
        self.file_name_label.configure(text=f"対象CSV: {base_name}")
        self.lbl_total_val.configure(text=f"{stats['total']} 名")
        self.lbl_pending_val.configure(text=f"{stats['pending']} 名")
        self.lbl_done_val.configure(text=f"{stats['done']} 名")

    def start_draw(self):
        if not self.manager.filename:
            messagebox.showwarning("警告", "CSVファイルを選択してください。")
            return

        student = self.manager.draw_student()
        if not student:
            messagebox.showinfo("情報", "投影実施可否が空欄の学生はいません（全員実施済みです）。")
            return

        self.id_label.configure(text=f"学籍番号: {student['student_id']}")
        self.name_label.configure(text=student['name'])
        self.kana_label.configure(text=f"({student['kana']})")
        self.show_result_screen()

    def record_result(self, result: str):
        try:
            self.manager.save_result(result)
            stats = self.manager.get_statistics()
            self.update_stats_display(self.manager.filename, stats)
            messagebox.showinfo("保存完了", f"結果（{result}）をCSVに保存しました。")
            self.show_wait_screen()
        except PermissionError as pe:
            messagebox.showerror("ファイルロックエラー", str(pe))
        except Exception as e:
            messagebox.showerror("エラー", f"保存中にエラーが発生しました:\n{e}")

    def confirm_exit(self):
        if messagebox.askyesno("確認", "プログラムを終了しますか？"):
            if self.on_exit_request:
                self.on_exit_request()
            else:
                self.quit()
