import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from core.roster_converter import convert_excel_to_csv, RosterConversionError
from ui.font_config import get_font_family

class ConverterView(ctk.CTkFrame):
    """名簿フォーマット変換機能のビュー（Excel -> CSV）"""

    def __init__(self, master, on_start_lottery=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.on_start_lottery = on_start_lottery
        self.font_family = get_font_family()

        self.input_excel_path = ""
        self.output_csv_path = ""

        self.create_widgets()

    def create_widgets(self):
        title_label = ctk.CTkLabel(
            self,
            text="名簿フォーマット変換",
            font=ctk.CTkFont(family=self.font_family, size=28, weight="bold")
        )
        title_label.pack(pady=(24, 6))

        subtitle_label = ctk.CTkLabel(
            self,
            text="Excel名簿から必須列（学籍番号・氏名・カナ）を抽出し、投影管理用CSVを生成します",
            font=ctk.CTkFont(family=self.font_family, size=15),
            text_color=("gray30", "gray75")
        )
        subtitle_label.pack(pady=(0, 16))

        card = ctk.CTkFrame(self, corner_radius=12)
        card.pack(padx=30, fill="x", pady=6)

        # Step 1: 入力ファイル
        step1_lbl = ctk.CTkLabel(
            card,
            text="[Step 1] 変換元のExcel名簿 (.xlsx)",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold")
        )
        step1_lbl.pack(anchor="w", padx=20, pady=(16, 6))

        step1_frame = ctk.CTkFrame(card, fg_color="transparent")
        step1_frame.pack(fill="x", padx=20, pady=(0, 14))

        self.entry_excel = ctk.CTkEntry(
            step1_frame,
            placeholder_text="Excelファイルを選択してください",
            font=ctk.CTkFont(family=self.font_family, size=15),
            height=44
        )
        self.entry_excel.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))

        btn_browse_excel = ctk.CTkButton(
            step1_frame,
            text="📂 参照...",
            command=self.browse_excel,
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            width=110,
            height=44
        )
        btn_browse_excel.pack(side=tk.RIGHT)

        # Step 2: 出力ファイル
        step2_lbl = ctk.CTkLabel(
            card,
            text="[Step 2] 出力先CSVファイル (.csv)",
            font=ctk.CTkFont(family=self.font_family, size=17, weight="bold")
        )
        step2_lbl.pack(anchor="w", padx=20, pady=(6, 6))

        step2_frame = ctk.CTkFrame(card, fg_color="transparent")
        step2_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.entry_csv = ctk.CTkEntry(
            step2_frame,
            placeholder_text="保存先CSVファイルを指定してください",
            font=ctk.CTkFont(family=self.font_family, size=15),
            height=44
        )
        self.entry_csv.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))

        btn_browse_csv = ctk.CTkButton(
            step2_frame,
            text="💾 参照...",
            command=self.browse_csv,
            font=ctk.CTkFont(family=self.font_family, size=15, weight="bold"),
            width=110,
            height=44
        )
        btn_browse_csv.pack(side=tk.RIGHT)

        # 変換実行ボタン（大型）
        self.btn_convert = ctk.CTkButton(
            self,
            text="⚡ 変換を実行する",
            command=self.run_conversion,
            font=ctk.CTkFont(family=self.font_family, size=22, weight="bold"),
            width=360,
            height=65,
            fg_color=("#87CEFA", "#2563EB"),
            hover_color=("#00BFFF", "#1D4ED8"),
            text_color=("black", "white")
        )
        self.btn_convert.pack(pady=16)

        # 完了メッセージ・連携エリア（初期非表示）
        self.result_card = ctk.CTkFrame(self, corner_radius=12, fg_color=("#E8F8F5", "#133E2B"))
        
        self.success_label = ctk.CTkLabel(
            self.result_card,
            text="",
            font=ctk.CTkFont(family=self.font_family, size=18, weight="bold"),
            text_color=("#117864", "#4ADE80")
        )
        self.success_label.pack(pady=(14, 8), padx=20)

        self.btn_goto_selector = ctk.CTkButton(
            self.result_card,
            text="🎯 この名簿で学生指名を開始する",
            command=self.goto_selector,
            font=ctk.CTkFont(family=self.font_family, size=18, weight="bold"),
            width=360,
            height=54,
            fg_color=("#2ECC71", "#16A34A"),
            hover_color=("#27AE60", "#15803D"),
            text_color="white"
        )
        self.btn_goto_selector.pack(pady=(4, 16))

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="変換元のExcelファイルを選択",
            filetypes=[("Excelファイル", "*.xlsx *.xlsm *.xltx *.xltm"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.input_excel_path = filename
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, filename)

            # 保存先CSVの初期値を自動提案
            dir_name = os.path.dirname(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            default_csv_name = f"演習投影名簿_{base_name}.csv"
            suggested_csv = os.path.join(dir_name, default_csv_name)

            self.output_csv_path = suggested_csv
            self.entry_csv.delete(0, tk.END)
            self.entry_csv.insert(0, suggested_csv)

            # 結果カードを隠す
            self.result_card.pack_forget()

    def browse_csv(self):
        initial_file = os.path.basename(self.entry_csv.get()) or "演習投影名簿_出力.csv"
        initial_dir = os.path.dirname(self.entry_csv.get()) or None

        filename = filedialog.asksaveasfilename(
            title="変換したCSVファイルの保存先を選択",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if filename:
            self.output_csv_path = filename
            self.entry_csv.delete(0, tk.END)
            self.entry_csv.insert(0, filename)
            self.result_card.pack_forget()

    def run_conversion(self):
        in_excel = self.entry_excel.get().strip()
        out_csv = self.entry_csv.get().strip()

        if not in_excel:
            messagebox.showwarning("警告", "変換元のExcelファイルを選択してください。")
            return
        if not os.path.exists(in_excel):
            messagebox.showerror("エラー", f"指定されたExcelファイルが存在しません:\n{in_excel}")
            return
        if not out_csv:
            messagebox.showwarning("警告", "保存先CSVファイルを指定してください。")
            return

        try:
            res = convert_excel_to_csv(in_excel, out_csv)
            count = res["total_extracted"]

            # 成功表示
            self.success_label.configure(
                text=f"✅ 変換が完了しました！（全 {count}名 抽出完了）"
            )
            self.result_card.pack(padx=30, fill="x", pady=10)
            messagebox.showinfo("変換成功", f"{count}名の学生データをCSVに変換・保存しました。\n保存先: {out_csv}")
        except RosterConversionError as rce:
            self.result_card.pack_forget()
            messagebox.showerror("変換エラー", str(rce))
        except Exception as e:
            self.result_card.pack_forget()
            messagebox.showerror("予期しないエラー", f"エラーが発生しました:\n{e}")

    def goto_selector(self):
        out_csv = self.entry_csv.get().strip()
        if self.on_start_lottery and out_csv and os.path.exists(out_csv):
            self.on_start_lottery(out_csv)
