import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox, filedialog
import csv
import random
import os

# 外観の設定
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

BASE_FONT_FAMILY = "MS Gothic"

class StudentSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("演習投影学生指名プログラム")
        self.root.geometry("500x500")
        self.root.resizable(False, False)

        self.filename = None
        self.data = []
        self.fieldnames = []
        self.selected_index = None
        self.selected_student = None

        self.create_wait_screen()
        self.create_result_screen()
        
        self.show_wait_screen()

    def create_wait_screen(self):
        self.wait_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        
        title_label = ctk.CTkLabel(self.wait_frame, text="演習投影学生指名プログラム", font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=28, weight="bold"))
        title_label.pack(pady=(40, 20))

        self.file_label = ctk.CTkLabel(self.wait_frame, text="対象ファイル:\n未選択", font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=16, weight="bold"), wraplength=500)
        self.file_label.pack(pady=30)

        # ファイル選択ボタン
        select_btn_color = ("#98FB98", "#3CB371") # ライト/ダーク用 (PaleGreen / MediumSeaGreen)
        select_file_btn = ctk.CTkButton(self.wait_frame, text="CSVファイルを選択", command=self.select_file, 
                                        font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=18, weight="bold"), 
                                        width=220, height=55, fg_color=select_btn_color, hover_color=("#90EE90", "#2E8B57"),
                                        text_color=("black", "white"))
        select_file_btn.pack(pady=15)

        # 抽選開始ボタン
        draw_btn_color = ("#87CEFA", "#4682B4")
        self.draw_btn = ctk.CTkButton(self.wait_frame, text="抽選開始", command=self.draw_lottery, 
                                      font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=18, weight="bold"), 
                                      fg_color=draw_btn_color, hover_color=("#00BFFF", "#4169E1"), 
                                      text_color=("black", "white"), width=220, height=55)
        self.draw_btn.pack(pady=15)

        # 終了ボタン
        exit_btn_color = ("#F08080", "#CD5C5C")
        exit_btn = ctk.CTkButton(self.wait_frame, text="プログラムを終了", command=self.confirm_exit, 
                                 font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=12, weight="bold"), width=150, height=35,
                                 fg_color=exit_btn_color, hover_color=("#CD5C5C", "#8B0000"), text_color=("black", "white"))
        exit_btn.pack(pady=40)

    def confirm_exit(self):
        if messagebox.askyesno("確認", "プログラムを終了しますか？"):
            self.root.destroy()

    def create_result_screen(self):
        self.result_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        
        self.student_label = ctk.CTkLabel(self.result_frame, text="", font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=24, weight="bold"))
        self.student_label.pack(pady=(50, 30))
        
        prompt_label = ctk.CTkLabel(self.result_frame, text="出席状況を選択してください", font=ctk.CTkFont(family=BASE_FONT_FAMILY, size=18))
        prompt_label.pack(pady=20)

        btn_frame = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        btn_frame.pack(pady=40)

        button_font = ctk.CTkFont(family=BASE_FONT_FAMILY, size=16, weight="bold")
        btn_width = 120
        btn_height = 60

        # 出席ボタン
        ok_btn_color = ("#90EE90", "#2E8B57")
        ok_btn = ctk.CTkButton(btn_frame, text="出席 (○)", command=lambda: self.save_result("○"), 
                               font=button_font, width=btn_width, height=btn_height, 
                               fg_color=ok_btn_color, hover_color=("#3CB371", "#006400"), text_color=("black", "white"))
        ok_btn.pack(side=tk.LEFT, padx=15)

        # 欠席ボタン
        ng_btn_color = ("#F08080", "#B22222")
        ng_btn = ctk.CTkButton(btn_frame, text="欠席 (×)", command=lambda: self.save_result("×"), 
                               font=button_font, width=btn_width, height=btn_height, 
                               fg_color=ng_btn_color, hover_color=("#CD5C5C", "#8B0000"), text_color=("black", "white"))
        ng_btn.pack(side=tk.LEFT, padx=15)

        # 戻るボタン
        skip_btn_color = ("#D3D3D3", "#696969")
        skip_btn = ctk.CTkButton(btn_frame, text="記入なし\n(戻る)", command=self.show_wait_screen, 
                                 font=button_font, width=btn_width, height=btn_height, 
                                 fg_color=skip_btn_color, hover_color=("#A9A9A9", "#404040"), text_color=("black", "white"))
        skip_btn.pack(side=tk.LEFT, padx=15)

    def show_wait_screen(self):
        self.result_frame.pack_forget()
        self.wait_frame.pack(fill="both", expand=True)

    def show_result_screen(self):
        self.wait_frame.pack_forget()
        self.result_frame.pack(fill="both", expand=True)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.filename = filename
            self.file_label.configure(text=f"対象ファイル:\n{os.path.basename(filename)}")

    def draw_lottery(self):
        if not self.filename:
            messagebox.showwarning("警告", "CSVファイルを選択してください。")
            return

        if not os.path.exists(self.filename):
            messagebox.showerror("エラー", f"ファイル '{self.filename}' が見つかりません。")
            return

        self.data = []
        try:
            with open(self.filename, mode='r', encoding='utf-8-sig', newline='') as f:
                reader = csv.DictReader(f)
                self.fieldnames = reader.fieldnames
                if not self.fieldnames or '学籍番号' not in self.fieldnames or '投影実施可否' not in self.fieldnames:
                    messagebox.showerror("エラー", "CSVファイルに必要な列（'学籍番号', '投影実施可否'）が存在しません。")
                    return
                for row in reader:
                    self.data.append(row)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました:\n{e}")
            return

        empty_indices = []
        for i, row in enumerate(self.data):
            val = row.get('投影実施可否', '').strip()
            if not val:
                empty_indices.append(i)

        if not empty_indices:
            messagebox.showinfo("情報", "投影実施可否が空欄の学生はいません。")
            return

        self.selected_index = random.choice(empty_indices)
        self.selected_student = self.data[self.selected_index].get('学籍番号', '不明')
        selected_name = self.data[self.selected_index].get('学生氏名', '不明')
        selected_kana = self.data[self.selected_index].get('学生氏名＿カナ', '不明')

        self.student_label.configure(text=f"学籍番号: {self.selected_student}\n氏名: {selected_name}\nカナ: {selected_kana}")
        self.show_result_screen()

    def save_result(self, result):
        if self.selected_index is None:
            return

        # メモリ上のデータを更新
        self.data[self.selected_index]['投影実施可否'] = result

        # CSVに保存
        try:
            with open(self.filename, mode='w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=self.fieldnames)
                writer.writeheader()
                writer.writerows(self.data)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの書き込み中にエラーが発生しました。\nファイルがExcel等で開かれている可能性があります。閉じてから再度お試しください。\n\n詳細: {e}")
            return # エラー時は結果画面に留まり、リトライ可能にする

        # 保存成功したら待機画面に戻る
        messagebox.showinfo("保存完了", f"{self.selected_student} の結果（{result}）を保存しました。")
        self.selected_index = None
        self.selected_student = None
        self.show_wait_screen()

if __name__ == "__main__":
    root = ctk.CTk()
    app = StudentSelectorApp(root)
    root.mainloop()
