import tkinter as tk
from tkinter import messagebox, filedialog
import csv
import random
import os

class StudentSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("演習投影学生指名プログラム")
        self.root.geometry("500x500")
        
        self.filename = None
        self.data = []
        self.fieldnames = []
        self.selected_index = None
        self.selected_student = None

        self.create_wait_screen()
        self.create_result_screen()
        
        self.show_wait_screen()

    def create_wait_screen(self):
        self.wait_frame = tk.Frame(self.root, padx=20, pady=40)
        
        title_label = tk.Label(self.wait_frame, text="演習投影学生指名", font=("Arial", 24, "bold"))
        title_label.pack(pady=20)

        self.file_label = tk.Label(self.wait_frame, text="対象ファイル: 未選択", wraplength=450, font=("Arial", 12))
        self.file_label.pack(pady=15)

        select_file_btn = tk.Button(self.wait_frame, text="CSVファイルを選択", command=self.select_file, font=("Arial", 16, "bold"), width=16, height=2, bg="lightgreen")
        select_file_btn.pack(pady=10)

        self.draw_btn = tk.Button(self.wait_frame, text="抽選開始", command=self.draw_lottery, font=("Arial", 16, "bold"), bg="lightblue", width=16, height=2)
        self.draw_btn.pack(pady=30)

        exit_btn = tk.Button(self.wait_frame, text="プログラムを終了", command=self.confirm_exit, font=("Arial", 10), width=15, bg="lightcoral")
        exit_btn.pack(pady=10)

    def confirm_exit(self):
        if messagebox.askyesno("確認", "プログラムを終了しますか？"):
            self.root.destroy()

    def create_result_screen(self):
        self.result_frame = tk.Frame(self.root, padx=20, pady=40)
        
        self.student_label = tk.Label(self.result_frame, text="", font=("Arial", 20, "bold"))
        self.student_label.pack(pady=30)
        
        prompt_label = tk.Label(self.result_frame, text="出席状況を選択してください", font=("Arial", 16))
        prompt_label.pack(pady=20)

        btn_frame = tk.Frame(self.result_frame)
        btn_frame.pack(pady=30)

        button_font = ("Arial", 12, "bold")
        btn_width = 9
        btn_height = 3

        ok_btn = tk.Button(btn_frame, text="出席", command=lambda: self.save_result("○"), font=button_font, width=btn_width, height=btn_height, bg="lightgreen")
        ok_btn.pack(side=tk.LEFT, padx=12)

        ng_btn = tk.Button(btn_frame, text="欠席", command=lambda: self.save_result("×"), font=button_font, width=btn_width, height=btn_height, bg="lightcoral")
        ng_btn.pack(side=tk.LEFT, padx=12)

        skip_btn = tk.Button(btn_frame, text="記入なし\n(戻る)", command=self.show_wait_screen, font=button_font, width=btn_width, height=btn_height, bg="lightgray")
        skip_btn.pack(side=tk.LEFT, padx=12)

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
            self.file_label.config(text=f"対象ファイル:\n{os.path.basename(filename)}")

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

        self.student_label.config(text=f"学籍番号: {self.selected_student}\n氏名: {selected_name}\nカナ: {selected_kana}")
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
    root = tk.Tk()
    app = StudentSelectorApp(root)
    root.mainloop()
