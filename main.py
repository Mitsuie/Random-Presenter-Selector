"""
演習投影学生指名プログラム 統合メインアプリケーション
名簿フォーマット変換（Excel -> CSV）および学生指名・出席記録を一元管理します。
"""

import sys
import tkinter as tk
from tkinter import messagebox
from ui.main_window import MainWindow

def main():
    try:
        app = MainWindow()
        app.mainloop()
    except Exception as e:
        # 予期しないエラー時のダイアログ
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("予期しないエラー", f"アプリケーション実行中にエラーが発生しました:\n{e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
