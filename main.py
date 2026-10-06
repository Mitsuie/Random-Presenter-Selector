"""
演習投影学生指名プログラム 統合メインアプリケーション
名簿フォーマット変換（Excel -> CSV）および学生指名・出席記録を一元管理します。
"""

import logging
import sys
import tkinter as tk
from tkinter import messagebox
from core.app_logging import setup_logging
from core.constants import APP_VERSION
from ui.main_window import MainWindow

def main():
    setup_logging()
    logging.info("起動 (v%s)", APP_VERSION)
    try:
        app = MainWindow()
        app.mainloop()
    except Exception as e:
        # 予期しないエラー時のダイアログ
        logging.exception("起動中に予期しないエラーが発生しました")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("予期しないエラー", f"アプリケーション実行中にエラーが発生しました:\n{e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
