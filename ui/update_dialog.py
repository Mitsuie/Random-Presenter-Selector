"""
アプリケーション更新ダイアログ
新バージョンのリリースノート表示、インストーラーの非同期ダウンロード、
進捗表示、および自動インストール・再起動実行を担当します。
"""
import os
import sys
import threading
import webbrowser
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional
from core.app_updater import UpdateInfo, download_installer, launch_installer_and_exit


class UpdateDialog:
    """新バージョン案内・ダウンロード・更新適用ダイアログ"""

    def __init__(self, parent: tk.Widget, update_info: UpdateInfo, current_version: str):
        self.parent = parent
        self.update_info = update_info
        self.current_version = current_version
        self.cancel_event = threading.Event()
        self.download_thread: Optional[threading.Thread] = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("アプリケーション アップデート")
        self.dialog.geometry("680x560")
        self.dialog.minsize(560, 440)
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 画面中央配置
        self.dialog.update_idletasks()
        w, h = 680, 560
        px = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.dialog.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")

        self._build_ui()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        # 1. 最下部ボタンバー (確実に表示されるよう side=BOTTOM で優先配置)
        btn_bar = tk.Frame(self.dialog)
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 16))

        self.browser_btn = tk.Button(btn_bar, text="GitHubで確認", command=self._open_browser, padx=12, pady=6)
        self.browser_btn.pack(side=tk.LEFT)

        self.close_btn = tk.Button(btn_bar, text="閉じる", command=self._on_close, padx=14, pady=6)
        self.close_btn.pack(side=tk.RIGHT, padx=(8, 0))

        if self.update_info.installer_download_url:
            self.update_btn = tk.Button(
                btn_bar,
                text="今すぐアップデート (自動インストール)",
                command=self._start_download,
                bg="#0969da",
                fg="#ffffff",
                font=("", 9, "bold"),
                padx=16,
                pady=6
            )
            self.update_btn.pack(side=tk.RIGHT)
        else:
            self.update_btn = None

        # 2. ダウンロード進捗領域
        self.progress_frame = tk.Frame(self.dialog)
        self.progress_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 10))

        self.progress_lbl = tk.Label(self.progress_frame, text="", fg="#57606a", anchor="w")
        self.progress_lbl.pack(fill=tk.X, pady=(0, 2))

        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X)
        self.progress_frame.pack_forget()  # 開始まで非表示

        # 3. メインコンテンツ（ヘッダー ＆ リリースノート）
        container = tk.Frame(self.dialog)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=(16, 10))

        title_lbl = tk.Label(container, text="🚀 新バージョンが利用可能です！", font=("", 13, "bold"), anchor="w")
        title_lbl.pack(fill=tk.X)

        ver_text = f"現在のバージョン: v{self.current_version}  ➔  最新バージョン: v{self.update_info.version}"
        if self.update_info.installer_size > 0:
            ver_text += f" ({self.update_info.installer_size / (1024*1024):.1f} MB)"
        tk.Label(container, text=ver_text, fg="#0969da", font=("", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(4, 10))

        tk.Label(container, text="リリースノート:", font=("", 9, "bold"), anchor="w").pack(fill=tk.X)

        text_frame = tk.Frame(container)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        scroll = ttk.Scrollbar(text_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.notes_text = tk.Text(text_frame, wrap="word", height=8, yscrollcommand=scroll.set, padx=8, pady=8)
        self.notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=self.notes_text.yview)

        self.notes_text.insert("1.0", self.update_info.release_notes.strip() or "詳細はGitHubをご覧ください。")
        self.notes_text.configure(state="disabled")

    def _open_browser(self):
        if self.update_info.release_url:
            webbrowser.open(self.update_info.release_url)

    def _start_download(self):
        if not self.update_btn:
            return
        self.update_btn.configure(state="disabled", text="ダウンロード中...")
        self.close_btn.configure(text="キャンセル")
        self.progress_frame.pack(fill=tk.X, padx=20, pady=(0, 10))
        self.progress_bar["value"] = 0

        def _worker():
            try:
                def _on_prog(dl, total):
                    if total > 0:
                        pct = (dl / total) * 100
                        msg = f"ダウンロード中: {dl/(1024*1024):.1f} MB / {total/(1024*1024):.1f} MB ({pct:.1f}%)"
                        self.dialog.after(0, lambda: self._update_prog_ui(pct, msg))

                path = download_installer(
                    self.update_info.installer_download_url,
                    target_filename=self.update_info.installer_name,
                    progress_callback=_on_prog,
                    cancel_event=self.cancel_event
                )
                self.dialog.after(0, lambda: self._on_download_complete(path))
            except InterruptedError:
                self.dialog.after(0, self._on_cancelled)
            except Exception as e:
                self.dialog.after(0, lambda: self._on_failed(str(e)))

        self.download_thread = threading.Thread(target=_worker, daemon=True)
        self.download_thread.start()

    def _update_prog_ui(self, pct, msg):
        self.progress_bar["value"] = pct
        self.progress_lbl.configure(text=msg)

    def _on_download_complete(self, path: str):
        file_name = os.path.basename(path)
        self.progress_lbl.configure(text=f"✓ ダウンロード完了: {file_name}")
        self.progress_bar["value"] = 100

        self.close_btn.configure(text="後で (閉じる)")
        if self.update_btn:
            self.update_btn.configure(state="normal", text="今すぐインストール", command=lambda: launch_installer_and_exit(path))
        self.browser_btn.configure(text="保存先フォルダを開く", command=lambda: self._open_folder(path))

        if messagebox.askyesno(
            "ダウンロード完了",
            f"インストーラーのダウンロードが完了しました。\n\n"
            "今すぐアプリを終了してインストーラーを起動しますか？\n\n"
            "・「はい」: アプリを終了し、インストーラーを起動して更新\n"
            "・「いいえ」: 起動せず、後でダイアログから実行",
            parent=self.dialog
        ):
            launch_installer_and_exit(path)

    def _open_folder(self, file_path: str):
        if sys.platform == "win32":
            subprocess.Popen(f'explorer /select,"{os.path.abspath(file_path)}"')
        else:
            webbrowser.open(os.path.dirname(file_path))

    def _on_failed(self, err: str):
        messagebox.showerror("エラー", f"ダウンロードに失敗しました:\n{err}", parent=self.dialog)
        if self.update_btn:
            self.update_btn.configure(state="normal", text="今すぐアップデート")
        self.close_btn.configure(text="閉じる")
        self.progress_frame.pack_forget()

    def _on_cancelled(self):
        if self.update_btn:
            self.update_btn.configure(state="normal", text="今すぐアップデート")
        self.close_btn.configure(text="閉じる")
        self.progress_frame.pack_forget()

    def _on_close(self):
        if self.download_thread and self.download_thread.is_alive():
            if messagebox.askyesno("確認", "ダウンロードを中止して閉じますか？", parent=self.dialog):
                self.cancel_event.set()
                self.dialog.destroy()
        else:
            self.dialog.destroy()
