"""
アプリケーション更新ダイアログ
新バージョンのリリースノート表示、インストーラーの非同期ダウンロード、
進捗表示、および自動インストール・再起動実行を担当します。
"""
import logging
import os
import sys
import threading
import webbrowser
import subprocess
import tkinter as tk
from tkinter import messagebox
from typing import Optional
import customtkinter as ctk
from core.app_updater import UpdateInfo, download_installer, launch_installer
from ui.font_config import get_font_family

logger = logging.getLogger(__name__)


class UpdateDialog:
    """新バージョン案内・ダウンロード・更新適用ダイアログ"""

    WIDTH = 680
    HEIGHT = 560

    def __init__(self, parent: ctk.CTk, update_info: UpdateInfo, current_version: str):
        self.parent = parent
        self.update_info = update_info
        self.current_version = current_version
        self.cancel_event = threading.Event()
        self.download_thread: Optional[threading.Thread] = None
        self.font_family = get_font_family()
        self._last_pct = -1

        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title("アプリケーション アップデート")
        self.dialog.minsize(560, 440)
        self.dialog.transient(parent)

        # 画面中央配置
        px = parent.winfo_rootx() + (parent.winfo_width() - self.WIDTH) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.HEIGHT) // 2
        self.dialog.geometry(f"{self.WIDTH}x{self.HEIGHT}+{max(0, px)}+{max(0, py)}")

        self._build_ui()
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
        # CTkToplevel は表示直後に grab できない場合があるため、少し待ってからモーダル化する
        self.dialog.after(100, self._make_modal)

    def _make_modal(self):
        try:
            self.dialog.grab_set()
            self.dialog.focus_set()
        except Exception:
            pass

    def _font(self, size: int, bold: bool = False) -> ctk.CTkFont:
        return ctk.CTkFont(family=self.font_family, size=size, weight="bold" if bold else "normal")

    def _build_ui(self) -> None:
        # 1. 最下部ボタンバー (確実に表示されるよう side=BOTTOM で優先配置)
        btn_bar = ctk.CTkFrame(self.dialog, fg_color="transparent")
        btn_bar.pack(side="bottom", fill="x", padx=20, pady=(0, 16))

        neutral = dict(fg_color=("gray80", "gray28"), hover_color=("gray70", "gray38"), text_color=("black", "white"))

        self.browser_btn = ctk.CTkButton(
            btn_bar, text="GitHubで確認", command=self._open_browser, font=self._font(14), **neutral
        )
        self.browser_btn.pack(side="left")

        self.close_btn = ctk.CTkButton(
            btn_bar, text="閉じる", command=self._on_close, font=self._font(14), width=110, **neutral
        )
        self.close_btn.pack(side="right", padx=(8, 0))

        if self.update_info.installer_download_url:
            self.update_btn = ctk.CTkButton(
                btn_bar,
                text="今すぐアップデート (自動インストール)",
                command=self._start_download,
                font=self._font(14, bold=True),
                fg_color=("#0969da", "#1f6feb"),
                hover_color=("#054da7", "#1158c7"),
                text_color="white"
            )
            self.update_btn.pack(side="right")
        else:
            self.update_btn = None

        # 2. ダウンロード進捗領域
        self.progress_frame = ctk.CTkFrame(self.dialog, fg_color="transparent")

        self.progress_lbl = ctk.CTkLabel(
            self.progress_frame, text="", anchor="w", font=self._font(13), text_color=("gray35", "gray70")
        )
        self.progress_lbl.pack(fill="x", pady=(0, 2))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x")

        # 3. メインコンテンツ（ヘッダー ＆ リリースノート）
        container = ctk.CTkFrame(self.dialog, fg_color="transparent")
        container.pack(side="top", fill="both", expand=True, padx=20, pady=(16, 10))

        ctk.CTkLabel(container, text="🚀 新バージョンが利用可能です！", font=self._font(18, bold=True), anchor="w").pack(fill="x")

        ver_text = f"現在のバージョン: v{self.current_version}  ➔  最新バージョン: v{self.update_info.version}"
        if self.update_info.installer_size > 0:
            ver_text += f" ({self.update_info.installer_size / (1024*1024):.1f} MB)"
        ctk.CTkLabel(
            container, text=ver_text, font=self._font(14, bold=True), text_color=("#0969da", "#58a6ff"), anchor="w"
        ).pack(fill="x", pady=(4, 10))

        ctk.CTkLabel(container, text="リリースノート:", font=self._font(13, bold=True), anchor="w").pack(fill="x")

        self.notes_text = ctk.CTkTextbox(container, wrap="word", font=self._font(13))
        self.notes_text.pack(fill="both", expand=True, pady=(4, 0))
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
        self.progress_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 10))
        self.progress_bar.set(0)
        self._last_pct = -1

        def _worker():
            try:
                def _on_prog(dl, total):
                    if total <= 0:
                        return
                    pct = int(dl * 100 / total)
                    # 進捗表示の更新は1%ごとに間引く（イベントキューを溢れさせない）
                    if pct == self._last_pct:
                        return
                    self._last_pct = pct
                    msg = f"ダウンロード中: {dl/(1024*1024):.1f} MB / {total/(1024*1024):.1f} MB ({pct}%)"
                    self._post(lambda: self._update_prog_ui(pct, msg))

                path = download_installer(
                    self.update_info.installer_download_url,
                    target_filename=self.update_info.installer_name,
                    progress_callback=_on_prog,
                    cancel_event=self.cancel_event,
                    expected_size=self.update_info.installer_size
                )
                self._post(lambda: self._on_download_complete(path))
            except InterruptedError:
                self._post(self._on_cancelled)
            except Exception as e:
                logger.warning("更新のダウンロードに失敗しました", exc_info=True)
                # except を抜けると e は削除されるため、メインスレッドで使う値を先に取り出す
                err = str(e)
                self._post(lambda: self._on_failed(err))

        self.download_thread = threading.Thread(target=_worker, daemon=True)
        self.download_thread.start()

    def _post(self, func):
        """ワーカースレッドからメインスレッドへ処理を渡す（ダイアログが閉じられていたら何もしない）"""
        try:
            self.dialog.after(0, func)
        except (RuntimeError, tk.TclError):
            pass

    def _update_prog_ui(self, pct, msg):
        self.progress_bar.set(pct / 100)
        self.progress_lbl.configure(text=msg)

    def _on_download_complete(self, path: str):
        file_name = os.path.basename(path)
        self.progress_lbl.configure(text=f"✓ ダウンロード完了: {file_name}")
        self.progress_bar.set(1)

        self.close_btn.configure(text="後で (閉じる)")
        if self.update_btn:
            self.update_btn.configure(state="normal", text="今すぐインストール", command=lambda: self._install(path))
        self.browser_btn.configure(text="保存先フォルダを開く", command=lambda: self._open_folder(path))

        if messagebox.askyesno(
            "ダウンロード完了",
            f"インストーラーのダウンロードが完了しました。\n\n"
            "今すぐアプリを終了してインストーラーを起動しますか？\n\n"
            "・「はい」: アプリを終了し、インストーラーを起動して更新\n"
            "・「いいえ」: 起動せず、後でダイアログから実行",
            parent=self.dialog
        ):
            self._install(path)

    def _install(self, path: str):
        """インストーラーを起動し、アプリを終了する"""
        try:
            launch_installer(path)
        except Exception as e:
            logger.exception("インストーラーを起動できませんでした")
            messagebox.showerror("エラー", f"インストーラーを起動できませんでした:\n{e}", parent=self.dialog)
            return
        logger.info("インストーラーを起動して終了します: %s", path)
        self.parent.destroy()

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
