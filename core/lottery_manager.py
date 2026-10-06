import os
import csv
import random
from typing import Optional, Dict, List, Any
from core.utils import atomic_write_text

REQUIRED_CSV_COLUMNS = ['学籍番号', '投影実施可否']

# 読み込みを試すエンコーディング（Excelで保存し直したCSVは cp932 になる）
CSV_ENCODINGS = ['utf-8-sig', 'cp932']

class LotteryError(Exception):
    """学生指名ロジックのエラー基底クラス"""
    pass

class LotteryManager:
    """CSV名簿の読み込み・統計・抽選・出席記録を管理するクラス"""

    def __init__(self):
        self.filename: Optional[str] = None
        self.encoding: str = CSV_ENCODINGS[0]
        self.fieldnames: List[str] = []
        self.data: List[Dict[str, Any]] = []
        self.current_selected_index: Optional[int] = None
        self.last_record: Optional[Dict[str, str]] = None
        self._loaded_mtime: Optional[float] = None

    def load_csv(self, filename: str) -> Dict[str, int]:
        """
        CSVファイルを読み込み、検証を行う。

        Args:
            filename: 読み込むCSVのファイルパス

        Returns:
            Dict[str, int]: 読み込み後の統計情報 (total, pending, done)

        Raises:
            FileNotFoundError: ファイルが存在しない場合
            LotteryError: 必須列が不足している、またはパースエラーの場合
        """
        if not os.path.exists(filename):
            raise FileNotFoundError(f"ファイルが見つかりません: {filename}")

        fieldnames, records, encoding = self._read_csv(filename)

        self.filename = filename
        self.encoding = encoding
        self.fieldnames = fieldnames
        self.data = records
        self.current_selected_index = None
        self.last_record = None
        self._loaded_mtime = os.path.getmtime(filename)

        return self.get_statistics()

    def _read_csv(self, filename: str):
        """CSVを対応エンコーディングで順に試して読み込む"""
        last_error: Optional[Exception] = None
        for encoding in CSV_ENCODINGS:
            try:
                with open(filename, mode='r', encoding=encoding, newline='') as f:
                    reader = csv.DictReader(f)
                    fieldnames = reader.fieldnames
                    if not fieldnames or not all(col in fieldnames for col in REQUIRED_CSV_COLUMNS):
                        raise LotteryError(
                            f"CSVファイルに必要な列（{', '.join(REQUIRED_CSV_COLUMNS)}）が存在しません。"
                        )
                    records = [dict(row) for row in reader]
                return list(fieldnames), records, encoding
            except UnicodeDecodeError as e:
                last_error = e
                continue
            except LotteryError:
                raise
            except Exception as e:
                raise LotteryError(f"CSVファイルの読み込みに失敗しました: {e}") from e
        raise LotteryError(f"CSVファイルの文字コードを判別できませんでした: {last_error}")

    def get_statistics(self) -> Dict[str, int]:
        """
        現在のデータの統計情報を取得する。

        Returns:
            Dict[str, int]: {"total": 全人数, "pending": 未投影人数, "done": 実施済人数}
        """
        total = len(self.data)
        pending = 0
        done = 0

        for row in self.data:
            val = str(row.get('投影実施可否', '') or '').strip()
            if not val:
                pending += 1
            else:
                done += 1

        return {
            "total": total,
            "pending": pending,
            "done": done
        }

    def get_empty_indices(self) -> List[int]:
        """投影実施可否が空欄（未投影）のインデックスリストを取得する"""
        indices = []
        for i, row in enumerate(self.data):
            val = str(row.get('投影実施可否', '') or '').strip()
            if not val:
                indices.append(i)
        return indices

    def draw_student(self) -> Optional[Dict[str, Any]]:
        """
        未投影の学生からランダムに1名選出する。

        Returns:
            Optional[Dict[str, Any]]:
                選出された学生情報 {
                    "index": int,
                    "student_id": str,
                    "name": str,
                    "kana": str
                }
                未投影の学生がいない場合は None
        """
        empty_indices = self.get_empty_indices()
        if not empty_indices:
            self.current_selected_index = None
            return None

        chosen_idx = random.choice(empty_indices)
        self.current_selected_index = chosen_idx
        row = self.data[chosen_idx]

        return {
            "index": chosen_idx,
            "student_id": row.get('学籍番号', '不明'),
            "name": row.get('学生氏名', '不明'),
            "kana": row.get('学生氏名＿カナ', '不明')
        }

    def save_result(self, result: str, index: Optional[int] = None) -> bool:
        """
        選出された学生の出席状況（○, × 等）を保存し、CSVファイルに上書きする。

        保存前にCSVが外部（Excel等）で変更されていた場合は読み込み直してから記録し、
        外部での変更を失わないようにする。

        Args:
            result: 記録する値（"○", "×" 等）
            index: 対象の学生インデックス（Noneの場合は現在の選出学生）

        Returns:
            bool: 外部で変更されたCSVを読み込み直した場合は True

        Raises:
            LotteryError: 対象学生が選択されていない場合、またはCSV書き込みに失敗した場合
            PermissionError: ファイルがExcel等で開かれていて書き込めない場合
        """
        target_idx = index if index is not None else self.current_selected_index
        if target_idx is None or target_idx >= len(self.data):
            raise LotteryError("結果を記録する対象の学生が選ばれていません。")

        if not self.filename:
            raise LotteryError("対象のCSVファイルが設定されていません。")

        student_id = self.data[target_idx].get('学籍番号')
        reloaded = False
        if self._is_modified_externally():
            target_idx = self._reload_and_find(student_id)
            reloaded = True
            if self._value(target_idx):
                self.current_selected_index = None
                raise LotteryError(
                    f"CSVが外部で変更され、学籍番号 {student_id} の学生にはすでに記録があります。記録は保存されていません。"
                )

        self._set_value_and_save(target_idx, result)

        self.current_selected_index = None
        self.last_record = {
            "student_id": student_id,
            "name": self.data[target_idx].get('学生氏名', ''),
            "value": result,
        }
        return reloaded

    def undo_last(self) -> Dict[str, str]:
        """
        直前に記録した1件を取り消し（空欄に戻し）、CSVに保存する。

        Returns:
            Dict[str, str]: 取り消した記録 {"student_id", "name", "value"}

        Raises:
            LotteryError: 取り消せる記録がない、または外部で変更されていて取り消せない場合
            PermissionError: ファイルがExcel等で開かれていて書き込めない場合
        """
        record = self.last_record
        if not record or not self.filename:
            raise LotteryError("取り消せる記録がありません。")

        student_id = record["student_id"]
        if self._is_modified_externally():
            target_idx = self._reload_and_find(student_id)
        else:
            target_idx = self._find_index(student_id)
            if target_idx is None:
                raise LotteryError(f"学籍番号 {student_id} の学生が見つかりません。")

        if self._value(target_idx) != record["value"]:
            self.last_record = None
            raise LotteryError(
                f"CSVが外部で変更され、学籍番号 {student_id} の記録が変わっているため取り消せません。"
            )

        self._set_value_and_save(target_idx, '')
        self.last_record = None
        return record

    def _value(self, index: int) -> str:
        return str(self.data[index].get('投影実施可否', '') or '').strip()

    def _find_index(self, student_id: Optional[str]) -> Optional[int]:
        return next((i for i, r in enumerate(self.data) if r.get('学籍番号') == student_id), None)

    def _set_value_and_save(self, index: int, value: str) -> None:
        """メモリ上の記録を更新して保存する（書き込みに失敗したら元に戻す）"""
        row = self.data[index]
        old_value = row.get('投影実施可否', '')
        row['投影実施可否'] = value
        try:
            self._write_csv()
        except BaseException:
            row['投影実施可否'] = old_value
            raise

    def _is_modified_externally(self) -> bool:
        try:
            return os.path.getmtime(self.filename) != self._loaded_mtime
        except OSError:
            return False

    def _reload_and_find(self, student_id: Optional[str]) -> int:
        """
        CSVを読み込み直し、対象学生の新しいインデックスを学籍番号で探して返す。
        """
        fieldnames, records, encoding = self._read_csv(self.filename)
        self.fieldnames = fieldnames
        self.data = records
        self.encoding = encoding
        self._loaded_mtime = os.path.getmtime(self.filename)

        new_idx = self._find_index(student_id)
        if new_idx is None:
            self.current_selected_index = None
            raise LotteryError(
                f"CSVが外部で変更され、学籍番号 {student_id} の学生が見つからなくなりました。記録は保存されていません。"
            )
        return new_idx

    def _write_csv(self) -> None:
        def _write(f):
            writer = csv.DictWriter(f, fieldnames=self.fieldnames)
            writer.writeheader()
            writer.writerows(self.data)

        try:
            atomic_write_text(self.filename, _write, encoding=self.encoding)
        except PermissionError as pe:
            raise PermissionError(
                f"ファイルの書き込みに失敗しました。ファイルがExcel等で開かれている可能性があります。\n"
                f"ファイルを閉じてから再度お試しください。\n詳細: {pe}"
            ) from pe
        except Exception as e:
            raise LotteryError(f"ファイル書き込み中にエラーが発生しました: {e}") from e

        self._loaded_mtime = os.path.getmtime(self.filename)

    def clear_selection(self) -> None:
        """現在の選出状態をクリアする"""
        self.current_selected_index = None
