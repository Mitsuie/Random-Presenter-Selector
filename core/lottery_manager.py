import os
import csv
import random
from typing import Optional, Dict, List, Any

REQUIRED_CSV_COLUMNS = ['学籍番号', '投影実施可否']

class LotteryError(Exception):
    """学生指名ロジックのエラー基底クラス"""
    pass

class LotteryManager:
    """CSV名簿の読み込み・統計・抽選・出席記録を管理するクラス"""

    def __init__(self):
        self.filename: Optional[str] = None
        self.fieldnames: List[str] = []
        self.data: List[Dict[str, Any]] = []
        self.current_selected_index: Optional[int] = None

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

        records = []
        try:
            with open(filename, mode='r', encoding='utf-8-sig', newline='') as f:
                reader = csv.DictReader(f)
                fieldnames = reader.fieldnames
                if not fieldnames or not all(col in fieldnames for col in REQUIRED_CSV_COLUMNS):
                    raise LotteryError(
                        f"CSVファイルに必要な列（{', '.join(REQUIRED_CSV_COLUMNS)}）が存在しません。"
                    )
                for row in reader:
                    records.append(dict(row))
        except LotteryError:
            raise
        except Exception as e:
            raise LotteryError(f"CSVファイルの読み込みに失敗しました: {e}") from e

        self.filename = filename
        self.fieldnames = list(fieldnames)
        self.data = records
        self.current_selected_index = None

        return self.get_statistics()

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

    def save_result(self, result: str, index: Optional[int] = None) -> None:
        """
        選出された学生の出席状況（○, × 等）を保存し、CSVファイルに上書きする。

        Args:
            result: 記録する値（"○", "×" 等）
            index: 対象の学生インデックス（Noneの場合は現在の選出学生）

        Raises:
            LotteryError: 対象学生が選択されていない場合、またはCSV書き込みに失敗した場合
        """
        target_idx = index if index is not None else self.current_selected_index
        if target_idx is None or target_idx >= len(self.data):
            raise LotteryError("結果を記録する対象の学生が選ばれていません。")

        if not self.filename:
            raise LotteryError("対象のCSVファイルが設定されていません。")

        # メモリ上のデータを更新
        self.data[target_idx]['投影実施可否'] = result

        # CSVに上書き保存
        try:
            with open(self.filename, mode='w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=self.fieldnames)
                writer.writeheader()
                writer.writerows(self.data)
        except PermissionError as pe:
            raise PermissionError(
                f"ファイルの書き込みに失敗しました。ファイルがExcel等で開かれている可能性があります。\n"
                f"ファイルを閉じてから再度お試しください。\n詳細: {pe}"
            ) from pe
        except Exception as e:
            raise LotteryError(f"ファイル書き込み中にエラーが発生しました: {e}") from e

        self.current_selected_index = None

    def clear_selection(self) -> None:
        """現在の選出状態をクリアする"""
        self.current_selected_index = None
