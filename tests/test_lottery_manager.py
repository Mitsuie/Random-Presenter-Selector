import os
import csv
import tempfile
import unittest
from core.lottery_manager import LotteryManager, LotteryError

class TestLotteryManager(unittest.TestCase):
    """学生指名ロジック (core.lottery_manager) の単体テスト"""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.csv_path = os.path.join(self.temp_dir.name, "roster.csv")
        self.manager = LotteryManager()

        # テスト用初期CSVを作成 (総数3名: 未投影2名, 実施済1名)
        self.initial_data = [
            ["学籍番号", "学生氏名", "学生氏名＿カナ", "投影実施可否"],
            ["K001", "山田 太郎", "ヤマダ タロウ", ""],
            ["K002", "佐藤 花子", "サトウ ハナコ", "○"],
            ["K003", "鈴木 一郎", "スズキ イチロウ", ""],
        ]
        self._write_csv(self.csv_path, self.initial_data)

    def tearDown(self):
        self.temp_dir.cleanup()

    def _write_csv(self, path, rows):
        with open(path, mode="w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(rows)

    def test_load_csv_success(self):
        """正常系: CSV読み込みと統計集計（total, pending, done）が正しいこと"""
        stats = self.manager.load_csv(self.csv_path)
        self.assertEqual(stats["total"], 3)
        self.assertEqual(stats["pending"], 2)
        self.assertEqual(stats["done"], 1)

    def test_load_csv_file_not_found(self):
        """異常系: 存在しないCSVファイルを読み込んだ場合に FileNotFoundError が送出されること"""
        not_exist = os.path.join(self.temp_dir.name, "none.csv")
        with self.assertRaises(FileNotFoundError):
            self.manager.load_csv(not_exist)

    def test_load_csv_missing_column(self):
        """異常系: 必須列が不足している場合に LotteryError が送出されること"""
        invalid_csv = os.path.join(self.temp_dir.name, "invalid.csv")
        self._write_csv(invalid_csv, [["名前", "カナ"], ["田中", "タナカ"]])
        with self.assertRaises(LotteryError):
            self.manager.load_csv(invalid_csv)

    def test_get_empty_indices(self):
        """未投影（投影実施可否が空欄）の行インデックスが正しく抽出されること"""
        self.manager.load_csv(self.csv_path)
        empty_indices = self.manager.get_empty_indices()
        self.assertEqual(empty_indices, [0, 2])

    def test_draw_student_success(self):
        """正常系: 未投影の学生から正しく1名が選出されること"""
        self.manager.load_csv(self.csv_path)
        student = self.manager.draw_student()

        self.assertIsNotNone(student)
        self.assertIn(student["index"], [0, 2])
        self.assertIn(student["student_id"], ["K001", "K003"])
        self.assertEqual(self.manager.current_selected_index, student["index"])

    def test_draw_student_when_all_done(self):
        """全員実施済みの場合に draw_student が None を返すこと"""
        # 全員実施済みのCSVを作成
        done_csv = os.path.join(self.temp_dir.name, "all_done.csv")
        self._write_csv(done_csv, [
            ["学籍番号", "学生氏名", "学生氏名＿カナ", "投影実施可否"],
            ["K001", "山田 太郎", "ヤマダ タロウ", "○"],
            ["K002", "佐藤 花子", "サトウ ハナコ", "×"],
        ])
        self.manager.load_csv(done_csv)
        student = self.manager.draw_student()

        self.assertIsNone(student)
        self.assertIsNone(self.manager.current_selected_index)

    def test_save_result_success(self):
        """正常系: 選出された学生に対して出席・欠席が記録され、CSVファイルに上書き保存されること"""
        self.manager.load_csv(self.csv_path)
        student = self.manager.draw_student()
        chosen_idx = student["index"]

        # 結果を「○」で保存
        self.manager.save_result("○")

        # 選択状態がリセットされること
        self.assertIsNone(self.manager.current_selected_index)

        # 統計値が更新されること (done: 1 -> 2, pending: 2 -> 1)
        stats = self.manager.get_statistics()
        self.assertEqual(stats["done"], 2)
        self.assertEqual(stats["pending"], 1)

        # CSVファイルの実体にも反映されていることを検証
        with open(self.csv_path, mode="r", encoding="utf-8-sig", newline="") as f:
            reader = list(csv.DictReader(f))
        self.assertEqual(reader[chosen_idx]["投影実施可否"], "○")

    def test_save_result_without_selection(self):
        """異常系: 学生が選出されていない状態で save_result を呼び出した場合に LotteryError が送出されること"""
        self.manager.load_csv(self.csv_path)
        with self.assertRaises(LotteryError):
            self.manager.save_result("○")

    def test_clear_selection(self):
        """clear_selection を呼び出すと現在の選出インデックスがリセットされること"""
        self.manager.load_csv(self.csv_path)
        self.manager.draw_student()
        self.assertIsNotNone(self.manager.current_selected_index)

        self.manager.clear_selection()
        self.assertIsNone(self.manager.current_selected_index)

if __name__ == "__main__":
    unittest.main()
