import os
import csv
import tempfile
import unittest
import openpyxl
from core.roster_converter import convert_excel_to_csv, RosterConversionError

class TestRosterConverter(unittest.TestCase):
    """Excel名簿変換モジュール (core.roster_converter) の単体テスト"""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_convert_excel_to_csv_success(self):
        """正常系: 必須列を含むExcelファイルから正しいCSVが抽出・出力されること"""
        excel_path = os.path.join(self.temp_dir.name, "test_roster.xlsx")
        csv_path = os.path.join(self.temp_dir.name, "output.csv")

        # テスト用Excelファイルの作成
        wb = openpyxl.Workbook()
        ws = wb.active
        # タイトル行（無視される行）
        ws.append(["演習クラス名簿", "", "", ""])
        # ヘッダー行
        ws.append(["No.", "学籍番号", "学生氏名", "学生氏名＿カナ", "備考"])
        # データ行
        ws.append([1, "K001", "山田 太郎", "ヤマダ タロウ", ""])
        ws.append([2, "K002", "佐藤 花子", "サトウ ハナコ", ""])
        ws.append([3, "K003", "鈴木 一郎", "スズキ イチロウ", ""])
        # 空行（無視されるべき）
        ws.append(["", "", "", "", ""])
        wb.save(excel_path)

        # 変換実行
        result = convert_excel_to_csv(excel_path, csv_path)

        # 戻り値の検証
        self.assertEqual(result["total_extracted"], 3)
        self.assertEqual(result["output_path"], csv_path)
        self.assertTrue(os.path.exists(csv_path))

        # 出力されたCSVの中身の検証
        with open(csv_path, mode="r", encoding="utf-8-sig", newline="") as f:
            reader = list(csv.reader(f))

        # ヘッダー検証
        expected_header = ["学籍番号", "学生氏名", "学生氏名＿カナ", "投影実施可否"]
        self.assertEqual(reader[0], expected_header)

        # データ行検証（3件）
        self.assertEqual(len(reader), 4)  # ヘッダー + 3データ行
        self.assertEqual(reader[1], ["K001", "山田 太郎", "ヤマダ タロウ", ""])
        self.assertEqual(reader[2], ["K002", "佐藤 花子", "サトウ ハナコ", ""])
        self.assertEqual(reader[3], ["K003", "鈴木 一郎", "スズキ イチロウ", ""])

    def test_file_not_found(self):
        """異常系: 存在しないファイルパスを指定した場合に FileNotFoundError が送出されること"""
        not_exist_path = os.path.join(self.temp_dir.name, "non_existent.xlsx")
        output_csv = os.path.join(self.temp_dir.name, "output.csv")
        with self.assertRaises(FileNotFoundError):
            convert_excel_to_csv(not_exist_path, output_csv)

    def test_missing_required_column(self):
        """異常系: 必須列（学生氏名＿カナなど）が欠落している場合に RosterConversionError が送出されること"""
        excel_path = os.path.join(self.temp_dir.name, "invalid_columns.xlsx")
        output_csv = os.path.join(self.temp_dir.name, "output.csv")

        wb = openpyxl.Workbook()
        ws = wb.active
        # 「学生氏名＿カナ」列が欠落したヘッダー
        ws.append(["学籍番号", "学生氏名", "備考"])
        ws.append(["K001", "山田 太郎", ""])
        wb.save(excel_path)

        with self.assertRaises(RosterConversionError):
            convert_excel_to_csv(excel_path, output_csv)

if __name__ == "__main__":
    unittest.main()
