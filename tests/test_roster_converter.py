import os
import csv
import tempfile
import unittest
from unittest.mock import patch
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

    def test_write_failure_keeps_existing_csv(self):
        """書き込み途中で失敗しても、既存の出力先CSVが壊れないこと"""
        excel_path = os.path.join(self.temp_dir.name, "roster.xlsx")
        csv_path = os.path.join(self.temp_dir.name, "output.csv")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["学籍番号", "学生氏名", "学生氏名＿カナ"])
        ws.append(["K001", "山田 太郎", "ヤマダ タロウ"])
        wb.save(excel_path)
        with open(csv_path, mode="w", encoding="utf-8-sig", newline="") as f:
            f.write("既存の記録\n")

        with patch("csv.writer", side_effect=OSError("disk full")):
            with self.assertRaises(RosterConversionError):
                convert_excel_to_csv(excel_path, csv_path)

        with open(csv_path, mode="r", encoding="utf-8-sig") as f:
            self.assertEqual(f.read(), "既存の記録\n")

    def _save_workbook(self, name, sheets):
        """sheets: [(シート名, 行のリスト), ...] からExcelファイルを作る"""
        path = os.path.join(self.temp_dir.name, name)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows in sheets:
            ws = wb.create_sheet(title)
            for row in rows:
                ws.append(row)
        wb.save(path)
        return path

    def test_header_variants_and_second_sheet(self):
        """列名の表記ゆれ（半角アンダースコア・空白）を吸収し、2枚目以降のシートも探すこと"""
        excel_path = self._save_workbook("variants.xlsx", [
            ("表紙", [["演習クラス名簿"], ["担当", "電大"]]),
            ("名簿", [
                ["学籍番号 ", "学生氏名", "学生氏名_カナ"],
                ["K001", "山田 太郎", "ヤマダ タロウ"],
                ["K002", "佐藤 花子"],  # カナが空欄でも学生は抽出する
            ]),
        ])
        csv_path = os.path.join(self.temp_dir.name, "out.csv")
        result = convert_excel_to_csv(excel_path, csv_path)

        self.assertEqual(result["sheet_name"], "名簿")
        self.assertEqual(result["total_extracted"], 2)
        with open(csv_path, mode="r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        self.assertEqual(rows[2], ["K002", "佐藤 花子", "", ""])

    def test_duplicate_ids_are_counted(self):
        excel_path = self._save_workbook("dup.xlsx", [("S", [
            ["学籍番号", "学生氏名", "学生氏名＿カナ"],
            ["K001", "山田 太郎", "ヤマダ タロウ"],
            ["K001", "山田 太郎", "ヤマダ タロウ"],
        ])])
        result = convert_excel_to_csv(excel_path, os.path.join(self.temp_dir.name, "out.csv"))
        self.assertEqual(result["total_extracted"], 2)
        self.assertEqual(result["duplicate_ids"], 1)

    def test_error_message_lists_found_columns(self):
        """必須列が見つからない場合、シート名と見つかった列名をメッセージに含めること"""
        excel_path = self._save_workbook("bad.xlsx", [("一覧", [["番号", "氏名"], ["1", "山田"]])])
        with self.assertRaises(RosterConversionError) as cm:
            convert_excel_to_csv(excel_path, os.path.join(self.temp_dir.name, "out.csv"))
        self.assertIn("一覧", str(cm.exception))
        self.assertIn("氏名", str(cm.exception))

    def test_xls_is_rejected_with_guidance(self):
        xls_path = os.path.join(self.temp_dir.name, "old.xls")
        with open(xls_path, "wb") as f:
            f.write(b"dummy")
        with self.assertRaises(RosterConversionError) as cm:
            convert_excel_to_csv(xls_path, os.path.join(self.temp_dir.name, "out.csv"))
        self.assertIn(".xlsx", str(cm.exception))

if __name__ == "__main__":
    unittest.main()
