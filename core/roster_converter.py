import os
import csv
import openpyxl

REQUIRED_COLUMNS = ['学籍番号', '学生氏名', '学生氏名＿カナ']
OUTPUT_HEADER = ['学籍番号', '学生氏名', '学生氏名＿カナ', '投影実施可否']

class RosterConversionError(Exception):
    """名簿変換時のエラーを表す例外クラス"""
    pass

def convert_excel_to_csv(excel_file_path: str, output_csv_path: str) -> dict:
    """
    Excelファイルから「学籍番号」「学生氏名」「学生氏名＿カナ」を抽出し、
    「投影実施可否」列を追加したCSVファイルとして保存する。

    Args:
        excel_file_path: 変換元のExcelファイルパス
        output_csv_path: 出力先CSVファイルパス

    Returns:
        dict: {
            "total_extracted": 抽出された学生数,
            "output_path": 出力されたCSVパス
        }

    Raises:
        FileNotFoundError: ファイルが存在しない場合
        RosterConversionError: 必須列が見つからない場合や変換エラー時
    """
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"指定されたExcelファイルが見つかりません: {excel_file_path}")

    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        sheet = wb.worksheets[0]
        row_iter = sheet.iter_rows(values_only=True)

        # ヘッダー行を探す
        col_id = None
        col_name = None
        col_kana = None

        for row in row_iter:
            if not row:
                continue
            # Noneを除去した文字列リスト化
            row_str = [str(cell).strip() if cell is not None else "" for cell in row]
            if all(col in row_str for col in REQUIRED_COLUMNS):
                col_id = row_str.index('学籍番号')
                col_name = row_str.index('学生氏名')
                col_kana = row_str.index('学生氏名＿カナ')
                break

        if col_id is None:
            raise RosterConversionError(
                f"選択されたExcelファイルに必須列（{', '.join(REQUIRED_COLUMNS)}）が見つかりませんでした。"
            )

        # 出力先ディレクトリの確保
        out_dir = os.path.dirname(os.path.abspath(output_csv_path))
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        extracted_count = 0
        with open(output_csv_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(OUTPUT_HEADER)

            for row in row_iter:
                if not row or len(row) <= max(col_id, col_name, col_kana):
                    continue
                val_id = row[col_id]
                if val_id is not None and str(val_id).strip() != '':
                    val_name = row[col_name] if row[col_name] is not None else ''
                    val_kana = row[col_kana] if row[col_kana] is not None else ''
                    writer.writerow([str(val_id).strip(), str(val_name).strip(), str(val_kana).strip(), ''])
                    extracted_count += 1

        return {
            "total_extracted": extracted_count,
            "output_path": output_csv_path
        }

    except RosterConversionError:
        raise
    except Exception as e:
        raise RosterConversionError(f"Excelファイルの読み込み/書き込み中にエラーが発生しました: {e}") from e
