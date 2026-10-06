import os
import csv
import unicodedata
import openpyxl
from core.utils import atomic_write_text

REQUIRED_COLUMNS = ['学籍番号', '学生氏名', '学生氏名＿カナ']
OUTPUT_HEADER = ['学籍番号', '学生氏名', '学生氏名＿カナ', '投影実施可否']

# ヘッダーを探す範囲（各シートの先頭行数）と、エラーメッセージに含める列名の数
HEADER_SEARCH_ROWS = 30
MAX_REPORTED_COLUMNS = 8

class RosterConversionError(Exception):
    """名簿変換時のエラーを表す例外クラス"""
    pass

def _normalize_header(value) -> str:
    """全角・半角の違いや空白を吸収して列名を比較できるようにする（例: 学生氏名_カナ）"""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return "".join(text.split())

_NORMALIZED_REQUIRED = [_normalize_header(c) for c in REQUIRED_COLUMNS]

def _find_header(sheet):
    """
    シートの先頭から必須列を含むヘッダー行を探す。

    Returns:
        (列インデックスのリスト or None, 行イテレータ, 先頭付近で見つかった列名のリスト)
    """
    row_iter = sheet.iter_rows(values_only=True)
    seen_headers = []
    for _, row in zip(range(HEADER_SEARCH_ROWS), row_iter):
        if not row:
            continue
        normalized = [_normalize_header(cell) for cell in row]
        if all(col in normalized for col in _NORMALIZED_REQUIRED):
            return [normalized.index(col) for col in _NORMALIZED_REQUIRED], row_iter, seen_headers
        if not seen_headers:
            seen_headers = [str(cell).strip() for cell in row if cell is not None and str(cell).strip()]
    return None, None, seen_headers

def _cell(row, index):
    return row[index] if index < len(row) and row[index] is not None else ''

def convert_excel_to_csv(excel_file_path: str, output_csv_path: str) -> dict:
    """
    Excelファイルから「学籍番号」「学生氏名」「学生氏名＿カナ」を抽出し、
    「投影実施可否」列を追加したCSVファイルとして保存する。
    先頭シートで見つからない場合は、ほかのシートも順に探す。

    Args:
        excel_file_path: 変換元のExcelファイルパス
        output_csv_path: 出力先CSVファイルパス

    Returns:
        dict: {
            "total_extracted": 抽出された学生数,
            "duplicate_ids": 重複している学籍番号の件数,
            "sheet_name": 抽出元のシート名,
            "output_path": 出力されたCSVパス
        }

    Raises:
        FileNotFoundError: ファイルが存在しない場合
        RosterConversionError: 必須列が見つからない場合や変換エラー時
    """
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"指定されたExcelファイルが見つかりません: {excel_file_path}")

    if os.path.splitext(excel_file_path)[1].lower() == '.xls':
        raise RosterConversionError(
            "旧形式のExcelファイル（.xls）には対応していません。\n"
            "Excelで開き、「名前を付けて保存」で .xlsx 形式に保存し直してから選択してください。"
        )

    wb = None
    try:
        wb = openpyxl.load_workbook(excel_file_path, read_only=True, data_only=True)

        columns, row_iter, sheet_name = None, None, None
        report = []
        for sheet in wb.worksheets:
            columns, row_iter, seen = _find_header(sheet)
            if columns:
                sheet_name = sheet.title
                break
            shown = '、'.join(seen[:MAX_REPORTED_COLUMNS]) or '（列名が見つかりません）'
            report.append(f"・シート「{sheet.title}」: {shown}")

        if not columns:
            raise RosterConversionError(
                f"選択されたExcelファイルに必須列（{', '.join(REQUIRED_COLUMNS)}）が見つかりませんでした。\n\n"
                "各シートの先頭で見つかった列名:\n" + "\n".join(report)
            )

        col_id, col_name, col_kana = columns
        rows = []
        for row in row_iter:
            if not row:
                continue
            val_id = str(_cell(row, col_id)).strip()
            if val_id:
                rows.append([val_id, str(_cell(row, col_name)).strip(), str(_cell(row, col_kana)).strip(), ''])

        def _write(f):
            writer = csv.writer(f)
            writer.writerow(OUTPUT_HEADER)
            writer.writerows(rows)

        # 出力先ディレクトリの確保
        out_dir = os.path.dirname(os.path.abspath(output_csv_path))
        os.makedirs(out_dir, exist_ok=True)
        atomic_write_text(output_csv_path, _write)

        ids = [r[0] for r in rows]
        return {
            "total_extracted": len(rows),
            "duplicate_ids": len(ids) - len(set(ids)),
            "sheet_name": sheet_name,
            "output_path": output_csv_path
        }

    except RosterConversionError:
        raise
    except Exception as e:
        raise RosterConversionError(f"Excelファイルの読み込み/書き込み中にエラーが発生しました: {e}") from e
    finally:
        if wb is not None:
            wb.close()
