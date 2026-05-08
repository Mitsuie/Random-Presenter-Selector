import openpyxl
import csv
import tkinter as tk
from tkinter import filedialog
import os

def excel_to_csv_with_column(excel_file_path):
    try:
        # Excelファイルを読み込む (数式ではなく計算結果を取得)
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        # 最初のシートを取得
        sheet = wb.worksheets[0]
        
        # イテレータを取得
        row_iter = sheet.iter_rows(values_only=True)
        
        # ヘッダー行を探す
        for row in row_iter:
            if '学籍番号' in row and '学生氏名' in row and '学生氏名＿カナ' in row:
                col_id = row.index('学籍番号')
                col_name = row.index('学生氏名')
                col_kana = row.index('学生氏名＿カナ')
                break
        else:
            # イテレータが最後まで到達した（breakされなかった）場合はヘッダーなし
            raise ValueError("選択されたExcelファイルに必須列（「学籍番号」「学生氏名」「学生氏名＿カナ」）が見つかりませんでした。")
        
        # 必須列が見つかった場合のみ、保存先ファイルのダイアログを表示
        default_file_name = "演習投影名簿_｛科目名｝" + ".csv"
        csv_file_path = filedialog.asksaveasfilename(
            title="変換したCSVファイルの保存先を選択",
            initialfile=default_file_name,
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        
        if not csv_file_path:
            print("ファイルの保存がキャンセルされました。")
            return
            
        # CSVファイルとして保存する
        # encoding='utf-8-sig' はWindows環境のExcelでCSVを開いた際の文字化けを防ぐための設定です
        with open(csv_file_path, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # 新しいヘッダーを作成
            new_header = ['学籍番号', '学生氏名', '学生氏名＿カナ', '投影実施可否']
            writer.writerow(new_header)
            
            # 残りのデータ行を処理（ヘッダーが見つかった次の行から再開される）
            for row in row_iter:
                # 学籍番号が空でない場合のみ出力
                if row[col_id] is not None and str(row[col_id]).strip() != '':
                    new_row = [
                        row[col_id],
                        row[col_name],
                        row[col_kana],
                        ''  # 投影実施可否の初期値は空欄
                    ]
                    writer.writerow(new_row)
            
        print(f"成功: {excel_file_path} から必須3列（「学籍番号」「学生氏名」「学生氏名＿カナ」）を抽出し、「投影実施可否」列を追加したCSVとして {csv_file_path} に変換・保存しました。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 使用例
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    # 変換元のExcelファイルのパスをダイアログで選択
    input_excel = filedialog.askopenfilename(
        title="変換元のExcelファイルを選択",
        filetypes=[("Excelファイル", "*.xlsx *.xlsm *.xltx *.xltm"), ("すべてのファイル", "*.*")]
    )
    
    if input_excel:
        excel_to_csv_with_column(input_excel)
    else:
        print("ファイルの選択がキャンセルされました。")

