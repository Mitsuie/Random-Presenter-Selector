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
        
        # デフォルトのファイル名を設定（元のファイル名から拡張子を.csvに変更）
        default_file_name = "演習投影名簿_｛科目名｝" + ".csv"
        
        # 保存先ファイルのダイアログを表示
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
            
            header_found = False
            col_id = -1
            col_name = -1
            col_kana = -1

            for row in sheet.iter_rows(values_only=True):
                row_list = list(row)
                
                if not header_found:
                    # 「学籍番号」が含まれる行をヘッダー行として探す
                    if '学籍番号' in row_list:
                        header_found = True
                        col_id = row_list.index('学籍番号')
                        col_name = row_list.index('学生氏名') if '学生氏名' in row_list else -1
                        col_kana = row_list.index('学生氏名＿カナ') if '学生氏名＿カナ' in row_list else -1
                        
                        # 新しいヘッダーを作成
                        new_header = ['学籍番号', '学生氏名', '学生氏名＿カナ', '投影実施可否']
                        writer.writerow(new_header)
                else:
                    # データ行の処理（学籍番号が空でない場合のみ出力）
                    if col_id != -1 and row_list[col_id] is not None and str(row_list[col_id]).strip() != '':
                        new_row = [
                            row_list[col_id] if col_id != -1 else '',
                            row_list[col_name] if col_name != -1 else '',
                            row_list[col_kana] if col_kana != -1 else '',
                            ''  # 投影実施可否の初期値は空欄
                        ]
                        writer.writerow(new_row)
        
        print(f"成功: {excel_file_path} に「投影実施可否」列を追加し、{csv_file_path} に変換・保存しました。")
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
        
    # 処理終了後にコンソールがすぐに閉じないようにキー入力を待つ
    input("\nEnterキーを押して終了してください...")
