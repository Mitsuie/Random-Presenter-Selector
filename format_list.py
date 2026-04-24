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
            
            for i, row in enumerate(sheet.iter_rows(values_only=True)):
                row_list = list(row)
                if i == 0:
                    # ヘッダー行に「投影実施可否」列を追加
                    row_list.append('投影実施可否')
                else:
                    # データ行には空欄を追加
                    row_list.append('')
                
                writer.writerow(row_list)
        
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
