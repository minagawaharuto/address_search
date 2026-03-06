# -*- coding: utf-8 -*-
import pandas as pd
from openpyxl import load_workbook
import os

# outputsフォルダ内のファイルを確認
output_folder = 'outputs'

print("=" * 60)
print("Outputsフォルダのファイル確認")
print("=" * 60)

files = [f for f in os.listdir(output_folder) if f.endswith('.xlsx')]
print(f"\nファイル数: {len(files)}")

if files:
    # 最初のファイルを詳しく確認
    file_path = os.path.join(output_folder, files[0])
    print(f"\n確認ファイル: {files[0]}")

    df = pd.read_excel(file_path, engine='openpyxl')
    print(f"\n列名: {list(df.columns)}")
    print(f"行数: {len(df)}")
    print("\n最初の3行:")
    print(df.head(3).to_string())
