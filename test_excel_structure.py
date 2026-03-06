import pandas as pd
import os

# サンプルファイルを確認
excel_files = [
    'サンプル_支店リスト.xlsx',
    'outputs/店舗名一覧_202512.xlsx',
    'uploads/202512.xlsx'
]

for file_path in excel_files:
    if os.path.exists(file_path):
        print(f"\n{'='*60}")
        print(f"ファイル: {file_path}")
        print('='*60)
        try:
            df = pd.read_excel(file_path)
            print(f"列名: {df.columns.tolist()}")
            print(f"データ件数: {len(df)}行")
            print("\n最初の3行:")
            print(df.head(3))
        except Exception as e:
            print(f"エラー: {e}")
    else:
        print(f"\nファイルが見つかりません: {file_path}")
