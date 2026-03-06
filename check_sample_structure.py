import pandas as pd
from openpyxl import load_workbook

# サンプルファイルを確認
file_path = 'サンプル_支店リスト.xlsx'

print("=" * 60)
print("サンプルファイルの構造確認")
print("=" * 60)

# pandasで読み込み
df = pd.read_excel(file_path, engine='openpyxl')
print(f"\n列名: {list(df.columns)}")
print(f"行数: {len(df)}")
print("\n最初の5行:")
print(df.head().to_string())

# 既存の住所、電話番号、営業時間の列があるか確認
print("\n" + "=" * 60)
print("既存の列の確認:")
print("=" * 60)
for col in df.columns:
    col_lower = str(col).lower()
    if '住所' in col_lower or 'address' in col_lower:
        print(f"✓ 住所列を発見: {col}")
    if '電話' in col_lower or 'phone' in col_lower or 'tel' in col_lower:
        print(f"✓ 電話番号列を発見: {col}")
    if '営業' in col_lower or 'hours' in col_lower or '時間' in col_lower:
        print(f"✓ 営業時間列を発見: {col}")
