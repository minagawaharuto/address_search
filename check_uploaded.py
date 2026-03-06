import pandas as pd
import sys

# UTF-8で出力
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'uploads/202512.xlsx'

try:
    df = pd.read_excel(file_path, engine='openpyxl')
    print(f"ファイル: {file_path}")
    print("=" * 80)
    print(f"列名: {df.columns.tolist()}")
    print(f"\nデータ件数: {len(df)}行")

    print("\n" + "=" * 80)
    print("列の検索:")
    print("=" * 80)

    business_type_col = None
    store_name_col = None

    for col in df.columns:
        col_str = str(col).lower()
        print(f"列名: '{col}' -> 小文字: '{col_str}'")
        if '業態' in col_str:
            business_type_col = col
            print(f"  >>> 業態名列として検出")
        if '店舗' in col_str or '店名' in col_str:
            store_name_col = col
            print(f"  >>> 店舗名列として検出")

    print("\n" + "=" * 80)
    print("検出結果:")
    print("=" * 80)
    print(f"業態名列: {business_type_col if business_type_col else '見つかりませんでした'}")
    print(f"店舗名列: {store_name_col if store_name_col else '見つかりませんでした'}")

    if business_type_col:
        print(f"\n業態名のユニークな値（最初の10件）:")
        unique_values = df[business_type_col].dropna().unique()
        for idx, val in enumerate(unique_values[:10], 1):
            print(f"  {idx}. {val}")
        print(f"\n総ユニーク数: {len(unique_values)}種類")

    print("\n最初の5行のデータ:")
    print(df.head(5))

except Exception as e:
    print(f"エラー: {e}")
    import traceback
    traceback.print_exc()
