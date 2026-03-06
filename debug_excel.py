import pandas as pd
import os

# サンプルファイルを読み込む
file_path = 'サンプル_支店リスト.xlsx'

if os.path.exists(file_path):
    print(f"ファイル: {file_path}")
    print("="*60)

    try:
        df = pd.read_excel(file_path)
        print(f"列名: {df.columns.tolist()}")
        print(f"\nデータ件数: {len(df)}行")
        print("\n最初の3行:")
        print(df.head(3))

        print("\n" + "="*60)
        print("列の検索テスト")
        print("="*60)

        for col in df.columns:
            col_str = str(col).lower()
            print(f"列名: '{col}' -> 小文字: '{col_str}'")

            if '業態' in col_str:
                print(f"  ✓ 業態列として検出")
            if '店舗' in col_str:
                print(f"  ✓ 店舗列として検出")
            if '店名' in col_str:
                print(f"  ✓ 店名列として検出")

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"ファイルが見つかりません: {file_path}")
