import pandas as pd
import os
import sys

# UTF-8で出力
sys.stdout.reconfigure(encoding='utf-8')

# サンプルファイルのパスを確認
sample_files = [
    'サンプル_支店リスト.xlsx',
    'uploads/サンプル_支店リスト.xlsx'
]

for file_path in sample_files:
    if os.path.exists(file_path):
        print(f"ファイル発見: {file_path}")
        print("=" * 80)

        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            print(f"列名: {df.columns.tolist()}")
            print(f"\nデータ件数: {len(df)}行")

            # 業態名の列を探す
            print("\n" + "=" * 80)
            print("業態名の列を探索:")
            print("=" * 80)

            business_type_col = None
            for col in df.columns:
                col_str = str(col).lower()
                print(f"列名: '{col}' -> 小文字: '{col_str}'")
                if '業態' in col_str:
                    business_type_col = col
                    print(f"  >>> 業態名列として検出: {col}")

            if business_type_col:
                print(f"\n業態名列: {business_type_col}")
                print("\n業態名のユニークな値（最初の20件）:")
                unique_values = df[business_type_col].dropna().unique()
                for idx, val in enumerate(unique_values[:20], 1):
                    print(f"  {idx}. {val}")

                print(f"\n総ユニーク数: {len(unique_values)}種類")

                # フィルタテスト
                print("\n" + "=" * 80)
                print("フィルタテスト:")
                print("=" * 80)

                # いくつかの業態名でフィルタをテスト
                test_filters = unique_values[:3] if len(unique_values) >= 3 else unique_values

                for test_filter in test_filters:
                    filter_str = str(test_filter).strip().lower()
                    matched = 0
                    for idx, row in df.iterrows():
                        business_type = row.get(business_type_col, '')
                        if business_type:
                            bt_str = str(business_type).strip().lower()
                            if bt_str == filter_str:
                                matched += 1

                    print(f"フィルタ「{test_filter}」: {matched}件マッチ")

            else:
                print("\n業態名の列が見つかりませんでした")

            break

        except Exception as e:
            print(f"エラー: {e}")
            import traceback
            traceback.print_exc()

else:
    print("サンプルファイルが見つかりませんでした")
    print("確認したパス:")
    for path in sample_files:
        print(f"  - {path}")
