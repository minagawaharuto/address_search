import pandas as pd
import os
import sys

# UTF-8で出力
sys.stdout.reconfigure(encoding='utf-8')

# 出力されたExcelファイルを確認
file_path = 'outputs/住所追加_202512.xlsx'

if os.path.exists(file_path):
    print(f"ファイル: {file_path}")
    print("=" * 80)

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"列名: {df.columns.tolist()}")
        print(f"\nデータ件数: {len(df)}行")

        # 検索結果の列があるか確認
        result_columns = ['正式名称', '住所', '電話番号', '営業時間']
        print("\n" + "=" * 80)
        print("検索結果の列の確認:")
        print("=" * 80)
        for col in result_columns:
            if col in df.columns:
                print(f"[OK] {col}列が存在します")
                # 空でないデータの件数をカウント
                non_empty = df[col].notna().sum()
                print(f"     データが入っている行数: {non_empty}件")

                # データが入っている場合は最初の3件を表示
                if non_empty > 0:
                    print(f"     最初の3件のデータ:")
                    sample_data = df[df[col].notna()][col].head(3)
                    for idx, val in enumerate(sample_data, 1):
                        val_str = str(val)[:50]  # 最初の50文字まで表示
                        print(f"       {idx}. {val_str}")
            else:
                print(f"[NG] {col}列が見つかりません")

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"ファイルが見つかりません: {file_path}")
