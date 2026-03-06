import pandas as pd
import os

# サンプルファイルを読み込む
file_path = 'サンプル_支店リスト.xlsx'

if os.path.exists(file_path):
    print(f"ファイルパス: {file_path}")
    print("="*60)

    try:
        df = pd.read_excel(file_path)
        print(f"列名: {df.columns.tolist()}")
        print(f"\nデータ件数: {len(df)}行")
        print("\n最初の5行:")
        print(df.head())

        # CSV変換をテスト
        print("\n" + "="*60)
        print("CSV変換テスト")
        print("="*60)

        # 店舗名と支店名の列を探す
        store_col = None
        branch_col = None

        for col in df.columns:
            col_str = str(col).lower()
            if '店舗' in col_str or 'store' in col_str:
                store_col = col
                print(f"店舗名列を発見: {col}")
            if '支店' in col_str or 'branch' in col_str:
                branch_col = col
                print(f"支店名列を発見: {col}")

        # 抽出する列を決定
        columns_to_extract = []
        if store_col:
            columns_to_extract.append(store_col)
        if branch_col:
            columns_to_extract.append(branch_col)

        # 列が見つからない場合は最初の2列を使用
        if not columns_to_extract:
            print("店舗名・支店名の列が見つかりません。最初の2列を使用します")
            columns_to_extract = df.columns[:2].tolist()
            print(f"使用する列: {columns_to_extract}")

        # データを抽出
        extracted_df = df[columns_to_extract].copy()

        # 列名を標準化
        if len(extracted_df.columns) >= 2:
            extracted_df.columns = ['店舗名', '支店名']
        elif len(extracted_df.columns) == 1:
            extracted_df.columns = ['支店名']

        # 空行を削除
        extracted_df = extracted_df.dropna(how='all')

        print(f"\n抽出結果: {len(extracted_df)}件")
        print(extracted_df.head())

        # CSVに保存
        csv_path = 'outputs/test_店舗リスト.csv'
        extracted_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"\nCSV保存完了: {csv_path}")

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"ファイルが見つかりません: {file_path}")
