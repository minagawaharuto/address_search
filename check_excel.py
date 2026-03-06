import pandas as pd

# エクセルファイルを読み込む
df = pd.read_excel('サンプル_支店リスト.xlsx')

print('列名:', df.columns.tolist())
print('\nデータ件数:', len(df))
print('\n最初の5行:')
print(df.head())
print('\nデータ型:')
print(df.dtypes)
