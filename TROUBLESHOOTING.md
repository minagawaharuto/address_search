# トラブルシューティングガイド

## よくあるエラーと解決方法

### 1. Excelファイル形式エラー

#### エラーメッセージ
```
openpyxl does not support file format, please check you can open it with Excel first.
Supported formats are: .xlsx,.xlsm,.xltx,.xltm
```

#### 原因
- 古いExcel形式（.xls）を使用している
- ファイルが破損している
- Excelファイルではないファイルを選択している

#### 解決方法

**方法1: Excelで形式を変換**
1. Excelでファイルを開く
2. 「ファイル」→「名前を付けて保存」
3. ファイル形式を「Excel ブック (.xlsx)」に変更して保存
4. 新しいファイルをアップロード

**方法2: 自動変換機能を使用**
- 更新版のアプリケーションでは、.xls形式も自動的に.xlsxに変換されます
- 依存パッケージを再インストール：
```bash
pip install -r requirements.txt
```

**方法3: pandasを使って変換**
```python
import pandas as pd

# .xlsファイルを読み込み
df = pd.read_excel('古いファイル.xls', engine='xlrd')

# .xlsx形式で保存
df.to_excel('新しいファイル.xlsx', index=False, engine='openpyxl')
```

### 2. xlrd/openpyxlインストールエラー

#### エラーメッセージ
```
ModuleNotFoundError: No module named 'xlrd'
```

#### 解決方法
```bash
pip install xlrd==2.0.1
pip install openpyxl==3.1.2
pip install pandas==2.1.4
```

または、全ての依存パッケージを再インストール：
```bash
pip install -r requirements.txt --upgrade
```

### 3. ファイルが空・データが読み込めない

#### 症状
- アップロード後、何も処理されない
- エラーメッセージ「Excelファイルが空です」

#### 原因
- Excelファイルにデータが入っていない
- シートが複数あり、アクティブシートにデータがない
- ヘッダー行のみでデータ行がない

#### 解決方法
1. Excelファイルを開いて、データが入っているか確認
2. 最初のシート（一番左のタブ）にデータがあるか確認
3. 最低限以下の形式でデータを入力：

```
| 支店名 |
|--------|
| 東京支店 |
| 大阪支店 |
```

### 4. ChromeDriver/Seleniumエラー

#### エラーメッセージ
```
WebDriverException: Message: 'chromedriver' executable needs to be in PATH
```

#### 解決方法

**自動インストールを使用（推奨）**
```bash
pip install webdriver-manager
```

アプリケーションが自動的にChromeDriverをダウンロードします。

**Google Chromeがインストールされていない場合**
1. Google Chromeをインストール: https://www.google.com/chrome/
2. アプリケーションを再起動

### 5. 住所が見つからない

#### 症状
- 結果に「住所が見つかりませんでした」と表示される

#### 原因
- Google検索結果に住所情報が含まれていない
- 会社名や支店名のスペルが間違っている
- 検索結果の構造が想定と異なる

#### 解決方法

1. **会社名と支店名を確認**
   - 正式名称を使用（例: 株式会社○○、○○株式会社）
   - 支店名の表記を確認（支店、営業所、事業所など）

2. **手動で検索して確認**
   - ブラウザで「会社名 支店名 住所」で検索
   - 結果に住所が表示されるか確認
   - 表示されない場合は、その支店の住所はWeb上に公開されていない可能性

3. **検索待機時間を増やす**
   - `app.py`の`time.sleep(2)`を`time.sleep(5)`に変更

### 6. メモリ不足エラー

#### エラーメッセージ
```
MemoryError
```

#### 原因
- 大量の支店データを一度に処理している
- Seleniumのメモリリークが発生

#### 解決方法

1. **データを分割して処理**
   - 100件ずつなど、小さなバッチで処理

2. **headlessモードを有効化**
   - `app.py`で既に有効化されていますが、念のため確認：
   ```python
   chrome_options.add_argument('--headless')
   ```

### 7. アップロードサイズエラー

#### エラーメッセージ
```
413 Request Entity Too Large
```

#### 原因
- ファイルサイズが16MBを超えている

#### 解決方法

**app.pyの設定を変更**
```python
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MBに変更
```

### 8. 日本語文字化け

#### 症状
- 住所に文字化けが発生

#### 解決方法

**Excelファイルを保存し直す**
1. Excelで「名前を付けて保存」
2. 「ツール」→「Web オプション」
3. エンコーディングを「UTF-8」に設定

### 9. ポート使用中エラー

#### エラーメッセージ
```
OSError: [Errno 48] Address already in use
```

#### 解決方法

**ポートを変更**
```python
# app.pyの最終行を変更
app.run(debug=True, host='0.0.0.0', port=5001)  # 5000→5001
```

**既存プロセスを終了**
```bash
# Windowsの場合
netstat -ano | findstr :5000
taskkill /PID プロセスID /F

# Mac/Linuxの場合
lsof -i :5000
kill -9 プロセスID
```

## デバッグモード

より詳細なログを確認するには、app.pyで以下を変更：

```python
logging.basicConfig(level=logging.DEBUG)  # INFOからDEBUGに変更
```

## サポート情報の収集

問題が解決しない場合、以下の情報を収集してください：

1. **環境情報**
```bash
python --version
pip list
```

2. **エラーログ**
   - コンソールに表示される完全なエラーメッセージ

3. **テストファイル**
   - 問題が発生するExcelファイルのサンプル（個人情報削除後）

4. **ブラウザのコンソール**
   - F12キーを押して、Consoleタブのエラーメッセージ

## よくある質問（FAQ）

### Q: 何件まで処理できますか？
A: 理論上は制限ありませんが、各検索に約5秒かかるため、100件で約8分かかります。大量の場合は分割処理を推奨します。

### Q: Google Maps APIを使えますか？
A: はい。より高精度な検索を希望する場合は、コードをGoogle Maps APIに対応させることができます（別途APIキーが必要）。

### Q: オフラインで使えますか？
A: いいえ。Google検索を使用するため、インターネット接続が必要です。

### Q: 住所の精度はどのくらいですか？
A: Google検索結果に依存します。公式サイトに住所が記載されている場合は高精度ですが、保証はできません。
