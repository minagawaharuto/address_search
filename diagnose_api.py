"""
Google Maps API診断スクリプト
APIキーの状態と利用可能なAPIを確認します
"""
import requests
import json
import sys
import io

# Windowsコンソールのエンコーディング問題を回避
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# config.txtからAPIキーを読み込む
def load_api_key():
    try:
        with open('config.txt', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('GOOGLE_MAPS_API_KEY='):
                    return line.split('=', 1)[1].strip()
    except Exception as e:
        print(f"[ERROR] config.txtの読み込みに失敗: {e}")
        return None

API_KEY = load_api_key()

if not API_KEY:
    print("[ERROR] APIキーが見つかりません")
    exit(1)

print("=" * 70)
print("Google Maps API 診断ツール")
print("=" * 70)
print(f"\n[OK] APIキー: {API_KEY[:20]}...{API_KEY[-4:]}")
print()

# テスト1: Places API (New) - Text Search
print("\n" + "=" * 70)
print("テスト1: Places API (New) - Text Search")
print("=" * 70)

url = "https://places.googleapis.com/v1/places:searchText"
headers = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": API_KEY,
    "X-Goog-FieldMask": "places.displayName,places.formattedAddress"
}
data = {
    "textQuery": "東京駅",
    "languageCode": "ja"
}

try:
    response = requests.post(url, json=data, headers=headers, timeout=10)
    print(f"ステータスコード: {response.status_code}")
    print(f"レスポンス:\n{json.dumps(response.json(), indent=2, ensure_ascii=False)}")

    if response.status_code == 200:
        print("[OK] Places API (New) は正常に動作しています")
    else:
        print("[ERROR] Places API (New) でエラーが発生しました")
except Exception as e:
    print(f"[ERROR] エラー: {e}")

# テスト2: Places API (旧版) - Text Search
print("\n" + "=" * 70)
print("テスト2: Places API (旧版) - Text Search")
print("=" * 70)

url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
params = {
    "query": "東京駅",
    "language": "ja",
    "key": API_KEY
}

try:
    response = requests.get(url, params=params, timeout=10)
    print(f"ステータスコード: {response.status_code}")
    result = response.json()
    print(f"ステータス: {result.get('status')}")

    if 'error_message' in result:
        print(f"エラーメッセージ: {result['error_message']}")

    print(f"レスポンス:\n{json.dumps(result, indent=2, ensure_ascii=False)}")

    if result.get('status') == 'OK':
        print("[OK] Places API (旧版) は正常に動作しています")
    else:
        print("[ERROR] Places API (旧版) でエラーが発生しました")
except Exception as e:
    print(f"[ERROR] エラー: {e}")

# テスト3: Geocoding API
print("\n" + "=" * 70)
print("テスト3: Geocoding API")
print("=" * 70)

url = "https://maps.googleapis.com/maps/api/geocode/json"
params = {
    "address": "東京駅",
    "language": "ja",
    "key": API_KEY
}

try:
    response = requests.get(url, params=params, timeout=10)
    print(f"ステータスコード: {response.status_code}")
    result = response.json()
    print(f"ステータス: {result.get('status')}")

    if 'error_message' in result:
        print(f"エラーメッセージ: {result['error_message']}")

    if result.get('status') == 'OK':
        print("[OK] Geocoding API は正常に動作しています")
        if result.get('results'):
            print(f"住所例: {result['results'][0]['formatted_address']}")
    else:
        print("[ERROR] Geocoding API でエラーが発生しました")
        print(f"レスポンス:\n{json.dumps(result, indent=2, ensure_ascii=False)}")
except Exception as e:
    print(f"[ERROR] エラー: {e}")

# 診断結果のまとめ
print("\n" + "=" * 70)
print("診断結果とアドバイス")
print("=" * 70)
print("""
[REQUEST_DENIED エラーの主な原因]

1. 課金が有効になっていない
   -> Google Cloud Console で請求先アカウントを設定
   -> https://console.cloud.google.com/billing

2. 必要なAPIが有効化されていない
   -> 「APIとサービス」->「ライブラリ」で以下を有効化：
      - Places API
      - Places API (New)
      - Geocoding API

3. APIキーの制限設定が厳しすぎる
   -> 「APIとサービス」->「認証情報」-> APIキーをクリック
   -> 「アプリケーションの制限」を「なし」に設定（テスト時）
   -> 「API の制限」を確認

4. APIキーが無効または期限切れ
   -> 新しいAPIキーを作成

5. プロジェクトの設定が不完全
   -> プロジェクトの請求が有効か確認
   -> プロジェクトが正しく選択されているか確認

[推奨される対応]

ステップ1: Google Cloud Console で課金を有効化
  https://console.cloud.google.com/billing

ステップ2: 以下の3つのAPIを有効化
  - Places API
  - Places API (New)
  - Geocoding API

ステップ3: APIキーの制限を一時的に「なし」に設定

ステップ4: 数分待ってから再度このスクリプトを実行

それでも解決しない場合は、新しいAPIキーを作成してください。
""")
