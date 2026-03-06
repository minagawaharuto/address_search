"""
Places API (New)で取得できる情報をテストするスクリプト
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
print("Places API (New) - 詳細情報取得テスト")
print("=" * 70)

# テスト検索: 三菱UFJ銀行 新宿支店
url = "https://places.googleapis.com/v1/places:searchText"

# すべての有用な情報を取得するためのField Mask
headers = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": API_KEY,
    "X-Goog-FieldMask": (
        "places.displayName,"
        "places.formattedAddress,"
        "places.addressComponents,"
        "places.nationalPhoneNumber,"
        "places.internationalPhoneNumber,"
        "places.location,"
        "places.types,"
        "places.id"
    )
}

data = {
    "textQuery": "三菱UFJ銀行 新宿支店",
    "languageCode": "ja"
}

try:
    response = requests.post(url, json=data, headers=headers, timeout=10)
    print(f"ステータスコード: {response.status_code}\n")

    if response.status_code == 200:
        result = response.json()

        if 'places' in result and len(result['places']) > 0:
            place = result['places'][0]

            print("=" * 70)
            print("取得できる情報の例:")
            print("=" * 70)

            # 建物名/施設名
            if 'displayName' in place:
                print(f"\n[建物名/施設名]")
                print(f"  {place['displayName'].get('text', 'N/A')}")

            # 住所
            if 'formattedAddress' in place:
                print(f"\n[住所]")
                print(f"  {place['formattedAddress']}")

            # 郵便番号（addressComponentsから抽出）
            if 'addressComponents' in place:
                print(f"\n[住所の詳細情報]")
                postal_code = None
                for component in place['addressComponents']:
                    types = component.get('types', [])
                    if 'postal_code' in types:
                        postal_code = component.get('longText', '')
                        print(f"  郵便番号: {postal_code}")

                    # その他の住所構成要素も表示
                    print(f"  {', '.join(types)}: {component.get('longText', 'N/A')}")

            # 電話番号
            if 'nationalPhoneNumber' in place:
                print(f"\n[国内電話番号]")
                print(f"  {place['nationalPhoneNumber']}")

            if 'internationalPhoneNumber' in place:
                print(f"\n[国際電話番号]")
                print(f"  {place['internationalPhoneNumber']}")

            # 位置情報
            if 'location' in place:
                print(f"\n[位置情報（緯度経度）]")
                print(f"  緯度: {place['location'].get('latitude', 'N/A')}")
                print(f"  経度: {place['location'].get('longitude', 'N/A')}")

            # 場所のタイプ
            if 'types' in place:
                print(f"\n[場所のタイプ]")
                print(f"  {', '.join(place['types'])}")

            print("\n" + "=" * 70)
            print("完全なレスポンス:")
            print("=" * 70)
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            print("検索結果が見つかりませんでした")
    else:
        print(f"[ERROR] API呼び出しに失敗しました")
        print(f"レスポンス: {response.text}")

except Exception as e:
    print(f"[ERROR] エラー: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 70)
print("取得可能な情報まとめ")
print("=" * 70)
print("""
Excelに出力できる情報:
1. 建物名/施設名 (displayName)
2. 住所 (formattedAddress)
3. 郵便番号 (addressComponents内のpostal_code)
4. 電話番号 (nationalPhoneNumber)
5. 緯度 (location.latitude)
6. 経度 (location.longitude)
7. 場所のタイプ (types)

これらの情報を自動的に抽出してExcelに追加します。
""")
