"""
営業時間などの詳細情報が取得できるかテストするスクリプト
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
print("Places API (New) - 営業時間など詳細情報取得テスト")
print("=" * 70)

# テスト検索: スターバックス 表参道
url = "https://places.googleapis.com/v1/places:searchText"

# 営業時間やその他の情報を取得するためのField Mask
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
        "places.regularOpeningHours,"
        "places.currentOpeningHours,"
        "places.businessStatus,"
        "places.websiteUri,"
        "places.rating,"
        "places.userRatingCount,"
        "places.priceLevel,"
        "places.id"
    )
}

data = {
    "textQuery": "スターバックス コーヒー アクセス表参道店",
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
            print("取得できる情報:")
            print("=" * 70)

            # 基本情報
            if 'displayName' in place:
                print(f"\n[施設名]")
                print(f"  {place['displayName'].get('text', 'N/A')}")

            if 'formattedAddress' in place:
                print(f"\n[住所]")
                print(f"  {place['formattedAddress']}")

            # 電話番号
            if 'nationalPhoneNumber' in place:
                print(f"\n[電話番号]")
                print(f"  {place['nationalPhoneNumber']}")

            # 営業時間
            if 'regularOpeningHours' in place:
                print(f"\n[営業時間（通常）]")
                opening_hours = place['regularOpeningHours']

                if 'weekdayDescriptions' in opening_hours:
                    print("  営業時間:")
                    for desc in opening_hours['weekdayDescriptions']:
                        print(f"    {desc}")

                if 'openNow' in opening_hours:
                    status = "営業中" if opening_hours['openNow'] else "営業時間外"
                    print(f"  現在の状態: {status}")

            if 'currentOpeningHours' in place:
                print(f"\n[営業時間（現在）]")
                current_hours = place['currentOpeningHours']

                if 'weekdayDescriptions' in current_hours:
                    print("  営業時間:")
                    for desc in current_hours['weekdayDescriptions']:
                        print(f"    {desc}")

                if 'openNow' in current_hours:
                    status = "営業中" if current_hours['openNow'] else "営業時間外"
                    print(f"  現在の状態: {status}")

            # ビジネスステータス
            if 'businessStatus' in place:
                print(f"\n[営業状況]")
                print(f"  {place['businessStatus']}")

            # ウェブサイト
            if 'websiteUri' in place:
                print(f"\n[ウェブサイト]")
                print(f"  {place['websiteUri']}")

            # 評価
            if 'rating' in place:
                print(f"\n[評価]")
                print(f"  評価: {place['rating']}")
                if 'userRatingCount' in place:
                    print(f"  レビュー数: {place['userRatingCount']}")

            # 価格帯
            if 'priceLevel' in place:
                print(f"\n[価格帯]")
                print(f"  {place['priceLevel']}")

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
print("取得可能な追加情報まとめ")
print("=" * 70)
print("""
Excelに追加できる情報:
1. 営業時間 (regularOpeningHours/currentOpeningHours)
   - 曜日別の営業時間
   - 現在営業中かどうか
2. ウェブサイト (websiteUri)
3. 評価/レビュー (rating, userRatingCount)
4. 価格帯 (priceLevel)
5. 営業状況 (businessStatus) - OPERATIONAL, CLOSED_TEMPORARILY等

これらの情報をExcelに追加することができます。
""")
