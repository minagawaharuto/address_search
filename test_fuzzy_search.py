"""
曖昧検索機能のテストスクリプト
"""
import sys
import io

# Windowsコンソールのエンコーディング問題を回避
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# app_googlemapsから関数をインポート
from app_googlemaps import generate_search_patterns, load_api_key

# APIキーを読み込む
load_api_key()

print("=" * 70)
print("曖昧検索パターン生成テスト")
print("=" * 70)

# テストケース
test_cases = [
    ("三菱UFJ銀行", "新宿"),
    ("三菱UFJ銀行", "新宿支店"),
    ("みずほ銀行", "渋谷店"),
    ("セブンイレブン", "六本木ヒルズ"),
    ("スターバックス", "表参道"),
    ("ファミリーマート", "秋葉原駅前"),
]

for company, branch in test_cases:
    patterns = generate_search_patterns(company, branch)
    print(f"\n会社名: {company}")
    print(f"支店名: {branch}")
    print(f"生成されたパターン ({len(patterns)}個):")
    for i, pattern in enumerate(patterns, 1):
        print(f"  {i}. {pattern}")

print("\n" + "=" * 70)
print("実際の検索テスト（曖昧な支店名）")
print("=" * 70)

# 実際に検索をテスト
from app_googlemaps import search_address_with_googlemaps

test_searches = [
    ("三菱UFJ銀行", "新宿"),  # 「支店」なし
    ("スターバックス", "表参道"),  # 「店」なし
]

for company, branch in test_searches:
    print(f"\n検索: {company} {branch}")
    print("-" * 70)

    result = search_address_with_googlemaps(company, branch)

    if result['status'] == 'success':
        print(f"[成功] {result['building_name']}")
        print(f"  住所: {result['address']}")
        print(f"  郵便番号: {result['postal_code']}")
        print(f"  電話番号: {result['phone']}")
        print(f"  使用した検索クエリ: {result.get('search_query', 'N/A')}")
    else:
        print(f"[失敗] {result.get('message', 'エラー')}")

print("\n" + "=" * 70)
print("テスト完了")
print("=" * 70)
