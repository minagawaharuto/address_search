"""
環境のセットアップを確認するスクリプト
使い方: python setup_check.py
"""
import sys
import os


def check_python_version():
    """Pythonバージョンの確認"""
    print("🐍 Python バージョン:")
    print(f"   {sys.version}")
    version_info = sys.version_info
    if version_info.major >= 3 and version_info.minor >= 8:
        print("   ✅ Python 3.8以上")
        return True
    else:
        print("   ❌ Python 3.8以上が必要です")
        return False


def check_packages():
    """必要なパッケージの確認"""
    print("\n📦 パッケージの確認:")

    required_packages = {
        'flask': '3.0.0',
        'selenium': '4.26.1',
        'beautifulsoup4': '4.12.2',
        'openpyxl': '3.1.2',
        'xlrd': '2.0.1',
        'pandas': '2.1.4',
        'webdriver_manager': '4.0.2',
        'lxml': '4.9.3',
        'requests': '2.31.0',
    }

    all_ok = True

    for package, expected_version in required_packages.items():
        try:
            if package == 'beautifulsoup4':
                import bs4
                version = bs4.__version__
                package_name = 'beautifulsoup4'
            elif package == 'webdriver_manager':
                import webdriver_manager
                version = webdriver_manager.__version__
                package_name = 'webdriver-manager'
            else:
                pkg = __import__(package)
                version = pkg.__version__
                package_name = package

            print(f"   ✅ {package_name}: {version}")
        except ImportError:
            print(f"   ❌ {package} がインストールされていません")
            all_ok = False
        except AttributeError:
            print(f"   ⚠ {package} はインストールされていますが、バージョンが確認できません")

    return all_ok


def check_chrome():
    """Google Chromeのインストール確認"""
    print("\n🌐 Google Chrome の確認:")

    import subprocess

    # Windowsの場合のChrome確認
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]

    chrome_found = False
    for path in chrome_paths:
        if os.path.exists(path):
            print(f"   ✅ Google Chrome が見つかりました")
            print(f"      パス: {path}")
            try:
                # バージョンを取得
                result = subprocess.run([path, '--version'], capture_output=True, text=True)
                if result.returncode == 0:
                    print(f"      {result.stdout.strip()}")
            except:
                pass
            chrome_found = True
            break

    if not chrome_found:
        print("   ❌ Google Chrome が見つかりません")
        print("      https://www.google.com/chrome/ からインストールしてください")
        return False

    return True


def check_chromedriver():
    """ChromeDriverの動作確認"""
    print("\n🚗 ChromeDriver の確認:")

    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from webdriver_manager.chrome import ChromeDriverManager

        print("   ChromeDriverをダウンロード中...")

        # ChromeDriverのダウンロード
        driver_path = ChromeDriverManager().install()
        print(f"   ✅ ChromeDriver ダウンロード成功")
        print(f"      パス: {driver_path}")

        # 実際に起動してみる
        print("   Chromeブラウザの起動テスト中...")
        chrome_options = Options()
        chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')

        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)

        # 簡単なテスト
        driver.get("https://www.google.com")
        print(f"   ✅ ブラウザ起動成功")
        print(f"      タイトル: {driver.title}")

        driver.quit()
        print("   ✅ すべてのテストが成功しました")
        return True

    except Exception as e:
        print(f"   ❌ エラーが発生しました: {e}")
        import traceback
        print("\n   詳細なエラー:")
        traceback.print_exc()
        return False


def main():
    print("=" * 60)
    print("住所検索システム - 環境チェック")
    print("=" * 60)
    print()

    results = []

    # Python バージョン
    results.append(("Python バージョン", check_python_version()))

    # パッケージ
    results.append(("必要なパッケージ", check_packages()))

    # Chrome
    results.append(("Google Chrome", check_chrome()))

    # ChromeDriver
    results.append(("ChromeDriver", check_chromedriver()))

    # 結果サマリー
    print("\n" + "=" * 60)
    print("📊 チェック結果サマリー")
    print("=" * 60)

    all_passed = True
    for name, passed in results:
        status = "✅ OK" if passed else "❌ NG"
        print(f"{status} - {name}")
        if not passed:
            all_passed = False

    print("=" * 60)

    if all_passed:
        print("\n✅ すべてのチェックが完了しました！")
        print("\n次のステップ:")
        print("1. アプリケーションを起動: python app.py")
        print("2. ブラウザで http://localhost:5000 を開く")
        print("3. Excelファイルをアップロードして住所検索を開始")
    else:
        print("\n❌ いくつかの問題が見つかりました")
        print("\n対処方法:")
        print("1. パッケージの再インストール:")
        print("   pip uninstall selenium webdriver-manager -y")
        print("   pip install -r requirements.txt --upgrade")
        print("\n2. Google Chromeのインストール:")
        print("   https://www.google.com/chrome/")
        print("\n3. システムの再起動")

    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nチェックをキャンセルしました")
    except Exception as e:
        print(f"\n❌ 予期しないエラー: {e}")
        import traceback
        traceback.print_exc()
