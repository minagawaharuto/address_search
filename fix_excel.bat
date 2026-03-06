@echo off
chcp 65001 >nul
echo ====================================
echo Excel ファイル修正ツール
echo ====================================
echo.

if "%~1"=="" (
    echo 使い方: このバッチファイルにExcelファイルをドラッグ^&ドロップしてください
    echo.
    echo または、コマンドプロンプトで:
    echo   fix_excel.bat "ファイルパス.xlsx"
    echo.
    pause
    exit /b
)

echo 対象ファイル: %~1
echo.

python fix_excel.py "%~1"

echo.
pause
