@echo off
chcp 65001 >nul
title 予定管理システム - 初回セットアップ

echo ============================================
echo   予定管理システム 初回セットアップ
echo ============================================
echo.

cd /d "%~dp0.."

REM --- Python 確認 ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Python が見つかりません。
    echo   https://www.python.org/downloads/ から Python 3.9 以上をインストールしてください。
    echo   インストール時に「Add Python to PATH」にチェックを入れてください。
    pause
    exit /b 1
)

echo [1/3] Python バージョン確認...
python --version

echo.
echo [2/3] 依存パッケージをインストール中...
pip install -r requirements_web.txt
if errorlevel 1 (
    echo [エラー] パッケージのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo [3/3] 動作確認中...
python -c "from web_app.app import create_app; create_app(); print('[OK] アプリケーション生成成功')"
if errorlevel 1 (
    echo [エラー] アプリケーションの生成に失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   セットアップ完了
echo ============================================
echo.
echo 起動方法:
echo   deploy\start_server.bat をダブルクリック
echo.
echo アクセス先:
echo   http://localhost:5000
echo   http://（このPCのIPアドレス）:5000
echo.

pause
