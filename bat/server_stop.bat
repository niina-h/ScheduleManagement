@echo off
chcp 65001 >nul
title 本番サーバー停止

echo ============================================
echo   予定管理システム - 本番サーバー停止
echo ============================================
echo.

REM --- 設定 ---
set SERVER=192.168.70.141

REM --- 接続確認 ---
echo [確認] %SERVER% への接続を確認中...
ping -n 1 -w 2000 %SERVER% >nul 2>&1
if errorlevel 1 (
    echo ★ エラー: %SERVER% に接続できません。
    pause
    exit /b 1
)
echo   → 接続OK
echo.

REM --- 停止 ---
echo サーバーを停止中...
wmic /node:"%SERVER%" process where "name='pythonw.exe'" call terminate >nul 2>&1
if errorlevel 1 (
    echo.
    echo ★ リモート停止に失敗しました。
    echo   本番サーバーにリモートデスクトップで接続し、
    echo   以下を手動で実行してください：
    echo.
    echo   taskkill /F /IM pythonw.exe
    echo.
    pause
    exit /b 1
)

echo   → 停止コマンドを送信しました
timeout /t 2 /nobreak >nul

echo.
echo ============================================
echo   サーバー停止完了
echo ============================================
echo.
pause
