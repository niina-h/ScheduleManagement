@echo off
chcp 65001 >nul
title 本番サーバー再起動

echo ============================================
echo   予定管理システム - 本番サーバー再起動
echo ============================================
echo.

REM --- 設定 ---
set SERVER=192.168.70.141
set APP_PATH=C:\App\ScheduleManagement

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
echo [1/2] サーバーを停止中...
wmic /node:"%SERVER%" process where "name='pythonw.exe'" call terminate >nul 2>&1
echo   → 停止コマンド送信
timeout /t 3 /nobreak >nul

REM --- 起動 ---
echo [2/2] サーバーを起動中...
wmic /node:"%SERVER%" process call create "cmd /c cd /d %APP_PATH% && start /b pythonw run_production.py" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ★ リモート起動に失敗しました。
    echo   本番サーバーにリモートデスクトップで接続し、
    echo   以下を手動で実行してください：
    echo.
    echo   taskkill /F /IM pythonw.exe
    echo   cd %APP_PATH%
    echo   start /b pythonw run_production.py
    echo.
    pause
    exit /b 1
)

echo   → 起動コマンド送信
echo.

REM --- 動作確認 ---
echo 起動を待機中（3秒）...
timeout /t 3 /nobreak >nul

echo 動作確認中...
powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://%SERVER%:5000/' -UseBasicParsing -TimeoutSec 5; Write-Host ('  → HTTP ' + $r.StatusCode + ' - OK') } catch { Write-Host '  → 応答なし（起動中の可能性があります）' }"

echo.
echo ============================================
echo   再起動完了
echo   URL: http://%SERVER%:5000/
echo ============================================
echo.
pause
