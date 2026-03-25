@echo off
chcp 65001 >nul
title 本番サーバー状態確認

echo ============================================
echo   予定管理システム - 本番サーバー状態確認
echo ============================================
echo.

REM --- 設定 ---
set SERVER=192.168.70.141

REM --- ネットワーク接続確認 ---
echo [1/3] ネットワーク接続...
ping -n 1 -w 2000 %SERVER% >nul 2>&1
if errorlevel 1 (
    echo   → NG（サーバーに接続できません）
    echo.
    pause
    exit /b 1
)
echo   → OK

REM --- プロセス確認 ---
echo [2/3] pythonw.exe プロセス...
wmic /node:"%SERVER%" process where "name='pythonw.exe'" get ProcessId 2>nul | findstr /r "[0-9]" >nul 2>&1
if errorlevel 1 (
    echo   → 停止中（pythonw.exe が見つかりません）
) else (
    echo   → 実行中
)

REM --- HTTP応答確認 ---
echo [3/3] HTTP応答（ポート5000）...
powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://%SERVER%:5000/' -UseBasicParsing -TimeoutSec 5; Write-Host ('  → HTTP ' + $r.StatusCode + ' - 正常稼働中') } catch { Write-Host '  → 応答なし' }"

echo.
echo ============================================
echo   確認完了
echo   URL: http://%SERVER%:5000/
echo ============================================
echo.
pause
