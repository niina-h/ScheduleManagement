@echo off
chcp 65001 >nul
title 本番デプロイ（プログラムのみ）

echo ============================================
echo   予定管理システム - プログラム更新
echo   ※ DBは上書きしません
echo ============================================
echo.

REM --- 設定 ---
set SOURCE=%~dp0..
set DEST=\\192.168.70.141\C$\App\ScheduleManagement

echo 開発元 : %SOURCE%
echo 本番先 : %DEST%
echo.

REM --- 接続確認 ---
echo [確認] 本番サーバーへの接続を確認中...
ping -n 1 -w 2000 192.168.70.141 >nul 2>&1
if errorlevel 1 (
    echo.
    echo ★ エラー: 192.168.70.141 に接続できません。
    echo   ネットワーク接続を確認してください。
    echo.
    pause
    exit /b 1
)

if not exist "%DEST%\" (
    echo.
    echo ★ エラー: %DEST% にアクセスできません。
    echo   共有フォルダのアクセス権を確認してください。
    echo.
    pause
    exit /b 1
)

echo   → 接続OK
echo.
echo ※ サーバーの停止・再起動は手動で行ってください
echo.
pause

REM --- ソースコード同期 ---
echo.
echo [1/3] プログラムファイルを同期中...
robocopy "%SOURCE%\web_app" "%DEST%\web_app" /E /PURGE /XD __pycache__ /NFL /NDL /NJH /NJS
if errorlevel 8 (
    echo ★ エラー: web_app の同期に失敗しました
    pause
    exit /b 1
)

robocopy "%SOURCE%\data" "%DEST%\data" /E /NFL /NDL /NJH /NJS
robocopy "%SOURCE%\deploy" "%DEST%\deploy" /E /NFL /NDL /NJH /NJS
robocopy "%SOURCE%\bat" "%DEST%\bat" /E /NFL /NDL /NJH /NJS

copy /Y "%SOURCE%\run_web.py" "%DEST%\run_web.py" >nul
copy /Y "%SOURCE%\run_production.py" "%DEST%\run_production.py" >nul
copy /Y "%SOURCE%\requirements_web.txt" "%DEST%\requirements_web.txt" >nul
copy /Y "%SOURCE%\CLAUDE.md" "%DEST%\CLAUDE.md" >nul
copy /Y "%SOURCE%\README.md" "%DEST%\README.md" >nul

echo   → プログラム同期完了

REM --- 本番先にdbフォルダがなければ作成 ---
echo [2/3] フォルダ構成を確認中...
if not exist "%DEST%\db" mkdir "%DEST%\db"
if not exist "%DEST%\output" mkdir "%DEST%\output"
if not exist "%DEST%\reports" mkdir "%DEST%\reports"
echo   → OK

REM --- 完了 ---
echo [3/3] 完了確認...
echo.
echo ============================================
echo   プログラム更新完了
echo   ※ DBはそのまま（変更なし）
echo   ※ 本番サーバーの再起動を行ってください
echo ============================================
echo.
pause
