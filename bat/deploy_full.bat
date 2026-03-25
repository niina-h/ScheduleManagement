@echo off
chcp 65001 >nul
title 本番デプロイ（プログラム＋DB）

echo ============================================
echo   予定管理システム - フルデプロイ
echo   ★ プログラム＋DB すべて上書き ★
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
echo ★★★ 警告 ★★★
echo   本番DBを開発DBで上書きします。
echo   本番の既存データはバックアップ後に上書きされます。
echo   本番サーバーを先に停止してから実行してください。
echo.
pause

REM --- 本番DBバックアップ ---
echo.
echo [1/4] 本番DBをバックアップ中...
set BACKUP_SUFFIX=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%
set BACKUP_SUFFIX=%BACKUP_SUFFIX: =0%

if not exist "%DEST%\db" mkdir "%DEST%\db"

if exist "%DEST%\db\web_app.db" (
    copy /Y "%DEST%\db\web_app.db" "%DEST%\db\web_app.db.bak_%BACKUP_SUFFIX%" >nul
    echo   → web_app.db.bak_%BACKUP_SUFFIX% に保存
)
if exist "%DEST%\db\schedule.db" (
    copy /Y "%DEST%\db\schedule.db" "%DEST%\db\schedule.db.bak_%BACKUP_SUFFIX%" >nul
    echo   → schedule.db.bak_%BACKUP_SUFFIX% に保存
)

REM --- ソースコード同期 ---
echo [2/4] プログラムファイルを同期中...
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

if not exist "%DEST%\output" mkdir "%DEST%\output"
if not exist "%DEST%\reports" mkdir "%DEST%\reports"

echo   → プログラム同期完了

REM --- DB上書きコピー ---
echo [3/4] DBを上書きコピー中...
copy /Y "%SOURCE%\db\web_app.db" "%DEST%\db\web_app.db" >nul
echo   → web_app.db をコピーしました
if exist "%SOURCE%\db\schedule.db" (
    copy /Y "%SOURCE%\db\schedule.db" "%DEST%\db\schedule.db" >nul
    echo   → schedule.db をコピーしました
)

REM --- 完了 ---
echo [4/4] 完了確認...
echo.
echo ============================================
echo   フルデプロイ完了（プログラム＋DB上書き）
echo   ※ 本番サーバーの再起動を行ってください
echo ============================================
echo.
pause
