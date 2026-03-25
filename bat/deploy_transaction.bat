@echo off
chcp 65001 >nul
title 本番デプロイ（プログラム＋トランザクション移行）

echo ============================================
echo   予定管理システム - プログラム＋トランザクション移行
echo   ★ マスタデータは本番を維持 ★
echo ============================================
echo.

REM --- 設定 ---
set SOURCE=%~dp0..
set DEST=\\192.168.70.141\C$\App\ScheduleManagement
set DEV_DB=%SOURCE%\db\web_app.db
set PROD_DB=%DEST%\db\web_app.db

echo 開発元 : %SOURCE%
echo 本番先 : %DEST%
echo.

REM --- 接続確認 ---
echo [確認] 本番サーバーへの接続を確認中...
ping -n 1 -w 2000 192.168.70.141 >nul 2>&1
if errorlevel 1 (
    echo ★ エラー: 192.168.70.141 に接続できません。
    pause
    exit /b 1
)
if not exist "%DEST%\" (
    echo ★ エラー: %DEST% にアクセスできません。
    pause
    exit /b 1
)
echo   → 接続OK
echo.

echo ★★★ 移行内容 ★★★
echo   [維持] ユーザー・作業マスタ・区分・部署・メール設定（本番データ）
echo   [移行] 週間予定・日次実績・コメント・繰越・休暇・操作ログ（開発データ）
echo   [更新] プログラムファイル一式
echo.
echo   ※ 本番サーバーを先に停止してから実行してください
echo.
pause

REM --- 本番DBバックアップ ---
echo.
echo [1/3] 本番DBをバックアップ中...
set BACKUP_SUFFIX=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%
set BACKUP_SUFFIX=%BACKUP_SUFFIX: =0%
if not exist "%DEST%\db" mkdir "%DEST%\db"
if exist "%PROD_DB%" (
    copy /Y "%PROD_DB%" "%PROD_DB%.bak_%BACKUP_SUFFIX%" >nul
    echo   → web_app.db.bak_%BACKUP_SUFFIX% に保存
)

REM --- プログラム同期 ---
echo [2/3] プログラムファイルを同期中...
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

REM --- トランザクションデータ移行（Pythonスクリプト） ---
echo [3/3] トランザクションデータを移行中...
python "%SOURCE%\bat\_migrate_transaction.py" "%DEV_DB%" "%PROD_DB%"
if errorlevel 1 (
    echo ★ エラー: データ移行に失敗しました
    pause
    exit /b 1
)

echo.
echo ============================================
echo   デプロイ完了
echo   - プログラム更新済み
echo   - マスタデータ: 本番を維持
echo   - トランザクション: 開発から移行済み
echo   ※ 本番サーバーの再起動を行ってください
echo ============================================
echo.
pause
