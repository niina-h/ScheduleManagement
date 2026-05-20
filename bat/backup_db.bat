@echo off
chcp 932 >nul
REM ============================================
REM Daily DB Backup (run on production server)
REM   Source : db\web_app.db
REM   Target : db\backups\web_app_YYYYMMDD_HHMMSS.db
REM   Keep   : latest 30 generations
REM   Log    : logs\backup.log
REM ============================================

set ROOT=C:\App\ScheduleManagement
set DB_SOURCE=%ROOT%\db\web_app.db
set BACKUP_DIR=%ROOT%\db\backups
set LOG_DIR=%ROOT%\logs
set LOG_FILE=%LOG_DIR%\backup.log

if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

for /f "tokens=2 delims==" %%a in ('wmic os get LocalDateTime /value 2^>nul') do set LDT=%%a
set TIMESTAMP=%LDT:~0,4%%LDT:~4,2%%LDT:~6,2%_%LDT:~8,2%%LDT:~10,2%%LDT:~12,2%
set BACKUP_FILE=%BACKUP_DIR%\web_app_%TIMESTAMP%.db

copy /Y "%DB_SOURCE%" "%BACKUP_FILE%" >nul
if errorlevel 1 (
    echo [%date% %time%] ERROR: backup copy failed >> "%LOG_FILE%"
    exit /b 1
)
echo [%date% %time%] OK: %BACKUP_FILE% >> "%LOG_FILE%"

REM Remove old backups (keep latest 30)
set /a KEEP=30
set /a COUNT=0
for /f "delims=" %%f in ('dir /b /o-d /a-d "%BACKUP_DIR%\web_app_*.db" 2^>nul') do (
    set /a COUNT+=1
    setlocal enabledelayedexpansion
    if !COUNT! GTR %KEEP% (
        del "%BACKUP_DIR%\%%f"
        echo [%date% %time%] DELETED: %%f >> "%LOG_FILE%"
    )
    endlocal
)
exit /b 0
