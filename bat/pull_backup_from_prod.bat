@echo off
chcp 932 >nul
REM ============================================
REM Off-site Backup Pull (run on developer PC)
REM   Source : \\192.168.70.141\C$\App\ScheduleManagement\db\web_app.db
REM   Target : c:\DEV(ClaudCode)\Backup\ScheduleManagement\web_app_YYYYMMDD_HHMMSS.db
REM   Keep   : latest 30 generations
REM ============================================

set PROD_DB=\\192.168.70.141\C$\App\ScheduleManagement\db\web_app.db
set BACKUP_DIR=c:\DEV(ClaudCode)\Backup\ScheduleManagement
set LOG_FILE=%BACKUP_DIR%\backup.log

if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"

ping -n 1 -w 2000 192.168.70.141 >nul 2>&1
if errorlevel 1 (
    echo [%date% %time%] ERROR: cannot reach 192.168.70.141 >> "%LOG_FILE%"
    exit /b 1
)

if not exist "%PROD_DB%" (
    echo [%date% %time%] ERROR: production DB not found at %PROD_DB% >> "%LOG_FILE%"
    exit /b 1
)

for /f "tokens=2 delims==" %%a in ('wmic os get LocalDateTime /value 2^>nul') do set LDT=%%a
set TIMESTAMP=%LDT:~0,4%%LDT:~4,2%%LDT:~6,2%_%LDT:~8,2%%LDT:~10,2%%LDT:~12,2%
set BACKUP_FILE=%BACKUP_DIR%\web_app_%TIMESTAMP%.db

copy /Y "%PROD_DB%" "%BACKUP_FILE%" >nul
if errorlevel 1 (
    echo [%date% %time%] ERROR: copy failed >> "%LOG_FILE%"
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
