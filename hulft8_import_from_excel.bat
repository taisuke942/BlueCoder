@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM Place this BAT beside hulft_snd.txt and hulft_rcv.txt.
REM Register to a TEST/EMPTY HULFT8 environment first.
REM Without -r: duplicate IDs are NOT replaced.
set "BASE=%~dp0"
set "LOG=%BASE%hulft_import_result.log"

where utliupdt >nul 2>&1
if errorlevel 1 (
    echo ERROR: utliupdt was not found in PATH.
    echo Run this BAT on the HULFT8 Windows server or specify the full EXE path.
    pause
    exit /b 1
)

if not exist "%BASE%hulft_snd.txt" (
    echo ERROR: hulft_snd.txt not found.
    pause
    exit /b 1
)
if not exist "%BASE%hulft_rcv.txt" (
    echo ERROR: hulft_rcv.txt not found.
    pause
    exit /b 1
)

echo This will REGISTER definition cards in HULFT8.
echo Source folder: %BASE%
echo Review target host and dependent HOST/GROUP/JOB settings first.
choice /M "Continue with new registrations (without -r)"
if errorlevel 2 exit /b 0

echo ==== HULFT management import ==== > "%LOG%"
echo ==== SEND management ==== >> "%LOG%"
utliupdt -f "%BASE%hulft_snd.txt" >> "%LOG%" 2>&1
set "SND_RC=%ERRORLEVEL%"

echo ==== RECEIVE management ==== >> "%LOG%"
utliupdt -f "%BASE%hulft_rcv.txt" >> "%LOG%" 2>&1
set "RCV_RC=%ERRORLEVEL%"

echo Send command exit code: %SND_RC%
echo Receive command exit code: %RCV_RC%
echo Check the full log for per-ID errors: "%LOG%"
if not "%SND_RC%"=="0" echo WARNING: send import reported a command error.
if not "%RCV_RC%"=="0" echo WARNING: receive import reported a command error.
pause
if not "%SND_RC%"=="0" exit /b 1
if not "%RCV_RC%"=="0" exit /b 1
exit /b 0
