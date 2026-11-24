@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM HULFT10 for Windows. Place this BAT in the SAME folder as:
REM   hulft10_snd.txt   (from Excel -> TXT VBA)
REM   hulft10_rcv.txt   (from Excel -> TXT VBA)
REM Safest first run: TEST / EMPTY HULFT10, NOT production.
REM No -r: existing management IDs are NOT overwritten.
REM If V8 and V10 coexist, specify full HULFT10 path below.
set "UTLIUPDT=utliupdt"
set "ROOT=%~dp0"
set "LOG=%ROOT%hulft10_import_result.log"

if not exist "%UTLIUPDT%" (
    where "%UTLIUPDT%" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: HULFT10 utliupdt command not found.
        echo Set UTLIUPDT to the full HULFT10 exe path.
        pause
        exit /b 3
    )
)
if not exist "%ROOT%hulft10_snd.txt" (
    echo ERROR: hulft10_snd.txt not found beside this BAT.
    pause
    exit /b 2
)
if not exist "%ROOT%hulft10_rcv.txt" (
    echo ERROR: hulft10_rcv.txt not found beside this BAT.
    pause
    exit /b 2
)

echo IMPORTANT: This registers settings in the CURRENT HULFT instance.
echo Verify you are running against the TEST HULFT10 environment.
echo Host/group/job/format information must already exist if referenced.
choice /M "Run HULFT10 management registration now"
if errorlevel 2 exit /b 0

echo ==== HULFT10 management registration ==== > "%LOG%"
echo ==== SEND ==== >> "%LOG%"
"%UTLIUPDT%" -f "%ROOT%hulft10_snd.txt" >> "%LOG%" 2>&1
set "RC_SND=%ERRORLEVEL%"

echo ==== RECEIVE ==== >> "%LOG%"
"%UTLIUPDT%" -f "%ROOT%hulft10_rcv.txt" >> "%LOG%" 2>&1
set "RC_RCV=%ERRORLEVEL%"

echo SEND exit code: %RC_SND%
echo RECEIVE exit code: %RC_RCV%
echo Read full log: "%LOG%"
echo Registration may partially succeed. Check EACH ID in HULFT10.
pause
if not "%RC_SND%"=="0" exit /b 1
if not "%RC_RCV%"=="0" exit /b 1
exit /b 0
