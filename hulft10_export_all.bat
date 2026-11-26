@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM HULFT10 Windows - ALL FOUR types, with diagnostic output.
REM Save this BAT as ANSI/CP932 or UTF-8 without BOM (ASCII contents).
REM This is an EXPORT ONLY: HULFT settings and history are not changed.
REM Always check that the resolved utllist/utligen are for HULFT10.
REM If both HULFT8 and HULFT10 are installed, use their FULL V10 EXE paths.
set "UTLIGEN=utligen"
set "UTLLIST=utllist"
set "ROOT=%~dp0hulft10_export_runs"
set "RUN=1"

if not exist "%UTLIGEN%" (
    where "%UTLIGEN%" >nul 2>&1
    if errorlevel 1 goto NoExe
)
if not exist "%UTLLIST%" (
    where "%UTLLIST%" >nul 2>&1
    if errorlevel 1 goto NoExe
)

if not exist "%ROOT%" mkdir "%ROOT%"
if not exist "%ROOT%" (
    echo ERROR: Unable to create output directory: "%ROOT%"
    pause
    exit /b 3
)

:NextRun
set "OUT=%ROOT%\run_%RUN%"
if not exist "%OUT%" goto NewRun
set /a RUN+=1
if %RUN% GTR 9999 goto FolderLimit
goto NextRun

:NewRun
mkdir "%OUT%" 2>nul
if not exist "%OUT%" (
    echo ERROR: Cannot create "%OUT%"
    pause
    exit /b 3
)

set "LOG=%OUT%\export_result.log"
> "%LOG%" echo ==== HULFT10 Export (4 types) ====
>> "%LOG%" echo Output directory: "%OUT%"
>> "%LOG%" echo UT LIGEN selected: "%UTLIGEN%"
>> "%LOG%" echo UT LLIST selected: "%UTLLIST%"
>> "%LOG%" echo ----- PATH resolution -----
if exist "%UTLIGEN%" (
    >> "%LOG%" echo utligen = "%UTLIGEN%"
) else (
    where "%UTLIGEN%" >> "%LOG%" 2>&1
)
if exist "%UTLLIST%" (
    >> "%LOG%" echo utllist = "%UTLLIST%"
) else (
    where "%UTLLIST%" >> "%LOG%" 2>&1
)

echo [1/4] SEND management ...
"%UTLIGEN%" -f "%OUT%\hulft10_source_snd.txt" -i snd -id "*" 1> "%OUT%\snd_mng_stdout.txt" 2> "%OUT%\snd_mng_stderr.txt"
set "RC_MS=%ERRORLEVEL%"
>> "%LOG%" echo SEND management exit code: %RC_MS%

echo [2/4] RECEIVE management ...
"%UTLIGEN%" -f "%OUT%\hulft10_source_rcv.txt" -i rcv -id "*" 1> "%OUT%\rcv_mng_stdout.txt" 2> "%OUT%\rcv_mng_stderr.txt"
set "RC_MR=%ERRORLEVEL%"
>> "%LOG%" echo RECEIVE management exit code: %RC_MR%

echo [3/4] SEND history (-v84) ...
"%UTLLIST%" -s -v84 1> "%OUT%\hulft10_snd_history.txt" 2> "%OUT%\snd_hist_stderr.txt"
set "RC_HS=%ERRORLEVEL%"
>> "%LOG%" echo SEND history exit code: %RC_HS%

echo [4/4] RECEIVE history (-v84) ...
set "RECV=%OUT%\hulft10_rcv_history.txt"
set "ERR_RECV=%OUT%\rcv_hist_stderr.txt"
call :ReceiveToFile
set "RC_HR=%ERRORLEVEL%"
>> "%LOG%" echo RECEIVE history exit code: %RC_HR%

>> "%LOG%" echo ----- File size / bytes -----
if exist "%OUT%\hulft10_source_snd.txt" (
    for %%F in ("%OUT%\hulft10_source_snd.txt") do >> "%LOG%" echo SEND management: %%~zF
) else (
    >> "%LOG%" echo SEND management file missing
)
if exist "%OUT%\hulft10_source_rcv.txt" (
    for %%F in ("%OUT%\hulft10_source_rcv.txt") do >> "%LOG%" echo RECEIVE management: %%~zF
) else (
    >> "%LOG%" echo RECEIVE management file missing
)
if exist "%OUT%\hulft10_snd_history.txt" (
    for %%F in ("%OUT%\hulft10_snd_history.txt") do >> "%LOG%" echo SEND history: %%~zF
) else (
    >> "%LOG%" echo SEND history file missing
)
if exist "%OUT%\hulft10_rcv_history.txt" (
    for %%F in ("%OUT%\hulft10_rcv_history.txt") do >> "%LOG%" echo RECEIVE history: %%~zF
) else (
    >> "%LOG%" echo RECEIVE history file missing
)

>> "%LOG%" echo ----- SEND history STDERR -----
type "%OUT%\snd_hist_stderr.txt" >> "%LOG%"
>> "%LOG%" echo ----- RECEIVE history diagnostics (stderr on console) -----
type "%OUT%\rcv_hist_stderr.txt" >> "%LOG%"
>> "%LOG%" echo ----- SEND management STDERR -----
type "%OUT%\snd_mng_stderr.txt" >> "%LOG%"
>> "%LOG%" echo ----- RECEIVE management STDERR -----
type "%OUT%\rcv_mng_stderr.txt" >> "%LOG%"

echo.
echo SEND management   : %RC_MS%
echo RECEIVE management: %RC_MR%
echo SEND history      : %RC_HS%
echo RECEIVE history   : %RC_HR%
echo.
echo Output directory: "%OUT%"
echo Check export_result.log and exported TXT before using Excel.
if not "%RC_MS%"=="0" echo WARNING: SEND management command failed.
if not "%RC_MR%"=="0" echo WARNING: RECEIVE management command failed.
if not "%RC_HS%"=="0" echo WARNING: SEND history command failed.
if not "%RC_HR%"=="0" echo WARNING: RECEIVE history command failed.
echo.
pause

if not "%RC_MS%"=="0" exit /b 1
if not "%RC_MR%"=="0" exit /b 1
if not "%RC_HS%"=="0" exit /b 1
if not "%RC_HR%"=="0" exit /b 1
exit /b 0

:NoExe
echo ERROR: Required utligen or utllist command was not found.
echo Use the full HULFT10 EXE paths in the UT LIGEN / UT LLIST variables.
pause
exit /b 2

:FolderLimit
echo ERROR: No unused run folder remains.
pause
exit /b 3

:ReceiveToFile
REM Use the same stdout-only redirection as the successful CMD test.
REM Stage in TEMP, then copy the result to the normal output folder.
REM stderr remains on screen, exactly as in the manual test.
> "%ERR_RECV%" echo RECEIVE diagnostics: stderr is displayed on the console.
if not defined TEMP (
    >> "%ERR_RECV%" echo ERROR: TEMP is not defined.
    exit /b 2001
)
set "RCV_STAGE=%TEMP%\hulft10_rcv_%RANDOM%_%RANDOM%"
mkdir "%RCV_STAGE%" 2>nul
if errorlevel 1 (
    >> "%ERR_RECV%" echo ERROR: Cannot create temporary folder "%RCV_STAGE%".
    exit /b 2001
)
set "RCV_TMP=%RCV_STAGE%\hulft10_rcv_test.txt"
>> "%ERR_RECV%" echo Staging file: "%RCV_TMP%"
echo Temporary receive output: "%RCV_TMP%"
if /i "%UTLLIST%"=="utllist" (
    utllist -r -v84 > "%RCV_TMP%"
) else (
    "%UTLLIST%" -r -v84 > "%RCV_TMP%"
)
set "RCV_CMD_RC=%ERRORLEVEL%"
>> "%ERR_RECV%" echo utllist receive exit code: %RCV_CMD_RC%
if not "%RCV_CMD_RC%"=="0" (
    echo RECEIVE command failed. Check the console and temporary file.
    >> "%ERR_RECV%" echo Failed; temporary files retained for diagnosis.
    exit /b %RCV_CMD_RC%
)
if not exist "%RCV_TMP%" (
    >> "%ERR_RECV%" echo ERROR: Temporary output file is missing.
    exit /b 2002
)
for %%F in ("%RCV_TMP%") do set "RCV_BYTES=%%~zF"
>> "%ERR_RECV%" echo Temporary output bytes: %RCV_BYTES%
if "%RCV_BYTES%"=="0" (
    >> "%ERR_RECV%" echo WARNING: No receive history bytes; temporary file retained.
    echo WARNING: RECEIVE temporary file is empty.
    exit /b 2002
)
if exist "%RECV%" (
    >> "%ERR_RECV%" echo ERROR: Destination already exists; not overwritten.
    exit /b 2003
)
copy /b /y "%RCV_TMP%" "%RECV%" >> "%ERR_RECV%" 2>&1
if errorlevel 1 (
    >> "%ERR_RECV%" echo ERROR: Copy failed; temporary file retained.
    echo RECEIVE copy failed. Source retained: "%RCV_TMP%"
    exit /b 2003
)
if not exist "%RECV%" exit /b 2003
for %%F in ("%RECV%") do set "RCV_COPIED_BYTES=%%~zF"
>> "%ERR_RECV%" echo Destination output bytes: %RCV_COPIED_BYTES%
if not "%RCV_BYTES%"=="%RCV_COPIED_BYTES%" (
    >> "%ERR_RECV%" echo ERROR: Copy size differs; temporary file retained.
    exit /b 2003
)
del /q "%RCV_TMP%" >nul 2>&1
rd "%RCV_STAGE%" >nul 2>&1
if exist "%RCV_STAGE%" >> "%ERR_RECV%" echo WARNING: Check remaining temporary files.
>> "%ERR_RECV%" echo RECEIVE export and copy completed.
echo RECEIVE history saved: "%RECV%"
exit /b 0
