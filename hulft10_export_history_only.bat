@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM HULFT10 Windows - export SEND / RECEIVE histories only (-v84).
REM Does not change HULFT settings or delete history.
REM Run on the HULFT10 server / an authorized HULFT10 command environment.
REM Save as ANSI (CP932) or UTF-8 WITHOUT BOM.
REM If HULFT8 and 10 coexist, use the full HULFT10 executable path:
REM set "UTLLIST=C:\path\to\HULFT10\bin\utllist.exe"
set "UTLLIST=utllist"
set "ROOT=%~dp0hulft10_history_runs"
set "RUN=1"

REM Find utllist and log the actual Windows PATH resolution.
if not exist "%UTLLIST%" (
    where "%UTLLIST%" >nul 2>&1
    if errorlevel 1 goto NotFound
)

if not exist "%ROOT%" mkdir "%ROOT%"
if not exist "%ROOT%" (
    echo ERROR: Cannot create output root: "%ROOT%"
    pause
    exit /b 3
)

:FindNewFolder
set "DIR=%ROOT%\run_%RUN%"
if not exist "%DIR%" goto CreateFolder
set /a RUN+=1
if %RUN% GTR 9999 goto TooMany
goto FindNewFolder

:CreateFolder
mkdir "%DIR%" 2>nul
if not exist "%DIR%" (
    echo ERROR: Cannot create "%DIR%"
    pause
    exit /b 3
)

set "SEND=%DIR%\hulft10_snd_history.txt"
set "RECV=%DIR%\hulft10_rcv_history.txt"
set "ERR_SEND=%DIR%\send_stderr.txt"
set "ERR_RECV=%DIR%\receive_stderr.txt"
set "LOG=%DIR%\history_result.log"

> "%LOG%" echo ==== HULFT10 history export diagnostic ====
>> "%LOG%" echo Output folder: "%DIR%"
>> "%LOG%" echo Configured utllist: "%UTLLIST%"

if exist "%UTLLIST%" (
    >> "%LOG%" echo Executable: "%UTLLIST%"
) else (
    >> "%LOG%" echo PATH candidate(s):
    where "%UTLLIST%" >> "%LOG%" 2>&1
)

echo.
echo [1/2] Exporting SEND history ...
>> "%LOG%" echo [SEND] utllist -s -v84
"%UTLLIST%" -s -v84 1> "%SEND%" 2> "%ERR_SEND%"
set "SEND_RC=%ERRORLEVEL%"
>> "%LOG%" echo SEND exit code: %SEND_RC%

echo [2/2] Exporting RECEIVE history ...
>> "%LOG%" echo [RECEIVE] utllist -r -v84
"%UTLLIST%" -r -v84 1> "%RECV%" 2> "%ERR_RECV%"
set "RECV_RC=%ERRORLEVEL%"
>> "%LOG%" echo RECEIVE exit code: %RECV_RC%

echo.>> "%LOG%"
if exist "%SEND%" (
    for %%F in ("%SEND%") do call :ReportSize SEND "%%~zF"
) else (
    >> "%LOG%" echo SEND output file missing.
)
if exist "%RECV%" (
    for %%F in ("%RECV%") do call :ReportSize RECEIVE "%%~zF"
) else (
    >> "%LOG%" echo RECEIVE output file missing.
)

>> "%LOG%" echo.
>> "%LOG%" echo ==== SEND stderr ====
type "%ERR_SEND%" >> "%LOG%"
>> "%LOG%" echo.
>> "%LOG%" echo ==== RECEIVE stderr ====
type "%ERR_RECV%" >> "%LOG%"

echo.
echo SEND exit code    : %SEND_RC%
echo RECEIVE exit code : %RECV_RC%
echo Files and diagnostic log:
echo "%DIR%"
echo.
if not "%SEND_RC%"=="0" (
    echo WARNING: SEND history export failed.
    echo Inspect send_stderr.txt and hulft10_snd_history.txt.
)
if not "%RECV_RC%"=="0" (
    echo WARNING: RECEIVE history export failed.
    echo Inspect receive_stderr.txt and hulft10_rcv_history.txt.
)
if "%SEND_RC%"=="0" if "%RECV_RC%"=="0" (
    echo Both commands returned success.
    echo Check TXT contents: empty history files can mean no retained records.
)
echo.
echo Review history_result.log, then press any key.
pause >nul
if not "%SEND_RC%"=="0" exit /b 1
if not "%RECV_RC%"=="0" exit /b 1
exit /b 0

:ReportSize
>> "%LOG%" echo %~1 output bytes: %~2
if "%~2"=="0" (
    >> "%LOG%" echo WARNING: %~1 output is 0 bytes.
    echo WARNING: %~1 history output is 0 bytes.
)
exit /b 0

:NotFound
echo ERROR: utllist was not found.
echo Run on HULFT10 Windows or change UTLLIST to the full EXE path.
pause
exit /b 2

:TooMany
echo ERROR: Too many history export folders.
pause
exit /b 3
