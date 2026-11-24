@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM HULFT10 for Windows: export 4 types to separate TXT files.
REM ASCII-only BAT: suitable for ANSI or UTF-8 without BOM.
REM The directory of this BAT determines the output folder.
REM If V8 and V10 are both installed, replace command names with full V10 exe paths.
set "UTLIGEN=utligen"
set "UTLLIST=utllist"
set "OUT=%~dp0hulft10_export"
set "LOG=%OUT%\export_result.log"

if not exist "%OUT%" mkdir "%OUT%"
if not exist "%OUT%" (
    echo ERROR: Cannot create "%OUT%".
    pause
    exit /b 1
)

where utligen >nul 2>&1
if errorlevel 1 goto MissingProgram
where utllist >nul 2>&1
if errorlevel 1 goto MissingProgram

REM Never overwrite files in an existing export folder.
if exist "%OUT%\hulft10_source_snd.txt" goto Exists
if exist "%OUT%\hulft10_source_rcv.txt" goto Exists
if exist "%OUT%\hulft10_snd_history.txt" goto Exists
if exist "%OUT%\hulft10_rcv_history.txt" goto Exists

echo ==== HULFT10 export ==== > "%LOG%"
echo Source commands: >> "%LOG%"
where utligen >> "%LOG%" 2>&1
where utllist >> "%LOG%" 2>&1

echo [1/4] SEND management ...
%UTLIGEN% -f "%OUT%\hulft10_source_snd.txt" -i snd -id "*" >> "%LOG%" 2>&1
if errorlevel 1 goto Failure

echo [2/4] RECEIVE management ...
%UTLIGEN% -f "%OUT%\hulft10_source_rcv.txt" -i rcv -id "*" >> "%LOG%" 2>&1
if errorlevel 1 goto Failure

echo [3/4] SEND history (V84) ...
%UTLLIST% -s -v84 > "%OUT%\hulft10_snd_history.txt" 2>> "%LOG%"
if errorlevel 1 goto Failure

echo [4/4] RECEIVE history (V84) ...
%UTLLIST% -r -v84 > "%OUT%\hulft10_rcv_history.txt" 2>> "%LOG%"
if errorlevel 1 goto Failure

echo SUCCESS. Exported files: >> "%LOG%"
dir /b "%OUT%\hulft10_*.txt" >> "%LOG%"
echo.
echo Export complete. Folder: "%OUT%"
echo Review "%LOG%" and data counts.
pause
exit /b 0

:Exists
echo ERROR: One or more output TXT files already exist.
echo Move the previous files or use a NEW export folder.
echo No files were exported this time.
pause
exit /b 2

:MissingProgram
echo ERROR: HULFT CLI command not found in PATH.
echo Run on the HULFT10 server or specify the full HULFT10 exe path.
pause
exit /b 3

:Failure
echo ERROR: Export interrupted. See "%LOG%".
echo Files exported before the error may exist. They are NOT complete backup.
pause
exit /b 1
