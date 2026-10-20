@echo off
setlocal DisableDelayedExpansion

rem Edit these settings. The default INI is next to this BAT.
rem set "INI_FILE=%~dp0test.ini"
set "INI_FILE=%~dp0test.ini"
set "NEW_VALUE=99DB"
set "INI_CHARSET=utf-8"

"%SystemRoot%\System32\cscript.exe" //nologo "%~dp0set_ini_test.vbs" "%INI_FILE%" "%NEW_VALUE%" "%INI_CHARSET%"
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (echo SUCCESS) else (echo FAILED - see the error above.)
rem Use /nopause when calling this BAT from another BAT or a job.
if /i not "%~1"=="/nopause" pause
exit /b %RESULT%
