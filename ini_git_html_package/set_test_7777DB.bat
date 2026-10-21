@echo off
setlocal DisableDelayedExpansion

rem Set the real INI path and the desired database name below.
set "INI_FILE=%~dp0test.ini"
set "TARGET_SECTION=ODBC\u63A5\u7D9A\u6587\u5B57\u5217"
set "TARGET_KEY=Database"
set "NEW_VALUE=99DB"
set "INI_CHARSET=shift_jis"

"%SystemRoot%\System32\cscript.exe" //nologo "%~dp0set_ini_test.vbs" "%INI_FILE%" "%NEW_VALUE%" "%INI_CHARSET%" "%TARGET_SECTION%" "%TARGET_KEY%"
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (echo SUCCESS) else (echo FAILED - see the error above.)
if /i not "%~1"=="/nopause" pause
exit /b %RESULT%
