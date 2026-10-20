@echo off
setlocal DisableDelayedExpansion

rem Set the target INI path, section, key, value, and charset below.
set "INI_FILE=%~dp0test.ini"
rem Japanese section name [Setto]. Unicode escapes avoid BAT code-page problems.
set "TARGET_SECTION=\u30BB\u30C3\u30C8"
set "TARGET_KEY=case"
set "NEW_VALUE=99DB"
set "INI_CHARSET=utf-8"

"%SystemRoot%\System32\cscript.exe" //nologo "%~dp0set_ini_test.vbs" "%INI_FILE%" "%NEW_VALUE%" "%INI_CHARSET%" "%TARGET_SECTION%" "%TARGET_KEY%"
set "RESULT=%ERRORLEVEL%"

echo.
if "%RESULT%"=="0" (echo SUCCESS) else (echo FAILED - see the error above.)
rem Use /nopause when calling this BAT from another BAT or a job.
if /i not "%~1"=="/nopause" pause
exit /b %RESULT%
