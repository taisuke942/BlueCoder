@echo off
setlocal EnableExtensions DisableDelayedExpansion
REM HULFT10 optional dependency export: host/group/job/format definitions.
REM Do NOT assume all IDs exist. Exit codes must be checked.
REM For full environment migration, also review sch/trg/mail as applicable.
set "OUT=%~dp0hulft10_dependencies"
if not exist "%OUT%" mkdir "%OUT%"
if not exist "%OUT%" (
  echo ERROR: Cannot create output directory.
  pause
  exit /b 1
)
where utligen >nul 2>&1
if errorlevel 1 (
  echo ERROR: utligen not found in PATH.
  pause
  exit /b 1
)
if exist "%OUT%\host.txt" goto Exists
if exist "%OUT%\tgrp.txt" goto Exists
if exist "%OUT%\job.txt" goto Exists
if exist "%OUT%\fmt.txt" goto Exists
if exist "%OUT%\mfmt.txt" goto Exists

echo Exporting HOST information...
utligen -f "%OUT%\host.txt" -i hst -id "*"
if errorlevel 1 goto Failed
echo Exporting TRANSFER GROUP information...
utligen -f "%OUT%\tgrp.txt" -i tgrp -id "*"
if errorlevel 1 goto Failed
echo Exporting JOB information...
utligen -f "%OUT%\job.txt" -i job -id "*"
if errorlevel 1 goto Failed
echo Exporting FORMAT information...
utligen -f "%OUT%\fmt.txt" -i fmt -id "*"
if errorlevel 1 goto Failed
echo Exporting MULTI FORMAT information...
utligen -f "%OUT%\mfmt.txt" -i mfmt -id "*"
if errorlevel 1 goto Failed
echo Done. Folder: "%OUT%"
pause
exit /b 0

:Exists
echo ERROR: Previous output exists. Use a new folder; no files overwritten.
pause
exit /b 2

:Failed
echo ERROR: One step failed; previous outputs might be partial.
pause
exit /b 1
