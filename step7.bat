@echo off
rem ============================================================
rem step7 : check the exit code of a PowerShell script
rem ------------------------------------------------------------
rem   See the header of step7.ps1 for the explanation.
rem
rem   %ERRORLEVEL% is 0 when the script ends normally (or "exit 0"),
rem   and 1 when the script stops with an error that no catch received.
rem   The script must set $ErrorActionPreference = "Stop" for that;
rem   without it, the script goes on after an error and ends with 0.
rem   "if errorlevel 1" means "exit code is 1 or more".
rem
rem   Do not use %ERRORLEVEL% inside ( ) of an if block:
rem   it is replaced before the block runs. Copy it to a variable first.
rem
rem   Comments are in English: Japanese text in a bat file
rem   may be garbled depending on the code page.
rem ============================================================

rem --- a script that handles its errors and ends with "exit 0" ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step7.ps1"
set CODE=%ERRORLEVEL%
if errorlevel 1 (
    echo step7.ps1 : FAILED, exit code = %CODE%
) else (
    echo step7.ps1 : OK, exit code = %CODE%
)

rem --- a script that stops with an error (invalid cell address) ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step6-sub.ps1" "???" "fail"
set CODE=%ERRORLEVEL%
if errorlevel 1 (
    echo step6-sub.ps1 : FAILED, exit code = %CODE%
) else (
    echo step6-sub.ps1 : OK, exit code = %CODE%
)

pause
