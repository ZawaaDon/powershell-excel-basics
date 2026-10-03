@echo off
rem ============================================================
rem step8 : a launcher bat (the entrance only)
rem ------------------------------------------------------------
rem   See the header of step8.ps1 for the explanation.
rem
rem   Keep the bat short: it only starts the script.
rem   The real work is in the .ps1 files (ps calls ps),
rem   so no value comes back to the bat and for /f is not needed.
rem
rem   pwsh            : PowerShell 7 (powershell = Windows PowerShell 5.1)
rem   %*              : all arguments given to this bat, passed as they are
rem   -ExecutionPolicy Bypass : run the script even if scripts are blocked
rem                     on this PC (this run only; the PC setting stays)
rem   pause           : keep the window open to read the result
rem
rem   Comments are in English: Japanese text in a bat file
rem   may be garbled depending on the code page.
rem ============================================================

rem --- the launcher: this part is all you need ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step8.ps1" %*
echo exit code = %ERRORLEVEL%

rem --- for learning: the same script with PowerShell 5.1 ---
rem step8.ps1 has "#Requires -Version 7", so 5.1 stops before the first line.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0step8.ps1" %*
echo exit code = %ERRORLEVEL%

pause
