@echo off
rem ============================================================
rem step6 : call a PowerShell script from a bat file
rem ------------------------------------------------------------
rem   Same calls as step6.ps1, written as a bat file.
rem   See the header of step6.ps1 for the explanation.
rem
rem   %~dp0           : folder of this bat file (ends with \)
rem   -File           : the script to run; its arguments follow
rem   -ExecutionPolicy Bypass : allow the script for this run only
rem   -NoProfile      : do not load the user profile
rem   %ERRORLEVEL%    : exit code of the last command (0 = OK)
rem                     set by cmd itself: do not "set ERRORLEVEL=..."
rem
rem   Comments are in English: Japanese text in a bat file
rem   may be garbled depending on the code page.
rem ============================================================

rem --- no arguments ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step6-sub.ps1"
echo exit code = %ERRORLEVEL%

rem --- with arguments: by position (Addr, Text) ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step6-sub.ps1" G2 "from bat"
echo exit code = %ERRORLEVEL%

rem --- with arguments: by name ---
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step6-sub.ps1" -Text "by name" -Addr G3
echo exit code = %ERRORLEVEL%

rem --- get a value back from the script ---
rem step6-macro.ps1 runs macros in macros.xlsm and macros.xlam,
rem and prints one line: (5 + 6) + 10 + 0.5 = 21.5
rem for /f puts the printed line into RESULT.
rem If the script fails, nothing is printed and RESULT stays empty.
set RESULT=
for /f "usebackq delims=" %%A in (`pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0step6-macro.ps1" 5 6`) do set RESULT=%%A
if defined RESULT (
    echo result = %RESULT%
) else (
    echo result = [failed]
)
rem Note: ERRORLEVEL can NOT tell success here.
rem The command inside for /f runs in another cmd, and its exit code
rem is thrown away. ERRORLEVEL keeps the value from before the for /f.
rem That is why we check whether RESULT is empty.

rem --- same as above, shorter: put the common part in a variable ---
set PS=pwsh -NoProfile -ExecutionPolicy Bypass -File
set RESULT=
for /f "usebackq delims=" %%A in (`%PS% "%~dp0step6-macro.ps1" 5 6`) do set RESULT=%%A
if defined RESULT (
    echo result = %RESULT%
) else (
    echo result = [failed]
)

rem --- get a value back AND the exit code: write it to a file ---
rem pwsh runs directly (not inside for /f), so ERRORLEVEL is its exit code.
rem Save ERRORLEVEL right away: the next command overwrites it.
rem set /p reads the first line of the file into RESULT.
set RESULT=
%PS% "%~dp0step6-macro.ps1" 5 6 > "%TEMP%\step6-result.txt"
set EXITCODE=%ERRORLEVEL%
set /p RESULT=<"%TEMP%\step6-result.txt"
del "%TEMP%\step6-result.txt"
if %EXITCODE%==0 (
    echo result = %RESULT%
) else (
    echo result = [failed] exit code = %EXITCODE%
)

pause
