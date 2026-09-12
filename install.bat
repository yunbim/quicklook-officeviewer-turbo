@echo off
rem =====================================================================
rem  QuickLook Office Turbo - installer launcher
rem  (Chinese: see INSTALL.txt; it is bilingual)
rem
rem  Double-click this file. The actual work is done by install.ps1, which
rem  sits next to it; this launcher only makes sure the script runs with the
rem  right policy (no profile, execution policy bypassed for this one run).
rem
rem  The logic lives in PowerShell on purpose: Windows always ships it, so
rem  nothing needs to be installed - and it is real, testable code instead
rem  of a pile of fragile cmd.exe quoting. This file is ASCII-only so it
rem  renders correctly in every console code page.
rem =====================================================================
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
set "RC=%ERRORLEVEL%"

echo.
pause
exit /b %RC%
