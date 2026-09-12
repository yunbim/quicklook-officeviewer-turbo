@echo off
rem =====================================================================
rem  QuickLook Office Turbo - rollback launcher
rem  (Chinese: see INSTALL.txt; it is bilingual)
rem
rem  Double-click this file to restore the ORIGINAL OfficeViewer plugin.
rem  The actual work is done by rollback.ps1, which sits next to it.
rem  This file is ASCII-only so it renders in every console code page.
rem =====================================================================
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0rollback.ps1" %*
set "RC=%ERRORLEVEL%"

echo.
pause
exit /b %RC%
