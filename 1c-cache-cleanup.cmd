@echo off
rem Runs the 1C cache cleanup script without changing the system execution policy.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp01c-cache-cleanup.ps1" %*
exit /b %ERRORLEVEL%
