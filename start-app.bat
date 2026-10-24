@echo off
cd /d "%~dp0"
call npm run app
if %ERRORLEVEL% NEQ 0 (
  echo Error launching application: %ERRORLEVEL%
  pause
)
