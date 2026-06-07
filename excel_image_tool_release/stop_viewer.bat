@echo off
cd /d "%~dp0"
set "TOOL_ENV_NO_PAUSE=1"
call "%~dp0setup_env.bat"
set "TOOL_ENV_NO_PAUSE="
if errorlevel 1 (
  pause
  exit /b 1
)
"%~dp0ExcelImageTool.exe" --stop
pause
