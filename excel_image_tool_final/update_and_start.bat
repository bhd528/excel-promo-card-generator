@echo off
cd /d "%~dp0"
set "TOOL_ENV_NO_PAUSE=1"
call "%~dp0setup_env.bat"
set "TOOL_ENV_NO_PAUSE="
if errorlevel 1 (
  pause
  exit /b 1
)
if exist "%~dp0ExcelImageTool.exe" (
  "%~dp0ExcelImageTool.exe" --update
) else (
  python launch_viewer.py --update
)
pause
