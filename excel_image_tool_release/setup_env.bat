@echo off
cd /d "%~dp0"
set "SETUP_STATUS=0"

echo ================================
echo Excel Image Tool - Environment
echo ================================
echo Checking runtime...

if exist "%~dp0ExcelImageTool.exe" (
  "%~dp0ExcelImageTool.exe" --help | findstr /C:"--update" >nul
  if errorlevel 1 (
    echo.
    echo [FAILED] Environment is not ready.
    echo ExcelImageTool.exe exists, but it is not the correct version.
    echo Please replace it with the latest release package.
    set "SETUP_STATUS=1"
    goto done
  )
  echo.
  echo [OK] Environment is ready.
  echo Found ExcelImageTool.exe. Python is not required.
  echo.
  echo Next step: double-click start_viewer.bat
  goto done
)

echo.
echo [FAILED] Environment is not ready.
echo Missing ExcelImageTool.exe.
echo Please use the complete release package.
set "SETUP_STATUS=1"

:done
if not "%TOOL_ENV_NO_PAUSE%"=="1" pause
exit /b %SETUP_STATUS%
