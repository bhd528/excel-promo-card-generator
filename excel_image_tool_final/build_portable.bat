@echo off
cd /d "%~dp0"
python -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
  echo PyInstaller not found. Installing build dependency...
  python -m pip install -r requirements-build.txt
  if errorlevel 1 (
    echo Failed to install build dependency.
    pause
    exit /b 1
  )
) else (
  echo PyInstaller already installed. Skip install.
)
python -m PyInstaller --noconfirm --clean --onefile --name ExcelImageTool launch_viewer.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
copy /Y dist\ExcelImageTool.exe ExcelImageTool.exe
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist
if exist ExcelImageTool.spec del /F /Q ExcelImageTool.spec
if exist __pycache__ rmdir /S /Q __pycache__
echo Portable exe created: ExcelImageTool.exe
pause
