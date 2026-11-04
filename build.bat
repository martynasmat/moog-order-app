@echo off
setlocal
cd /d "%~dp0"
set "PYINSTALLER_CONFIG_DIR=%~dp0build\pyinstaller-cache"

".venv\Scripts\python.exe" -m pip install -e ".[build]"
if errorlevel 1 goto failed

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name MoogOrders --specpath build --paths . --hidden-import win32timezone gui\__main__.py
if errorlevel 1 goto failed

echo.
echo Ready: dist\MoogOrders.exe
pause
exit /b 0

:failed
echo.
echo Build failed. See the error above.
pause
exit /b 1
