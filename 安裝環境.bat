@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo 正在準備毛豬結算環境...
echo.

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys, pip; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
    if not errorlevel 1 goto install_packages
)

py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    goto create_environment
)
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto create_environment
)
echo 找不到 Python 3.10 或更新版本。
echo 請先安裝 Python，安裝時勾選 Add python.exe to PATH，
echo 再重新雙擊這個安裝檔。
goto failed

:create_environment
if not exist "requirements.txt" (
    echo 找不到 requirements.txt，請先完整解壓縮整個資料夾。
    goto failed
)
%PYTHON_CMD% -m venv ".venv"
if errorlevel 1 goto failed
:install_packages
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements.txt"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -c "import tkinter, openpyxl, docx, win32com.client; tkinter.Tcl()"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip check
if errorlevel 1 goto failed
echo.
echo 安裝完成！以後請雙擊「啟動毛豬結算.bat」。
echo Excel 與 Word 請使用電腦已安裝的桌面版。
if /i not "%~1"=="--no-pause" pause
exit /b 0

:failed
echo.
echo 環境尚未安裝完成。請查看上方錯誤；下載套件時需要網路。
if /i not "%~1"=="--no-pause" pause
exit /b 1
