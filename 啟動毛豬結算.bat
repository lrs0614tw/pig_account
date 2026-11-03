@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto missing_environment
".venv\Scripts\python.exe" -c "import tkinter, openpyxl, docx, win32com.client; tkinter.Tcl()" >nul 2>&1
if errorlevel 1 goto missing_environment
if not exist "毛豬結算\pig_accounting.py" (
    echo 找不到主程式，請先完整解壓縮整個資料夾。
    pause
    exit /b 1
)
cd /d "%~dp0毛豬結算"
"%~dp0.venv\Scripts\python.exe" "pig_accounting.py"
if errorlevel 1 (
    echo.
    echo 程式發生錯誤，請保留上方錯誤訊息。
    pause
    exit /b 1
)
exit /b 0

:missing_environment
echo 請先執行「安裝環境.bat」，完成後再啟動程式。
echo 若資料夾是從其他電腦複製過來，也請重新安裝環境。
pause
exit /b 1
