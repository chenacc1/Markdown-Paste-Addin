@echo off
chcp 65001 >nul
set SCRIPT=%~dp0md2docx.py
set OUTPUT=%USERPROFILE%\Desktop\paste_result.docx

echo ========================================
echo   MarkdownPasteAddin v3.0
echo   粘贴转 Word
echo ========================================
echo.

echo 正在检测 Python...
echo.

REM Auto-detect Python from common locations
set PYTHON=
for %%p in (python python3) do (
    where %%p >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%i in ('where %%p 2^>nul') do set PYTHON=%%i
        goto :found
    )
)

REM Check common install paths
for %%d in (
    "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python313\python.exe"
    "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe"
    "C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "D:\anaconda3\python.exe"
    "C:\ProgramData\anaconda3\python.exe"
) do (
    if exist %%d (
        set PYTHON=%%~d
        goto :found
    )
)

echo [ERROR] Python not found. Please install Python 3.9+ from https://python.org
echo Or run: winget install python
pause
exit /b 1

:found
echo Python found: %PYTHON%
echo Converting clipboard to Word...
echo.

"%PYTHON%" "%SCRIPT%" "%OUTPUT%" --toc --format 2> "%USERPROFILE%\Desktop\paste_error.log"

if exist "%OUTPUT%" (
    echo.
    echo ========================================
    echo   Done! Opening Word...
    echo ========================================
    start "" "%OUTPUT%"
) else (
    echo.
    echo ========================================
    echo   FAILED
    echo ========================================
    echo.
    echo Error log saved to: %USERPROFILE%\Desktop\paste_error.log
    echo.
    echo Troubleshooting:
    echo   1. Make sure you copied content first (Ctrl+C in DeepSeek)
    echo   2. Run: pip install -r requirements.txt
    echo   3. Try opening GUI instead: python gui_app.py
    echo.
    type "%USERPROFILE%\Desktop\paste_error.log"
)
pause
