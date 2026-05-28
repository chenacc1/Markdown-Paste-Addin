@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   批量 Markdown → Word 转换工具
echo ========================================
echo.
echo   将包含 .md 文件的文件夹拖放到本窗口
echo   或直接输入文件夹路径
echo ========================================
echo.

set /p INPUT="输入文件夹: "
if "%INPUT%"=="" (
    echo 未输入路径，退出。
    pause
    exit /b
)

set /p OUTPUT="输出文件夹 (默认: %INPUT%\docx): "
if "%OUTPUT%"=="" set OUTPUT=%INPUT%\docx

echo.
echo Converting all .md files in: %INPUT%
echo Output to: %OUTPUT%
echo.

D:\anaconda3\python.exe "%~dp0batch_convert.py" "%INPUT%" "%OUTPUT%"

echo.
pause
