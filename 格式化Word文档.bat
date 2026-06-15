@echo off
set SCRIPT=%~dp0format_docx.py

echo ========================================
echo   Word 文档格式统一工具 v3.2
echo ========================================
echo.
echo   将 .docx 文件拖放到本窗口，按 Enter
echo   或者直接输入 .docx 文件路径
echo.
echo   格式预设：
echo     一级标题: 黑体 16pt 加粗 (大纲级别1)
echo     二级标题: 黑体 14pt 加粗 (大纲级别2)
echo     三级标题: 黑体 12pt 加粗 (大纲级别3)
echo     图题:     宋体 10pt 居中
echo     表题:     宋体 10pt 居中加粗
echo     正文:     宋体 12pt 1.5倍行距 首行缩进
echo.
echo   支持中英文格式自动识别
echo   (python format_docx.py --lang en)
echo ========================================
echo.

set /p FILE="文件路径: "
if "%FILE%"=="" (
    echo 未输入文件路径，退出。
    pause
    exit /b
)

D:\anaconda3\python.exe "%SCRIPT%" "%FILE%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo 格式化完成！
) else (
    echo.
    echo 格式化失败，请检查文件路径是否正确。
)
pause
