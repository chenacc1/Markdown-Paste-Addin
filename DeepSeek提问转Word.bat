@echo off
echo ========================================
echo   DeepSeek API → Word 转换工具
echo ========================================
echo.
echo   直接向 DeepSeek 提问，答案自动转为 Word
echo   需要设置 DEEPSEEK_API_KEY 环境变量
echo ========================================
echo.

if "%DEEPSEEK_API_KEY%"=="" (
    echo [警告] DEEPSEEK_API_KEY 未设置！
    echo.
    echo 请先设置环境变量:
    echo   set DEEPSEEK_API_KEY=sk-xxxx
    echo.
    echo 或在命令行指定:
    echo   python deepseek_api.py --api-key sk-xxxx "问题" output.docx
    echo.
)

set /p QUESTION="请输入问题: "
if "%QUESTION%"=="" (
    echo 未输入问题，退出。
    pause
    exit /b
)

set OUTPUT=%USERPROFILE%\Desktop\deepseek-answer.docx

echo.
echo Asking DeepSeek...
echo.

D:\anaconda3\python.exe "%~dp0deepseek_api.py" "%QUESTION%" "%OUTPUT%"

if exist "%OUTPUT%" (
    echo.
    echo Done! Opening Word...
    start "" "%OUTPUT%"
) else (
    echo.
    echo 转换失败，请检查 API key 和网络连接。
)
pause
