@echo off
title MarkdownPasteAddin Bridge Server

echo ========================================
echo   MarkdownPasteAddin 桥接服务 v2.0
echo ========================================
echo.
echo   用于 Chrome 扩展连接
echo   保持本窗口打开，不要关闭
echo.
echo   服务地址: http://localhost:9876
echo   健康检查: http://localhost:9876/api/health
echo.
echo   按 Ctrl+C 停止服务
echo ========================================
echo.

D:\anaconda3\python.exe "%~dp0bridge_server.py" %*

pause
