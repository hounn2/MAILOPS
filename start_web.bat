@echo off
chcp 65001
cls
echo ==========================================
echo Outlook邮件助手 - Web控制界面
echo ==========================================
echo.
echo 正在启动Web服务...
echo.
echo 请在浏览器中打开: http://localhost:5000
echo.
echo 按 Ctrl+C 停止服务
echo ==========================================
echo.

python web_app.py

pause
