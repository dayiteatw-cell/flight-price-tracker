@echo off
chcp 65001 >nul
echo 正在開啟機票價格智慧監控網頁儀表板...
start "" "%~dp0dashboard.html"
exit
