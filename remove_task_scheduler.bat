@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "schtasks /delete /tn 'FlightPriceTracker_0100' /f; schtasks /delete /tn 'FlightPriceTracker_0900' /f; Write-Host '✅ 已成功移除自動排程！' -ForegroundColor Yellow"
pause
