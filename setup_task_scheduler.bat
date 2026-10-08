@echo off
schtasks /create /tn "FlightPriceTracker_0100" /tr "\"%~dp0run_scheduled.bat\"" /sc daily /st 01:00 /f >nul
schtasks /create /tn "FlightPriceTracker_0900" /tr "\"%~dp0run_scheduled.bat\"" /sc daily /st 09:00 /f >nul
echo.
echo ========================================================
echo   Windows Task Scheduler Installed Successfully!
echo   [1] Daily 01:00 AM (Midnight Flight Release)
echo   [2] Daily 09:00 AM (Morning Price Adjustment)
echo ========================================================
echo.
pause
