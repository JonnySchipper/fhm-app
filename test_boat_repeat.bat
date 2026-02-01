@echo off
echo ========================================
echo Boat Test - Auto Repeat Mode
echo ========================================
echo.
echo This will run the boat test continuously.
echo Edit test_boat.py and save - it will auto-regenerate!
echo Press Ctrl+C to stop
echo.
pause

:loop
cls
echo ========================================
echo Running boat test...
echo ========================================
python test_boat.py
echo.
echo Waiting 2 seconds before next run...
timeout /t 2 /nobreak >nul
goto loop

