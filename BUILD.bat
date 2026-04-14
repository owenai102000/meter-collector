@echo off
echo ================================================
echo Canon Meter Collector - Windows Build Script
echo ================================================
echo.

echo Step 1: Installing Python libraries...
pip install requests pyinstaller

echo.
echo Step 2: Building .exe (this may take a few minutes)...
pyinstaller --onefile --windowed --name "MeterCollector" windows-meter-collector.py

echo.
echo ================================================
echo Build Complete!
echo ================================================
echo.
echo Your .exe is in the dist folder:
cd dist
echo.
echo To configure, edit these values in windows-meter-collector.py:
echo   - FIREBASE_URL
echo   - FIREBASE_SECRET
echo   - CUSTOMER_ID
echo   - COPIER_IP
echo   - COPIER_PORT
echo   - COPIER_USERNAME
echo   - COPIER_PASSWORD
echo.
pause
