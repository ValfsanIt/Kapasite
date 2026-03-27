@echo off
setlocal

cd /d "%~dp0"
set "KAPASITE_URL=http://172.30.134.9:8050/"

echo KAPASITE sunucu modu baslatiliyor...
echo Adres: %KAPASITE_URL%
echo.

for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":8050 .*LISTENING"') do (
    taskkill /F /PID %%P >nul 2>&1
)

if exist ".\venv\Scripts\python.exe" (
    ".\venv\Scripts\python.exe" -c "from app import app; app.run(host='0.0.0.0', port=8050, debug=True, use_reloader=False)"
) else (
    python -c "from app import app; app.run(host='0.0.0.0', port=8050, debug=True, use_reloader=False)"
)

echo.
echo Uygulama kapandi. Cikmak icin bir tusa basin.
pause >nul
