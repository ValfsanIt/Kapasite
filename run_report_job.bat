@echo off
setlocal

set "ROOT=%~dp0"
pushd "%ROOT%"

if not exist "venv\Scripts\python.exe" (
  echo [run_report_job] HATA: venv\Scripts\python.exe bulunamadi.
  popd
  exit /b 1
)

if not exist "logs" mkdir "logs"

rem HIZLI TEST MODU kapali (tum costcenter)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "LOG_FILE=%ROOT%logs\report_job_%TS%.log"

echo [run_report_job] Basladi: %DATE% %TIME%
echo [run_report_job] Log: %LOG_FILE%

"venv\Scripts\python.exe" "report_job.py" >> "%LOG_FILE%" 2>&1
set "ERR=%ERRORLEVEL%"

if "%ERR%"=="0" (
  echo [run_report_job] Tamamlandi. exit_code=0
) else (
  echo [run_report_job] HATA. exit_code=%ERR%
)

popd
exit /b %ERR%
