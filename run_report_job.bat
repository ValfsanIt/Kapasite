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

rem Uretimde her zaman tum costcenter (HEPSI) calissin.
rem Ortamda daha once KAP_REPORT_SINGLE_CC tanimliysa burada sifirlanir.
set "KAP_REPORT_SINGLE_CC="

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "LOG_FILE=%ROOT%logs\report_job_%TS%.log"

echo [run_report_job] Basladi: %DATE% %TIME%
echo [run_report_job] Log kapali. Cikti terminalde gosterilecek.
set "PYTHONUNBUFFERED=1"

"venv\Scripts\python.exe" -u "report_job.py"
set "ERR=%ERRORLEVEL%"

if "%ERR%"=="0" (
  echo [run_report_job] Tamamlandi. exit_code=0
) else (
  echo [run_report_job] HATA. exit_code=%ERR%
)

popd
exit /b %ERR%
