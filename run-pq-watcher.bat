@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start /min "" "%~f0" run & exit)
:run
title PQ-Watcher - MORFunctions-dev
cd /d "%~dp0"
echo ============================================================
echo  PQ Watcher: MORFunctions-dev.xlsm ^<^> power-queries\
echo  (Add-ins typically have no PQ - provisioned for uniformity)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync_watch.py "MORFunctions-dev.xlsm" --pq-dir power-queries --sync-script scripts\pq_sync.py
echo.
echo === Watcher stopped. Press any key to close ===
pause
