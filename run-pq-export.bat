@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start "" "%~f0" run & exit)
:run
title PQ: Force Export (Dev XLSM -> power-queries)
cd /d "%~dp0"
echo ============================================================
echo  Force Export: Queries ^> power-queries\
echo  (Add-ins typically have no PQ — launcher provisioned for
echo   uniformity; expect zero queries)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync.py export-all "MORFunctions-dev.xlsm"
echo.
echo === Done. Press any key to close ===
pause
