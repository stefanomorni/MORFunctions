@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start "" "%~f0" run & exit)
:run
title PQ: Force Import (power-queries -> Dev XLSM)
cd /d "%~dp0"
echo ============================================================
echo  Force Import: power-queries\ ^> MORFunctions-dev
echo  (Add-ins typically have no PQ — launcher provisioned for
echo   uniformity)
echo ============================================================
echo.
"%PY_EXEC%" scripts\pq_sync.py import-all "MORFunctions-dev.xlsm"
echo.
echo === Done. Press any key to close ===
pause
