@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start "" "%~f0" run & exit)
:run
title Ribbon: Force Import (ribbon\ -> Dev XLSM)
cd /d "%~dp0"
echo ============================================================
echo  Force Import: ribbon\ ^> MORFunctions-dev.xlsm
echo ============================================================
echo.
"%PY_EXEC%" scripts\ribbon_sync.py import "MORFunctions-dev.xlsm" --input-dir ribbon
echo.
echo === Done. Press any key to close ===
pause
