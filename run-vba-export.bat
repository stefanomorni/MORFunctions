@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start "" "%~f0" run & exit)
:run
title VBA: Force Export (Dev XLSM -> vba-files)
cd /d "%~dp0"
echo ============================================================
echo  Force Export: MORFunctions-dev VBProject ^> vba-files\
echo  NOTE: Target must be MORFunctions-dev.xlsm, NOT the .xlam
echo ============================================================
echo.
"%PY_EXEC%" scripts\vba_sync_watch.py export-all "MORFunctions-dev.xlsm"
echo.
echo === Done. Press any key to close ===
pause
