@echo off
for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Resolve-OfficePython.ps1" -ProjectRoot "%~dp0."`) do set "PY_EXEC=%%i"
if not defined PY_EXEC (
  echo [XX] Could not resolve Office Python via Resolve-OfficePython.ps1
  exit /b 1
)

if "%1"=="" (start "" "%~f0" run & exit)
:run
title Add-In Publish: MORFunctions-dev.xlsm -> MORFunctions.xlam
cd /d "%~dp0"
echo ============================================================
echo  PUBLISH ADD-IN
echo  Source : MORFunctions-dev.xlsm  (dev workbook)
echo  Target : MORFunctions.xlam      (canonical add-in)
echo ============================================================
echo.
echo [1/3] Exporting current VBA state to vba-files\ ...
"%PY_EXEC%" scripts\vba_sync_watch.py export-all "MORFunctions-dev.xlsm"
if errorlevel 1 (
    echo [ERROR] VBA export failed. Aborting publish.
    pause & exit /b 1
)
echo.
echo [2/3] Saving dev workbook as staged add-in ...
"%PY_EXEC%" scripts\publish_addin.py stage "MORFunctions-dev.xlsm" "_stage_MORFunctions.xlam"
if errorlevel 1 (
    echo [ERROR] SaveAs xlAddIn failed. Aborting publish.
    pause & exit /b 1
)
echo.
echo [3/3] Copying staged add-in to canonical location ...
"%PY_EXEC%" scripts\publish_addin.py publish "_stage_MORFunctions.xlam" "MORFunctions.xlam"
if errorlevel 1 (
    echo [ERROR] Copy to canonical .xlam failed.
    pause & exit /b 1
)
echo.
echo ============================================================
echo  Publish complete. Reload the add-in in Excel if needed:
echo  File ^> Options ^> Add-Ins ^> Manage: Excel Add-Ins ^> Browse
echo ============================================================
pause
