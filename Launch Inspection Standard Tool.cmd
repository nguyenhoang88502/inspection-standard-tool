@echo off
setlocal
pushd "%~dp0"

echo Preparing Inspection Standard Tool data...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0system\inspection-standard-launcher.ps1" %*

if errorlevel 1 (
  echo.
  echo Inspection Standard Tool preparation failed.
  echo Review system\logs\Inspection_Standard_Tool_launcher.log.
  pause
)

popd
