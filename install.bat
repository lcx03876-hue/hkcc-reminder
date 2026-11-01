@echo off
setlocal
cd /d "%~dp0"

if not "%~1"=="" set "HKCC_EXPECTED_USER=%~1"
if not defined HKCC_EXPECTED_USER set "HKCC_EXPECTED_USER=%USERNAME%"

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Requesting administrator privileges...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList '%USERNAME%' -Verb RunAs"
  exit /b
)

echo Starting installer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -ExpectedUser "%HKCC_EXPECTED_USER%"
if errorlevel 1 pause
endlocal
