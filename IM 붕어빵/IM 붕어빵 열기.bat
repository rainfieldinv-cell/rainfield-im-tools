@echo off
cd /d "%~dp0"
title IM converter

rem Double-click to start the IM converter and open it in the browser.
rem If it is already running, this just opens the browser.

set PORT=8501
set URL=http://localhost:%PORT%

netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 goto open

echo.
echo   Starting the IM converter... (10-20 sec)
echo.
start "IM converter server" /min python -m streamlit run app.py --server.port %PORT% --server.headless true

for /l %%i in (1,1,90) do (
    netstat -ano | findstr ":%PORT% " | findstr "LISTENING" >nul 2>&1
    if not errorlevel 1 goto open
    ping -n 2 127.0.0.1 >nul
)

echo.
echo   Failed to start. Is Python installed?
echo.
pause
exit /b 1

:open
start "" "%URL%"
exit /b 0
