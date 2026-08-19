@echo off
title Auto Purchase Order Server

echo ========================================================
echo   Auto Purchase Order Automation System
echo ========================================================
echo.

set "UV_BIN="

where uv >nul 2>&1
if not errorlevel 1 set "UV_BIN=uv"

if "%UV_BIN%"=="" if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.cargo\bin\uv.exe"
if "%UV_BIN%"=="" if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.local\bin\uv.exe"
if "%UV_BIN%"=="" if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_BIN=%LOCALAPPDATA%\Programs\uv\uv.exe"
if "%UV_BIN%"=="" if exist "%APPDATA%\uv\uv.exe" set "UV_BIN=%APPDATA%\uv\uv.exe"

if not "%UV_BIN%"=="" goto :START_UV

echo [Notice] Downloading and installing uv...
powershell -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm https://astral.sh/uv/install.ps1 | iex"

if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.cargo\bin\uv.exe"
if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.local\bin\uv.exe"
if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_BIN=%LOCALAPPDATA%\Programs\uv\uv.exe"
if exist "%APPDATA%\uv\uv.exe" set "UV_BIN=%APPDATA%\uv\uv.exe"

if not "%UV_BIN%"=="" goto :START_UV

goto :START_PYTHON

:START_UV
echo [1/3] Synchronizing dependencies with uv...
"%UV_BIN%" sync
if %errorlevel% neq 0 (
    echo [Warning] uv sync failed. Falling back to Python...
    goto :START_PYTHON
)

echo.
echo [2/3] Checking template files...
if not exist "templates\price_master.xlsx" (
    "%UV_BIN%" run python init_templates.py
)

echo.
echo [3/3] Starting web server...
echo ========================================================
echo   Web URL: http://127.0.0.1:8000
echo   (Close this window to stop the server)
echo ========================================================
echo.
start http://127.0.0.1:8000
"%UV_BIN%" run python app.py
goto :DONE

:START_PYTHON
echo.
echo [Notice] Attempting execution with standard Python...
where python >nul 2>&1
if errorlevel 1 goto :NO_PYTHON

echo [1/3] Checking packages...
python -m pip install -r requirements.txt --quiet

echo [2/3] Checking templates...
if not exist "templates\price_master.xlsx" (
    python init_templates.py
)

echo [3/3] Starting web server...
echo ========================================================
echo   Web URL: http://127.0.0.1:8000
echo ========================================================
echo.
start http://127.0.0.1:8000
python app.py
goto :DONE

:NO_PYTHON
echo.
echo ========================================================
echo [Error] Neither uv nor Python was found.
echo Please install uv from https://docs.astral.sh/uv/
echo or Python from https://python.org
echo ========================================================
echo.

:DONE
echo.
echo Program ended.
pause
