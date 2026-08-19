@echo off
title Auto Purchase Order Server
chcp 949 >nul
cls

echo ========================================================
echo   [Security] Auto Purchase Order System (v1.0)
echo ========================================================
echo.

:AUTH_LOOP
set "INPUT_PW="
set /p "INPUT_PW=Password: "

if "%INPUT_PW%"=="0708" goto :AUTH_SUCCESS
if "%INPUT_PW%"==" 0708" goto :AUTH_SUCCESS
if "%INPUT_PW%"=="0708 " goto :AUTH_SUCCESS

echo.
echo [Error] Incorrect Password. Please try again.
echo.
goto :AUTH_LOOP

:AUTH_SUCCESS
echo.
echo [Success] Password verified! Starting server...
echo.

set "UV_BIN="

where uv >nul 2>&1
if not errorlevel 1 set "UV_BIN=uv"

if "%UV_BIN%"=="" if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.cargo\bin\uv.exe"
if "%UV_BIN%"=="" if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.local\bin\uv.exe"
if "%UV_BIN%"=="" if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_BIN=%LOCALAPPDATA%\Programs\uv\uv.exe"
if "%UV_BIN%"=="" if exist "%APPDATA%\uv\uv.exe" set "UV_BIN=%APPDATA%\uv\uv.exe"

if not "%UV_BIN%"=="" goto :START_UV

echo [Notice] Installing uv package manager...
powershell -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm https://astral.sh/uv/install.ps1 | iex"

if exist "%USERPROFILE%\.cargo\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.cargo\bin\uv.exe"
if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_BIN=%USERPROFILE%\.local\bin\uv.exe"
if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_BIN=%LOCALAPPDATA%\Programs\uv\uv.exe"
if exist "%APPDATA%\uv\uv.exe" set "UV_BIN=%APPDATA%\uv\uv.exe"

if not "%UV_BIN%"=="" goto :START_UV

goto :START_PYTHON

:START_UV
echo [1/2] Syncing packages...
"%UV_BIN%" sync
if %errorlevel% neq 0 (
    echo [Warning] uv sync failed. Falling back to python...
    goto :START_PYTHON
)

echo.
echo [2/2] Starting Web Server...
echo ========================================================
echo   Web URL: http://127.0.0.1:8000
echo ========================================================
echo.
start http://127.0.0.1:8000
"%UV_BIN%" run python app.py
goto :DONE

:START_PYTHON
echo.
echo [Notice] Running with system Python...
where python >nul 2>&1
if errorlevel 1 goto :NO_PYTHON

echo [1/2] Installing requirements...
python -m pip install -r requirements.txt --quiet

echo.
echo [2/2] Starting Web Server...
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
echo [Error] Python or uv is not installed.
echo ========================================================
echo.

:DONE
echo.
echo Server stopped.
pause
