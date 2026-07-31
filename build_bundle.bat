@echo off
setlocal EnableExtensions

cd /d "%~dp0"

echo ============================================================
echo KEYZAR FENIX BOT - BUILD
echo ============================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Virtual environment not found:
    echo         %CD%\.venv
    echo.
    echo Create it first with:
    echo     py -3.11 -m venv .venv
    echo.
    pause
    exit /b 1
)

set "PYTHON=%CD%\.venv\Scripts\python.exe"

echo [1/7] Installing/updating build dependencies...
"%PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto :failed

"%PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 goto :failed

"%PYTHON%" -m pip install --upgrade pyinstaller
if errorlevel 1 goto :failed

echo.
echo [2/7] Cleaning old build output...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo [3/7] Building launcher EXE...
"%PYTHON%" -m PyInstaller --clean --noconfirm KeyzarFenixBot.spec
if errorlevel 1 goto :failed

echo.
echo [4/7] Copying updateable bot files...
xcopy "bot" "dist\KeyzarFenixBot\bot" /E /I /H /Y >nul
if errorlevel 1 goto :failed

echo.
echo [5/7] Copying updater and local configuration...
copy /Y "updater.py" "dist\KeyzarFenixBot\updater.py" >nul

if exist ".env" (
    copy /Y ".env" "dist\KeyzarFenixBot\.env" >nul
) else (
    echo [WARNING] .env was not found. The installer will not contain one.
)

echo.
echo [6/7] Creating required local folders...
if not exist "dist\KeyzarFenixBot\data\logs" mkdir "dist\KeyzarFenixBot\data\logs"
if not exist "dist\KeyzarFenixBot\data\pending_reports" mkdir "dist\KeyzarFenixBot\data\pending_reports"
if not exist "dist\KeyzarFenixBot\data\sent_reports" mkdir "dist\KeyzarFenixBot\data\sent_reports"
if not exist "dist\KeyzarFenixBot\data\screenshots" mkdir "dist\KeyzarFenixBot\data\screenshots"
if not exist "dist\KeyzarFenixBot\playwright_profile" mkdir "dist\KeyzarFenixBot\playwright_profile"

echo.
echo [7/7] Verifying output...
if not exist "dist\KeyzarFenixBot\KeyzarFenixBot.exe" goto :failed
if not exist "dist\KeyzarFenixBot\bot\app_runtime.py" goto :failed
if not exist "dist\KeyzarFenixBot\updater.py" goto :failed

echo.
echo ============================================================
echo BUILD COMPLETED SUCCESSFULLY
echo ============================================================
echo Output:
echo   %CD%\dist\KeyzarFenixBot
echo.
echo Test before creating the installer:
echo   dist\KeyzarFenixBot\KeyzarFenixBot.exe
echo.
pause
exit /b 0

:failed
echo.
echo ============================================================
echo BUILD FAILED
echo ============================================================
echo Review the error shown above.
echo.
pause
exit /b 1
