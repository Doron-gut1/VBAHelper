@echo off
setlocal

REM בודק אם קובץ ה-RAR קיים
if not exist "ExternalLibraries.rar" (
    echo [ERROR] ExternalLibraries.rar not found in the current directory.
    pause
    exit /b 1
)

REM בודק אם UNRAR קיים
where unrar >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 'unrar' is not installed or not in PATH.
    echo Please install WinRAR or add 'unrar' to system PATH.
    pause
    exit /b 1
)

REM מחלץ את הקובץ
echo Extracting ExternalLibraries.rar...
unrar x -y ExternalLibraries.rar

echo Extraction complete.
pause
