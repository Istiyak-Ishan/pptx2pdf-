@echo off
echo ===================================================
echo Building PPTX2PDF Windows Desktop Executable
echo ===================================================

IF EXIST .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

echo Cleaning previous build artifacts...
IF EXIST build RMDIR /S /Q build
IF EXIST dist RMDIR /S /Q dist

echo Running PyInstaller...
pyinstaller --clean pptx2pdf.spec

IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller build failed!
    exit /b %ERRORLEVEL%
)

echo.
echo ===================================================
echo SUCCESS! Executable built at: dist\PPTX2PDF.exe
echo ===================================================
pause
