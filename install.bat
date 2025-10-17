@echo off
echo ========================================
echo Excel Processor UI - Installation Script
echo ========================================
echo.

echo Checking if Node.js is installed...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed!
    echo Please download and install Node.js from https://nodejs.org/
    echo This will also install npm automatically.
    pause
    exit /b 1
)

echo Node.js is installed!
node --version
echo.

echo Installing backend dependencies...
npm install
if %errorlevel% neq 0 (
    echo ERROR: Failed to install backend dependencies!
    pause
    exit /b 1
)

echo.
echo Installing frontend dependencies...
cd frontend
npm install
if %errorlevel% neq 0 (
    echo ERROR: Failed to install frontend dependencies!
    pause
    exit /b 1
)

cd ..
echo.
echo ========================================
echo Installation completed successfully!
echo ========================================
echo.
echo To start the application:
echo.
echo Development mode:
echo   1. Start frontend: cd frontend ^&^& npm start
echo   2. Start backend: npm run dev
echo.
echo Production mode:
echo   1. Build frontend: npm run build
echo   2. Start server: npm start
echo.
echo Frontend will be available at: http://localhost:4200
echo Backend will be available at: http://localhost:3000
echo.
pause
