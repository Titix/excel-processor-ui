Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Excel Processor UI - Installation Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Checking if Node.js is installed..." -ForegroundColor Yellow
try {
    $nodeVersion = node --version
    Write-Host "Node.js is installed! Version: $nodeVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Node.js is not installed!" -ForegroundColor Red
    Write-Host "Please download and install Node.js from https://nodejs.org/" -ForegroundColor Red
    Write-Host "This will also install npm automatically." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Installing backend dependencies..." -ForegroundColor Yellow
npm install
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install backend dependencies!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Installing frontend dependencies..." -ForegroundColor Yellow
Set-Location frontend
npm install
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install frontend dependencies!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Set-Location ..
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Installation completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "To start the application:" -ForegroundColor Cyan
Write-Host ""
Write-Host "Development mode:" -ForegroundColor Yellow
Write-Host "  1. Start frontend: cd frontend; npm start" -ForegroundColor White
Write-Host "  2. Start backend: npm run dev" -ForegroundColor White
Write-Host ""
Write-Host "Production mode:" -ForegroundColor Yellow
Write-Host "  1. Build frontend: npm run build" -ForegroundColor White
Write-Host "  2. Start server: npm start" -ForegroundColor White
Write-Host ""
Write-Host "Frontend will be available at: http://localhost:4200" -ForegroundColor Green
Write-Host "Backend will be available at: http://localhost:3000" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
