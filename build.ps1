Write-Host "Installing dependencies..."
Set-Location frontend
npm install

Write-Host "Building Angular application..."
npx @angular/cli build --configuration production

Write-Host "Build completed successfully!"
