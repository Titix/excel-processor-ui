const { execSync } = require('child_process');
const path = require('path');

console.log('Installing dependencies...');
process.chdir(path.join(__dirname, 'frontend'));
execSync('npm install', { stdio: 'inherit' });

console.log('Building Angular application...');
execSync('npx @angular/cli build --configuration production', { stdio: 'inherit' });

console.log('Build completed successfully!');
