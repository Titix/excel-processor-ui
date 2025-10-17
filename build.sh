#!/bin/bash
set -e

echo "Installing dependencies..."
cd frontend
npm install

echo "Building Angular application..."
npx @angular/cli build --configuration production

echo "Build completed successfully!"
