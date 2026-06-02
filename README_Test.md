# Test Guide — excel-processor-ui

This document shows the command to run followed by a short explanation.

## Change to project directory
```bash
cd "C:\Users\Kecsap\excel-processor-ui"
```
Run all commands from the project root.

## Start the React development server (http://localhost:3000). Use for manual UI testing.
```bash
npm start
```

## Run the entire test suite (fast default)
```bash
npm test
```

## Run backend unit tests only (serial)
```bash
npm run test:unit:backend -- --runInBand
```
Runs backend-focused unit tests. Use `--runInBand` for CI or when isolation is needed.

## Run frontend unit tests only (serial)
```bash
npm run test:unit:frontend -- --runInBand
```
Runs fast frontend unit/smoke tests.

## Run a single frontend test file (verbose)
npm run test:unit:frontend -- --runInBand
```bash
npx jest tests/unit/frontend/<file>.test.tsx -i --runInBand --verbose
```
Use when developing or debugging a specific test file.

## Run integration tests only (serial)
npm run test:integration -- --runInBand
```bash
```
Integration tests are heavier and cover end-to-end flows and environment mocks.

## Alternate: explicit integration invocation
```bash
npx jest --config jest.config.js --testPathPattern="tests/integration" --runInBand
```

## Run full suite with coverage (outputs to coverage/)
```bash
npm run test:coverage
```
or
```bash
npx jest --config jest.config.js --coverage --coverageDirectory=coverage --runInBand
```
Generates lcov/html coverage report under `coverage/lcov-report/`.

## Open the HTML coverage report (Windows / PowerShell)
```powershell
start .\coverage\lcov-report\index.html
```
Opens the coverage report in the default browser.

---
PowerShell tips
- Some PowerShell builds don't support `&&`. Use `;` or run commands on separate lines.
  Example:
  cd "C:\Users\Kecsap\excel-processor-ui"; npm run test:integration -- --runInBand

Project notes
- Integration tests (heavier flows & mocks): tests/integration/frontend/
- Unit tests (fast): tests/unit/frontend/
- I consolidated the coverage-focused tests into integration to keep unit tests fast.

Troubleshooting
- If coverage thresholds fail, inspect coverage/lcov-report/index.html and add targeted tests for uncovered lines.
- Ensure File API mocks (showDirectoryPicker, File handles, FileReader/XLSX) are present for integration tests.
- For act() warnings, wrap state-updating test actions in React's act or await async updates.

CI recommendation
- Run unit tests for every PR.
- Run integration/coverage tests in a separate CI job (scheduled or pre-merge to main).

If you want IntelliJ run configurations, an npm script to open the coverage report, or further test-splitting, I can add them. 

