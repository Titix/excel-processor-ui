# Excel Processor - Test Suite

This directory contains comprehensive unit tests, integration tests, and test utilities for the Excel Processor application.

## 📁 Test Structure

```
test/
├── App.test.tsx              # Main React component tests
├── utils.test.ts             # Utility function tests
├── integration.test.tsx      # Integration tests
├── server.test.js            # Backend server tests
├── jest.config.js            # Jest configuration
├── setup.ts                  # Test setup file
├── globalSetup.ts            # Global test setup
├── globalTeardown.ts         # Global test teardown
├── runner.js                 # Test runner script
└── README.md                 # This file
```

## 🧪 Test Categories

### Unit Tests
- **App.test.tsx**: Tests for the main React component
  - Component rendering
  - User interactions (drag & drop, file upload)
  - File validation
  - Error handling
  - State management

- **utils.test.ts**: Tests for utility functions
  - File size formatting
  - File validation
  - Excel processing utilities
  - Download utilities
  - Message display utilities

### Integration Tests
- **integration.test.tsx**: End-to-end workflow tests
  - Complete file processing workflow
  - Multiple file uploads
  - Error handling integration
  - UI state consistency
  - Performance testing

### Server Tests
- **server.test.js**: Backend server tests
  - Health check endpoints
  - Static file serving
  - CORS configuration
  - Error handling

## 🚀 Running Tests

### Available Test Commands

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run tests with coverage report
npm run test:coverage

# Run unit tests only
npm run test:unit

# Run integration tests only
npm run test:integration

# Run server tests only
npm run test:server

# Run tests in CI mode
npm run test:ci

# Clean test artifacts
npm run test:clean
```

### Using the Test Runner

```bash
# Run specific test categories
node test/runner.js test:unit
node test/runner.js test:integration
node test/runner.js test:coverage

# Show help
node test/runner.js help
```

## 📊 Coverage Reports

Coverage reports are generated in the `coverage/` directory and include:

- **HTML Report**: `coverage/lcov-report/index.html`
- **LCOV Report**: `coverage/lcov.info`
- **Text Report**: Displayed in terminal

### Coverage Thresholds

- **Branches**: 80%
- **Functions**: 80%
- **Lines**: 80%
- **Statements**: 80%

## 🔧 Test Configuration

### Jest Configuration (`jest.config.js`)

- **Test Environment**: jsdom (for React components)
- **Setup Files**: `test/setup.ts`
- **Global Setup**: `test/globalSetup.ts`
- **Global Teardown**: `test/globalTeardown.ts`
- **Coverage**: Configured for src/ directory
- **Transform**: TypeScript and JavaScript files

### Test Setup (`setup.ts`)

- Mock implementations for browser APIs
- Console error/warning suppression
- Global test utilities
- Cleanup after each test

## 🎯 Test Utilities

### Global Test Utils

Available in all tests via `global.testUtils`:

```typescript
// Create mock files
const file = global.testUtils.createMockFile('test.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 1024);

// Create mock Excel files
const excelFile = global.testUtils.createMockExcelFile('test.xlsx', 1024);

// Create mock workbooks
const workbook = global.testUtils.createMockWorkbook('Sheet1', sheetData);

// Wait for async operations
await global.testUtils.waitFor(() => {
  expect(screen.getByText('Success')).toBeInTheDocument();
});
```

### Mock Implementations

- **XLSX Library**: Mocked for Excel processing
- **FileReader**: Mocked for file reading operations
- **URL APIs**: Mocked for download functionality
- **Browser APIs**: IntersectionObserver, ResizeObserver, etc.

## 📝 Writing Tests

### Test File Naming Convention

- Unit tests: `*.test.ts` or `*.test.tsx`
- Integration tests: `integration.test.tsx`
- Server tests: `server.test.js`

### Test Structure

```typescript
describe('Component Name', () => {
  beforeEach(() => {
    // Setup before each test
  });

  test('should do something', () => {
    // Test implementation
  });

  test('should handle errors', async () => {
    // Async test with error handling
  });
});
```

### Best Practices

1. **Descriptive Test Names**: Use clear, descriptive test names
2. **Arrange-Act-Assert**: Structure tests clearly
3. **Mock External Dependencies**: Mock file system, network calls, etc.
4. **Test Edge Cases**: Include error scenarios and edge cases
5. **Clean Up**: Ensure tests don't affect each other
6. **Async Testing**: Use `waitFor` for async operations

## 🐛 Debugging Tests

### Running Individual Tests

```bash
# Run specific test file
npm test -- App.test.tsx

# Run tests matching pattern
npm test -- --testNamePattern="should handle file upload"

# Run tests in specific directory
npm test -- test/
```

### Debug Mode

```bash
# Run tests with debug output
npm test -- --verbose

# Run tests with coverage and debug
npm run test:coverage -- --verbose
```

## 🔍 Test Monitoring

### Continuous Integration

The test suite is configured for CI environments with:

- **CI Mode**: `npm run test:ci`
- **Coverage Reports**: Generated automatically
- **Exit Codes**: Proper exit codes for CI systems
- **Timeout**: 30 seconds for integration tests

### Performance Testing

Integration tests include performance benchmarks:

- **Large File Processing**: Tests with 1000+ rows
- **Memory Usage**: Monitors memory consumption
- **Processing Time**: Ensures operations complete within 5 seconds

## 📚 Additional Resources

- [Jest Documentation](https://jestjs.io/docs/getting-started)
- [React Testing Library](https://testing-library.com/docs/react-testing-library/intro/)
- [Testing Best Practices](https://testing-library.com/docs/guiding-principles)
- [Jest Configuration](https://jestjs.io/docs/configuration)

## 🤝 Contributing

When adding new features:

1. Write tests first (TDD approach)
2. Ensure all tests pass
3. Maintain coverage thresholds
4. Update this README if needed
5. Run full test suite before committing

## 📞 Support

For test-related issues:

1. Check test logs for specific errors
2. Verify mock implementations
3. Ensure proper cleanup between tests
4. Check Jest configuration
5. Review test setup files
