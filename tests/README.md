# Testing Strategy Documentation

## Overview
This project implements a comprehensive testing strategy with three levels of testing:

1. **Unit Tests** - Test individual components and functions in isolation
2. **Integration Tests** - Test component interactions and workflows
3. **End-to-End Tests** - Test complete user workflows in real browser environment

## Test Structure

```
tests/
├── e2e/                    # End-to-End tests (Playwright)
│   ├── excel-processing.spec.ts
│   └── basic-functionality.spec.ts
├── integration/           # Integration tests (Jest + React Testing Library)
│   ├── app-integration.test.tsx
│   └── setup.ts
└── unit/                  # Unit tests (Jest + React Testing Library)
    ├── setup.ts
    ├── utils.test.ts
    ├── App.simple.test.tsx
    ├── simple.test.tsx
    ├── LanguageSelector.test.tsx
    └── LanguageContext.test.tsx
```

## Running Tests

### Unit Tests
```bash
npm run test              # Run all unit tests
npm run test:watch        # Run tests in watch mode
npm run test:coverage     # Run tests with coverage report
npm run test:unit         # Run only unit tests
```

### Integration Tests
```bash
npm run test:integration  # Run integration tests
```

### End-to-End Tests
```bash
npm run test:e2e          # Run E2E tests (headless)
npm run test:e2e:headed   # Run E2E tests (with browser UI)
npm run test:e2e:ui       # Run E2E tests with Playwright UI
```

### All Tests
```bash
npm run test:all          # Run unit + integration + E2E tests
```

## Test Coverage

### Current Coverage Status
- **Unit Tests**: 25.75% overall coverage
  - `LanguageSelector.tsx`: 100% coverage
  - `LanguageContext.tsx`: 100% coverage
  - `utils.ts`: 100% coverage
  - `App.tsx`: 16.66% coverage (low due to complex DOM interactions)

### Coverage Goals
- **Unit Tests**: Target 80%+ for pure functions and isolated components
- **Integration Tests**: Target 60-70% for component interactions
- **E2E Tests**: Provide confidence in complete workflows

## Test Types Explained

### Unit Tests
- **Purpose**: Test individual functions and components in isolation
- **Tools**: Jest + React Testing Library
- **Scope**: Pure functions, utility functions, individual components
- **Examples**: LanguageSelector, LanguageContext, utility functions

### Integration Tests
- **Purpose**: Test component interactions and workflows
- **Tools**: Jest + React Testing Library + MSW
- **Scope**: Component interactions, state management, user workflows
- **Examples**: Complete App workflow, file processing, language switching

### End-to-End Tests
- **Purpose**: Test complete user workflows in real browser
- **Tools**: Playwright
- **Scope**: Complete user journeys, browser compatibility, real file operations
- **Examples**: Folder selection → file processing → download workflow

## Mocking Strategy

### File System Access API
- **Unit Tests**: Mocked at component level
- **Integration Tests**: Mocked globally with MSW
- **E2E Tests**: Mocked with Playwright's `addInitScript`

### XLSX Library
- **Unit Tests**: Mocked per test
- **Integration Tests**: Mocked globally
- **E2E Tests**: Mocked with Playwright's `addInitScript`

### Browser APIs
- **URL.createObjectURL**: Mocked for download testing
- **document.createElement**: Mocked for anchor element creation
- **Blob**: Mocked for file creation

## Test Data

### Sample Excel Data
```typescript
const mockExcelData = {
  SheetNames: ['Sheet1'],
  Sheets: {
    'Sheet1': {
      '!ref': 'A1:C3',
      'A1': { v: 'Name', t: 's' },
      'B1': { v: 'Age', t: 's' },
      'C1': { v: 'City', t: 's' },
      'A2': { v: 'John', t: 's' },
      'B2': { v: 25, t: 'n' },
      'C2': { v: 'New York', t: 's' },
      'A3': { v: 'Jane', t: 's' },
      'B3': { v: 30, t: 'n' },
      'C3': { v: 'London', t: 's' }
    }
  }
};
```

### Test Scenarios

#### Duplicate Detection
- **Unit Tests**: Test duplicate detection logic in isolation
- **Integration Tests**: Test duplicate removal in complete workflow
- **E2E Tests**: Test duplicate removal with real file processing

#### File Processing
- **Unit Tests**: Test individual file processing functions
- **Integration Tests**: Test multi-file processing workflow
- **E2E Tests**: Test complete file processing with download

#### Error Handling
- **Unit Tests**: Test error handling in individual functions
- **Integration Tests**: Test error handling in component workflows
- **E2E Tests**: Test error handling in complete user journeys

## Best Practices

### Test Organization
- Group related tests in `describe` blocks
- Use descriptive test names
- Follow AAA pattern (Arrange, Act, Assert)

### Mocking
- Mock external dependencies
- Use realistic test data
- Reset mocks between tests

### Assertions
- Use specific assertions (`toBeVisible`, `toHaveTextContent`)
- Test both positive and negative cases
- Test edge cases and error conditions

### Performance
- Use `waitFor` for async operations
- Avoid unnecessary waits
- Use `maxWorkers: 1` for Jest to prevent crashes

## Troubleshooting

### Common Issues

#### Jest Crashes
- **Cause**: Multiple Jest versions, Node.js native crashes
- **Solution**: Use Jest 27.5.1, set `maxWorkers: 1`

#### DOM Container Issues
- **Cause**: Missing root element for React rendering
- **Solution**: Create DOM container in `beforeEach` hook

#### Coverage Issues
- **Cause**: Complex DOM interactions, File System Access API
- **Solution**: Use integration and E2E tests for complex workflows

#### Playwright Issues
- **Cause**: Server not running, port conflicts
- **Solution**: Ensure server runs on port 3000, use `webServer` config

### Debug Commands
```bash
# Debug Jest tests
npm run test -- --verbose --no-coverage

# Debug Playwright tests
npx playwright test --debug

# Check test coverage
npm run test:coverage
```

## Future Improvements

### Test Coverage
- Increase App.tsx unit test coverage
- Add more edge case tests
- Improve error scenario coverage

### Performance
- Optimize test execution time
- Reduce test flakiness
- Improve test reliability

### Maintenance
- Regular test updates
- Dependency updates
- Test documentation updates
