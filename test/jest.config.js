module.exports = {
  testEnvironment: 'jsdom',
  setupFilesAfterEnv: ['<rootDir>/test/setup.ts'],
  transform: {
    '^.+\\.(ts|tsx)$': 'ts-jest',
    '^.+\\.(js|jsx)$': 'babel-jest',
  },
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
  testMatch: [
    '<rootDir>/test/**/*.test.(ts|tsx|js|jsx)',
  ],
  collectCoverageFrom: [
    'src/frontend/**/*.{ts,tsx}',
    '!src/frontend/**/*.d.ts',
    '!src/frontend/index.tsx',
  ],
  coverageThreshold: {
    global: {
      branches: 100,
      functions: 100,
      lines: 100,
      statements: 100,
    },
  },
  coverageReporters: ['text', 'lcov', 'html'],
  coverageDirectory: 'coverage',
  clearMocks: true,
  restoreMocks: true,
  verbose: true,
  testTimeout: 10000,
  transformIgnorePatterns: [
    'node_modules/(?!(.*\\.mjs$))',
  ],
};