module.exports = {
  testEnvironment: 'jsdom',
  roots: ['<rootDir>/src', '<rootDir>/tests'],
  testEnvironmentOptions: {
    customExportConditions: ['node', 'node-addons'],
    resources: 'usable'
  },
  setupFilesAfterEnv: ['<rootDir>/tests/setup.ts'],
  setupFiles: [
    '<rootDir>/tests/unit/frontend/setup.ts',
    '<rootDir>/tests/unit/backend/setup.ts',
    '<rootDir>/tests/integration/frontend/setup.ts',
    '<rootDir>/tests/integration/backend/setup.ts'
  ],
  transform: {
    '^.+\\.(ts|tsx)$': ['ts-jest', {
      tsconfig: {
        ...require('./tsconfig.json').compilerOptions,
        module: 'commonjs',
        paths: require('./tsconfig.json').compilerOptions.paths,
      },
      isolatedModules: true,
    }],
    '^.+\\.(js|jsx)$': 'babel-jest',
  },
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
  testMatch: [
    '<rootDir>/tests/unit/frontend/**/*.test.(ts|tsx|js|jsx)',
    '<rootDir>/tests/unit/backend/**/*.test.(ts|tsx|js|jsx)',
    '<rootDir>/tests/integration/frontend/**/*.test.(ts|tsx)',
    '<rootDir>/tests/integration/backend/**/*.test.(ts|tsx)',
  ],
  collectCoverageFrom: [
    'src/**/*.{ts,tsx}',
    '!src/**/*.d.ts',
    '!src/frontend/index.tsx',
    '!src/frontend/App.css',
    '!src/backend/**',
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
  coverageProvider: 'babel',
  maxWorkers: 1,
  clearMocks: true,
  restoreMocks: true,
  verbose: true,
  testTimeout: 10000,
  transformIgnorePatterns: [
    'node_modules/(?!(.*\\.mjs$))',
  ],
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/src/$1',
    '^@test/(.*)$': '<rootDir>/test/$1',
    '\\.(css|less|scss|sass)$': 'identity-obj-proxy',
  },
};