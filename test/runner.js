#!/usr/bin/env node

/**
 * Test Runner Script for Excel Processor
 * Provides various test execution options
 */

const { execSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);
const command = args[0] || 'test';

const testConfig = {
  // Jest configuration file
  config: path.join(__dirname, 'jest.config.js'),
  
  // Test directories
  testDir: path.join(__dirname, 'test'),
  srcDir: path.join(__dirname, 'src'),
  
  // Coverage directory
  coverageDir: path.join(__dirname, 'coverage'),
};

const commands = {
  // Run all tests
  test: () => {
    console.log('🧪 Running all tests...');
    execSync(`jest --config ${testConfig.config}`, { stdio: 'inherit' });
  },
  
  // Run tests in watch mode
  'test:watch': () => {
    console.log('👀 Running tests in watch mode...');
    execSync(`jest --config ${testConfig.config} --watch`, { stdio: 'inherit' });
  },
  
  // Run tests with coverage
  'test:coverage': () => {
    console.log('📊 Running tests with coverage...');
    execSync(`jest --config ${testConfig.config} --coverage`, { stdio: 'inherit' });
  },
  
  // Run unit tests only
  'test:unit': () => {
    console.log('🔬 Running unit tests...');
    execSync(`jest --config ${testConfig.config} --testPathPattern="(utils|App).test"`, { stdio: 'inherit' });
  },
  
  // Run integration tests only
  'test:integration': () => {
    console.log('🔗 Running integration tests...');
    execSync(`jest --config ${testConfig.config} --testPathPattern="integration.test"`, { stdio: 'inherit' });
  },
  
  // Run server tests only
  'test:server': () => {
    console.log('🖥️ Running server tests...');
    execSync(`jest --config ${testConfig.config} --testPathPattern="server.test"`, { stdio: 'inherit' });
  },
  
  // Run tests in CI mode
  'test:ci': () => {
    console.log('🤖 Running tests in CI mode...');
    execSync(`jest --config ${testConfig.config} --ci --coverage --watchAll=false`, { stdio: 'inherit' });
  },
  
  // Clean test artifacts
  'test:clean': () => {
    console.log('🧹 Cleaning test artifacts...');
    try {
      execSync(`rmdir /s /q ${testConfig.coverageDir}`, { stdio: 'inherit' });
    } catch (error) {
      // Ignore errors if directory doesn't exist
    }
    console.log('✅ Test artifacts cleaned');
  },
  
  // Show help
  help: () => {
    console.log(`
🧪 Excel Processor Test Runner

Available commands:
  test              Run all tests
  test:watch        Run tests in watch mode
  test:coverage     Run tests with coverage report
  test:unit         Run unit tests only
  test:integration  Run integration tests only
  test:server       Run server tests only
  test:ci           Run tests in CI mode
  test:clean        Clean test artifacts
  help              Show this help message

Examples:
  node test/runner.js test
  node test/runner.js test:coverage
  node test/runner.js test:unit
    `);
  }
};

// Execute command
if (commands[command]) {
  try {
    commands[command]();
  } catch (error) {
    console.error(`❌ Error running command '${command}':`, error.message);
    process.exit(1);
  }
} else {
  console.error(`❌ Unknown command: ${command}`);
  console.log('Run "node test/runner.js help" for available commands');
  process.exit(1);
}
