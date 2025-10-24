// Simple test for LanguageSelector without React Testing Library
import React from 'react';
import { LanguageProvider } from '../../src/contexts/LanguageContext';

// Mock CSS imports
jest.mock('../../src/components/LanguageSelector.css', () => ({}));

// Simple test that doesn't use React Testing Library
describe('LanguageSelector - Simple Test', () => {
  test('LanguageSelector component can be imported', () => {
    // This test just verifies the component can be imported without crashing
    expect(() => {
      const LanguageSelector = require('../../src/components/LanguageSelector').default;
      expect(LanguageSelector).toBeDefined();
    }).not.toThrow();
  });

  test('LanguageProvider can be imported', () => {
    expect(LanguageProvider).toBeDefined();
  });

  test('LanguageProvider is a React component', () => {
    expect(typeof LanguageProvider).toBe('function');
  });
});
