// Simple test for LanguageContext without React Testing Library
import React from 'react';
import { LanguageProvider, useLanguage, formatMessage } from '@/contexts/LanguageContext';
import { Language } from '@/languages/index';

// Mock CSS imports
jest.mock('../../src/components/LanguageSelector.css', () => ({}));

// Simple test that doesn't use React Testing Library
describe('LanguageContext - Simple Test', () => {
  test('LanguageContext exports can be imported', () => {
    expect(LanguageProvider).toBeDefined();
    expect(useLanguage).toBeDefined();
    expect(formatMessage).toBeDefined();
  });

  test('LanguageProvider is a React component', () => {
    expect(typeof LanguageProvider).toBe('function');
  });

  test('useLanguage is a hook function', () => {
    expect(typeof useLanguage).toBe('function');
  });

  test('formatMessage is a utility function', () => {
    expect(typeof formatMessage).toBe('function');
  });

  test('formatMessage works correctly', () => {
    const message = 'Hello {name}, your age is {age}.';
    const placeholders = { name: 'John', age: 30 };
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Hello John, your age is 30.');
  });
});
