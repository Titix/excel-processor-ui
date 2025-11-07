// Simple test for LanguageSelector without React Testing Library
import React from 'react';
import { LanguageProvider } from '@/contexts/LanguageContext';
import LanguageSelector from '@/components/LanguageSelector';

// Mock CSS imports
jest.mock('../../src/components/LanguageSelector.css', () => ({}));

// Simple test that doesn't use React Testing Library
describe('LanguageSelector - Simple Test', () => {
  test('LanguageSelector component can be imported', () => {
    // This test just verifies the component can be imported without crashing
    expect(LanguageSelector).toBeDefined();
  });

  test('LanguageProvider can be imported', () => {
    expect(LanguageProvider).toBeDefined();
  });

  test('LanguageProvider is a React component', () => {
    expect(typeof LanguageProvider).toBe('function');
  });
});
