import React from 'react';
import { render, screen, fireEvent, act } from '@testing-library/react';
import '@testing-library/jest-dom';

import { LanguageProvider, useLanguage, formatMessage } from '@/contexts/LanguageContext';
import { Language } from '@/languages/index';

// Mock CSS imports
jest.mock('../src/components/LanguageSelector.css', () => ({}));

// Test component that uses the language context
const TestComponent: React.FC = () => {
  const { language, setLanguage, t } = useLanguage();

  return (
    <div>
      <div data-testid="current-language">{language}</div>
      <div data-testid="current-text">{t.language}</div>
      <button onClick={() => setLanguage('hu')}>Set Hungarian</button>
      <button onClick={() => setLanguage('en')}>Set English</button>
    </div>
  );
};

// Test component that should throw error when used outside provider
const TestComponentWithoutProvider: React.FC = () => {
  useLanguage();
  return <div>Should not render</div>;
};

describe('LanguageContext', () => {
  beforeEach(() => {
    // Clear any existing root element
    const existingRoot = document.getElementById('root');
    if (existingRoot) {
      existingRoot.remove();
    }
    
    // Create a div element for React to render into
    const div = document.createElement('div');
    div.id = 'root';
    document.body.appendChild(div);
    
    // Clear localStorage
    localStorage.clear();
  });

  afterEach(() => {
    const root = document.getElementById('root');
    if (root) {
      root.remove();
    }
    
    // Clear localStorage
    localStorage.clear();
  });

  const renderWithProvider = () => {
    const container = document.getElementById('root');
    expect(container).toBeTruthy();
    
    return render(
      <LanguageProvider>
        <TestComponent />
      </LanguageProvider>,
      { container: container! }
    );
  };

  test('provides default language context', () => {
    renderWithProvider();
    
    expect(screen.getByTestId('current-language')).toHaveTextContent('en');
    expect(screen.getByTestId('current-text')).toHaveTextContent('Language');
  });

  test('changes language when setLanguage is called', () => {
    renderWithProvider();
    
    expect(screen.getByTestId('current-language')).toHaveTextContent('en');
    
    const hungarianButton = screen.getByText('Set Hungarian');
    fireEvent.click(hungarianButton);
    
    expect(screen.getByTestId('current-language')).toHaveTextContent('hu');
    expect(screen.getByTestId('current-text')).toHaveTextContent('Nyelv');
  });

  test('saves language to localStorage', () => {
    renderWithProvider();
    
    const hungarianButton = screen.getByText('Set Hungarian');
    fireEvent.click(hungarianButton);
    
    expect(localStorage.getItem('excel-processor-language')).toBe('hu');
  });

  test('loads language from localStorage on mount', () => {
    // Set language in localStorage before rendering
    localStorage.setItem('excel-processor-language', 'hu');
    
    renderWithProvider();
    
    expect(screen.getByTestId('current-language')).toHaveTextContent('hu');
    expect(screen.getByTestId('current-text')).toHaveTextContent('Nyelv');
  });

  test('ignores invalid language from localStorage', () => {
    // Set invalid language in localStorage
    localStorage.setItem('excel-processor-language', 'invalid');
    
    renderWithProvider();
    
    // Should still use default language
    expect(screen.getByTestId('current-language')).toHaveTextContent('en');
    expect(screen.getByTestId('current-text')).toHaveTextContent('Language');
  });

  test('throws error when useLanguage is used outside provider', () => {
    const container = document.getElementById('root');
    expect(container).toBeTruthy();
    
    // Suppress console.error for this test
    const originalError = console.error;
    console.error = jest.fn();
    
    expect(() => {
      render(<TestComponentWithoutProvider />, { container: container! });
    }).toThrow('useLanguage must be used within a LanguageProvider');
    
    // Restore console.error
    console.error = originalError;
  });

  test('formatMessage replaces placeholders correctly', () => {
    const message = 'Hello {name}, you have {count} messages';
    const placeholders = { name: 'John', count: 5 };
    
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Hello John, you have 5 messages');
  });

  test('formatMessage handles missing placeholders', () => {
    const message = 'Hello {name}, you have {count} messages';
    const placeholders = { name: 'John' }; // Missing count
    
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Hello John, you have {count} messages');
  });

  test('formatMessage handles empty placeholders', () => {
    const message = 'Hello {name}';
    const placeholders = {};
    
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Hello {name}');
  });

  test('formatMessage handles numeric placeholders', () => {
    const message = 'Count: {count}';
    const placeholders = { count: 42 };
    
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Count: 42');
  });

  test('formatMessage handles multiple same placeholders', () => {
    const message = '{greeting} {name}, {greeting} again!';
    const placeholders = { greeting: 'Hello', name: 'John' };
    
    const result = formatMessage(message, placeholders);
    expect(result).toBe('Hello John, {greeting} again!');
  });
});
