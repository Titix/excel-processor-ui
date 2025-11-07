import React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import '@testing-library/jest-dom';

import LanguageSelector from '@/components/LanguageSelector';
import { LanguageProvider } from '@/contexts/LanguageContext';

// Mock CSS imports
jest.mock('../../src/components/LanguageSelector.css', () => ({}));

describe('LanguageSelector', () => {
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
  });

  afterEach(() => {
    const root = document.getElementById('root');
    if (root) {
      root.remove();
    }
  });

  const renderLanguageSelector = () => {
    const container = document.getElementById('root');
    expect(container).toBeTruthy();
    
    return render(
      <LanguageProvider>
        <LanguageSelector />
      </LanguageProvider>,
      { container: container! }
    );
  };

  test('renders language selector', () => {
    renderLanguageSelector();
    
    expect(screen.getByText('Language:')).toBeInTheDocument();
    expect(screen.getByRole('combobox')).toBeInTheDocument();
  });

  test('shows language options', () => {
    renderLanguageSelector();
    
    const select = screen.getByRole('combobox');
    expect(select).toBeInTheDocument();
    
    // Check that both language options are present
    expect(screen.getByText('🇺🇸 English')).toBeInTheDocument();
    expect(screen.getByText('🇭🇺 Magyar')).toBeInTheDocument();
  });

  test('changes language when option is selected', () => {
    renderLanguageSelector();
    
    const select = screen.getByRole('combobox');
    expect(select).toHaveValue('en'); // Default language
    
    // Change to Hungarian
    fireEvent.change(select, { target: { value: 'hu' } });
    expect(select).toHaveValue('hu');
    
    // Change back to English
    fireEvent.change(select, { target: { value: 'en' } });
    expect(select).toHaveValue('en');
  });

  test('has correct HTML attributes', () => {
    renderLanguageSelector();
    
    const select = screen.getByRole('combobox');
    expect(select).toHaveAttribute('id', 'language-select');
    expect(select).toHaveClass('language-select');
    
    const label = screen.getByText('Language:');
    expect(label).toHaveAttribute('for', 'language-select');
    expect(label).toHaveClass('language-label');
  });

  test('renders with correct structure', () => {
    renderLanguageSelector();
    
    const container = screen.getByRole('combobox').closest('.language-selector');
    expect(container).toBeInTheDocument();
    
    const label = container?.querySelector('.language-label');
    expect(label).toBeInTheDocument();
    
    const select = container?.querySelector('.language-select');
    expect(select).toBeInTheDocument();
  });
});
