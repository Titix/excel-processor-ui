import React from 'react';
import { render, screen } from '@testing-library/react';
import '@testing-library/jest-dom';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

import App from '../../src/frontend/App';
import { LanguageProvider } from '../../src/contexts/LanguageContext';

// Mock XLSX library
const mockXLSX = {
  read: jest.fn().mockReturnValue({
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
  }),
  write: jest.fn().mockReturnValue(new Uint8Array([1, 2, 3, 4, 5])),
  utils: {
    sheet_to_json: jest.fn().mockReturnValue([
      ['Name', 'Age', 'City'],
      ['John', 25, 'New York'],
      ['Jane', 30, 'London']
    ]),
    decode_range: jest.fn().mockReturnValue({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } }),
    encode_cell: jest.fn().mockImplementation(({ r, c }) => {
      const col = String.fromCharCode(65 + c);
      return `${col}${r + 1}`;
    })
  }
};

// Mock window.XLSX
Object.defineProperty(window, 'XLSX', {
  value: mockXLSX,
  writable: true,
});

// Mock File System Access API
const mockDirectoryHandle = {
  name: 'test-folder',
  entries: jest.fn(),
};

const mockFileHandle = {
  kind: 'file',
  getFile: jest.fn(),
};

const mockFile = {
  name: 'test.xlsx',
  size: 1024,
};

// Mock showDirectoryPicker
Object.defineProperty(window, 'showDirectoryPicker', {
  value: jest.fn(),
  writable: true,
});

describe('Simple App Test', () => {
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

  test('renders app without crashing', () => {
    const container = document.getElementById('root');
    expect(container).toBeTruthy();
    
    render(
      <LanguageProvider>
        <App />
      </LanguageProvider>,
      { container: container! }
    );
    
    expect(screen.getByText('Excel File Processor')).toBeInTheDocument();
  });
});

