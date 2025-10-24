import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
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
  getFileHandle: jest.fn(),
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

// Mock URL and Blob APIs
Object.defineProperty(window, 'URL', {
  value: {
    createObjectURL: jest.fn().mockReturnValue('mock-url'),
    revokeObjectURL: jest.fn()
  },
  writable: true
});

// Mock document.createElement and appendChild
const mockAnchor = {
  href: '',
  download: '',
  click: jest.fn()
};

Object.defineProperty(document, 'createElement', {
  value: jest.fn().mockReturnValue(mockAnchor),
  writable: true
});

Object.defineProperty(document.body, 'appendChild', {
  value: jest.fn(),
  writable: true
});

Object.defineProperty(document.body, 'removeChild', {
  value: jest.fn(),
  writable: true
});

describe('App Integration Tests', () => {
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
    
    // Reset all mocks
    jest.clearAllMocks();
    
    // Reset XLSX mocks
    mockXLSX.read.mockReturnValue({
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
    });
    
    mockXLSX.utils.sheet_to_json.mockReturnValue([
      ['Name', 'Age', 'City'],
      ['John', 25, 'New York'],
      ['Jane', 30, 'London']
    ]);
    
    // Reset File System Access API mocks
    mockDirectoryHandle.entries.mockReturnValue([
      ['test1.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(new File(['test content 1'], 'test1.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
      }],
      ['test2.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(new File(['test content 2'], 'test2.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
      }]
    ]);
    
    (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);
  });

  afterEach(() => {
    const root = document.getElementById('root');
    if (root) {
      root.remove();
    }
  });

  const renderApp = () => {
    const container = document.getElementById('root');
    expect(container).toBeTruthy();
    
    return render(
      <LanguageProvider>
        <App />
      </LanguageProvider>,
      { container: container! }
    );
  };

  test('complete workflow integration', async () => {
    renderApp();
    
    // Test initial state
    expect(screen.getByText('Excel File Processor')).toBeInTheDocument();
    expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    
    // Test folder selection
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
      expect(screen.getByText('2 Excel files found')).toBeInTheDocument();
    });
    
    // Test file list display
    expect(screen.getByText('Excel Files Found')).toBeInTheDocument();
    expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
    expect(screen.getByText('test2.xlsx')).toBeInTheDocument();
    
    // Test file selection
    const checkboxes = screen.getAllByRole('checkbox');
    expect(checkboxes[0]).not.toBeChecked();
    fireEvent.click(checkboxes[0]);
    expect(checkboxes[0]).toBeChecked();
    
    // Test processing
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
    });
    
    // Test download button appearance
    expect(screen.getByRole('button', { name: 'Save Processed File' })).toBeInTheDocument();
  });

  test('file selection and deselection workflow', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    
    // Test individual file selection
    fireEvent.click(checkboxes[0]);
    expect(checkboxes[0]).toBeChecked();
    expect(checkboxes[1]).not.toBeChecked();
    
    fireEvent.click(checkboxes[1]);
    expect(checkboxes[0]).toBeChecked();
    expect(checkboxes[1]).toBeChecked();
    
    // Test deselect all
    const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
    fireEvent.click(deselectAllButton);
    
    checkboxes.forEach(checkbox => {
      expect(checkbox).not.toBeChecked();
    });
    
    // Test select all
    const selectAllButton = screen.getByRole('button', { name: 'Select All' });
    fireEvent.click(selectAllButton);
    
    checkboxes.forEach(checkbox => {
      expect(checkbox).toBeChecked();
    });
  });

  test('duplicate removal workflow', async () => {
    // Mock data with duplicates
    mockXLSX.utils.sheet_to_json.mockReturnValue([
      ['Name', 'Age', 'City'],
      ['John', 25, 'New York'],
      ['John', 25, 'New York'], // Duplicate
      ['Jane', 30, 'London']
    ]);
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
      expect(screen.getByText(/Duplicates removed: 1/)).toBeInTheDocument();
    });
  });

  test('error handling workflow', async () => {
    // Mock file loading error
    mockDirectoryHandle.entries.mockReturnValue([
      ['error.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockRejectedValue(new Error('File not found'))
      }]
    ]);
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Failed to load file/)).toBeInTheDocument();
    });
  });

  test('language switching integration', async () => {
    renderApp();
    
    // Test default language
    expect(screen.getByText('Language:')).toBeInTheDocument();
    
    // Change language
    const languageSelect = screen.getByRole('combobox');
    fireEvent.change(languageSelect, { target: { value: 'hu' } });
    
    // Verify language change
    expect(screen.getByText('Nyelv:')).toBeInTheDocument();
  });

  test('state reset workflow', async () => {
    renderApp();
    
    // Complete workflow first
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
    });
    
    // Test state reset
    const chooseNewButton = screen.getByRole('button', { name: 'Choose New Folder' });
    fireEvent.click(chooseNewButton);
    
    // Should return to initial state
    expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    expect(screen.queryByText('test-folder')).not.toBeInTheDocument();
  });

  test('processing with no files selected', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    // Deselect all files
    const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
    fireEvent.click(deselectAllButton);
    
    // Try to process
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText('Please select at least one file to process')).toBeInTheDocument();
    });
  });
});
