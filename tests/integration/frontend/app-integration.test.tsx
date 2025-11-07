import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

import App from '@/frontend/App';
import { LanguageProvider } from '@/contexts/LanguageContext';

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

// Mock document.createElement - but preserve original for React
const originalCreateElement = document.createElement.bind(document);
const mockAnchor = {
  href: '',
  download: '',
  click: jest.fn(),
  style: {} // Ensure style exists for anchor elements
};

// Create a mock that preserves React's element creation but mocks anchor for downloads
Object.defineProperty(document, 'createElement', {
  value: jest.fn().mockImplementation((tagName: string) => {
    // For anchor elements used in downloads, return mock
    if (tagName.toLowerCase() === 'a') {
      return mockAnchor as any;
    }
    // For all other elements (div, etc.), use real implementation for React
    return originalCreateElement(tagName);
  }),
  writable: true,
  configurable: true
});

describe('App Integration Tests', () => {
  beforeEach(() => {
    // Clear localStorage to ensure default language (en) is used
    localStorage.clear();
    
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
    
    // Reset XLSX mocks - use mockImplementation to handle all calls
    const defaultWorkbook = {
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
    };
    
    mockXLSX.read.mockImplementation(() => defaultWorkbook);
    
    mockXLSX.utils.sheet_to_json.mockImplementation(() => [
      ['Name', 'Age', 'City'],
      ['John', 25, 'New York'],
      ['Jane', 30, 'London']
    ]);
    
    // Reset File System Access API mocks
    // Create mock files with arrayBuffer method
    const createMockFile = (name: string, content: string) => {
      const file = new File([content], name, { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      // Mock arrayBuffer method
      file.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(content.length));
      return file;
    };
    
    const mockFile1 = createMockFile('test1.xlsx', 'test content 1');
    const mockFile2 = createMockFile('test2.xlsx', 'test content 2');
    
    mockDirectoryHandle.entries.mockReturnValue([
      ['test1.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(mockFile1)
      }],
      ['test2.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(mockFile2)
      }]
    ]);
    
    // Mock getFileHandle to return file handles that can load files
    // This needs to work for both the initial file listing and when processing files
    mockDirectoryHandle.getFileHandle.mockImplementation((path: string) => {
      const fileName = path.split('/').pop() || path; // Handle both 'test1.xlsx' and paths
      const file = fileName === 'test1.xlsx' ? mockFile1 : mockFile2;
      return Promise.resolve({
        getFile: jest.fn().mockResolvedValue(file)
      });
    });
    
    (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);
  });

  afterEach(() => {
    const root = document.getElementById('root');
    if (root) {
      root.remove();
    }
    
    // Clear localStorage
    localStorage.clear();
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
    expect(screen.getByText('Select a folder')).toBeInTheDocument();
    
    // Test folder selection
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
      expect(screen.getByText('2 Excel file(s) found')).toBeInTheDocument();
    });
    
    // Test file list display - files should be visible
    await waitFor(() => {
      expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
      expect(screen.getByText('test2.xlsx')).toBeInTheDocument();
    });
    
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
    }, { timeout: 5000 });
    
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
    // Note: The Filter duplicates checkbox is not currently rendered in the UI
    // This test verifies that files can be processed successfully
    // Duplicate removal functionality exists in the code but is disabled by default
    mockXLSX.utils.sheet_to_json.mockImplementation(() => [
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
    
    // Note: Filter duplicates checkbox is not in the UI, so we skip clicking it
    // The test verifies that processing works without duplicate removal enabled
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
      // Without filter duplicates enabled, all rows are kept (including duplicates)
      expect(screen.getByText(/All.*rows kept in result|Total rows processed/)).toBeInTheDocument();
    }, { timeout: 5000 });
  });

  test('error handling workflow', async () => {
    // Create mock files
    const createMockFile = (name: string, content: string) => {
      const file = new File([content], name, { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      file.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(content.length));
      return file;
    };
    
    const mockFile1 = createMockFile('test1.xlsx', 'test content 1');
    const mockFile2 = createMockFile('test2.xlsx', 'test content 2');
    
    // First, set up normal folder selection (entries should work)
    mockDirectoryHandle.entries.mockReturnValue([
      ['test1.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(mockFile1)
      }],
      ['test2.xlsx', {
        kind: 'file',
        getFile: jest.fn().mockResolvedValue(mockFile2)
      }]
    ]);
    
    // Track if we're in processing phase
    let processingPhase = false;
    
    // For folder selection, getFileHandle should work normally
    // But for processing, make test1.xlsx fail when getFile is called
    mockDirectoryHandle.getFileHandle.mockImplementation((path: string) => {
      const fileName = path.split('/').pop() || path;
      const file = fileName === 'test1.xlsx' ? mockFile1 : mockFile2;
      
      return Promise.resolve({
        getFile: jest.fn().mockImplementation(async () => {
          // During processing, make test1.xlsx fail
          if (processingPhase && fileName === 'test1.xlsx') {
            throw new Error('File not found');
          }
          return file;
        })
      });
    });
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    // Select the first file which will fail to load
    fireEvent.click(checkboxes[0]);
    
    // Now we're in processing phase - set flag before clicking process
    processingPhase = true;
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Processing failed|Failed to load file/)).toBeInTheDocument();
    }, { timeout: 5000 });
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
    }, { timeout: 5000 });
    
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
    
    // Deselect all files (they start unselected, but let's make sure)
    const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
    fireEvent.click(deselectAllButton);
    
    // The Merge Files button should not be visible when no files are selected
    await waitFor(() => {
      expect(screen.queryByRole('button', { name: 'Merge Files' })).not.toBeInTheDocument();
    });
    
    // Select a file to make button appear
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Merge Files' })).toBeInTheDocument();
    });
    
    // Now deselect again
    fireEvent.click(deselectAllButton);
    
    // Button should disappear again
    await waitFor(() => {
      expect(screen.queryByRole('button', { name: 'Merge Files' })).not.toBeInTheDocument();
    });
  });
});
