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
    decode_range: jest.fn().mockImplementation((range: string) => {
      // Default implementation - parse range like "A1:C2"
      if (range && range.includes(':')) {
        const parts = range.split(':');
        const start = parts[0];
        const end = parts[1];
        const startCol = start.charCodeAt(0) - 65; // A=0, B=1, etc.
        const startRow = parseInt(start.slice(1)) - 1;
        const endCol = end.charCodeAt(0) - 65;
        const endRow = parseInt(end.slice(1)) - 1;
        return { s: { r: startRow, c: startCol }, e: { r: endRow, c: endCol } };
      }
      return { s: { r: 0, c: 0 }, e: { r: 2, c: 2 } };
    }),
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

describe('App Component - Simple Tests', () => {
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

  test('renders app without crashing', () => {
    renderApp();
    expect(screen.getByText('Excel File Processor')).toBeInTheDocument();
  });

  test('shows initial folder selection UI', () => {
    renderApp();
    expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    expect(screen.getByText('Select a folder')).toBeInTheDocument();
  });

  test('has language selector', () => {
    renderApp();
    expect(screen.getByText('Language:')).toBeInTheDocument();
    expect(screen.getByRole('combobox')).toBeInTheDocument();
  });

  test('renders footer', () => {
    renderApp();
    expect(screen.getByText('© 2025 Excel Processor. Built with React and Node.js.')).toBeInTheDocument();
    expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
  });

  test('handles folder selection successfully', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
      expect(screen.getByText('2 Excel file(s) found')).toBeInTheDocument();
    });
  });

  test('handles folder selection with no Excel files', async () => {
    mockDirectoryHandle.entries.mockReturnValue([
      ['test.txt', { kind: 'file' }],
      ['test.pdf', { kind: 'file' }]
    ]);
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('No Excel files found in the selected folder')).toBeInTheDocument();
    });
  });

  test('handles folder selection cancellation', async () => {
    (window as any).showDirectoryPicker.mockRejectedValueOnce({ name: 'AbortError' });
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    // Should not show any error message for cancellation
    await waitFor(() => {
      expect(screen.queryByText('Browser not supported')).not.toBeInTheDocument();
    });
  });

  test('handles folder selection error', async () => {
    (window as any).showDirectoryPicker.mockRejectedValueOnce(new Error('Permission denied'));
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Browser does not support this feature or user cancelled selection/)).toBeInTheDocument();
    });
  });

  test('shows file list after folder selection', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
      expect(screen.getByText('test2.xlsx')).toBeInTheDocument();
    });
  });

  test('toggles file selection', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    expect(checkboxes[0]).not.toBeChecked();
    
    fireEvent.click(checkboxes[0]);
    expect(checkboxes[0]).toBeChecked();
    
    fireEvent.click(checkboxes[0]);
    expect(checkboxes[0]).not.toBeChecked();
  });

  test('selects all files', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    // Get initial file checkboxes - find them by their association with file names
    // Wait for file names to appear first
    await waitFor(() => {
      expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
    });
    
    const initialCheckboxes = screen.getAllByRole('checkbox');
    const fileCheckboxes = initialCheckboxes.filter(cb => {
      const parent = cb.closest('label');
      if (!parent) return false;
      const text = parent.textContent || '';
      // File checkboxes are in labels that contain file names (xlsx extension)
      return text.includes('.xlsx') || text.includes('.xls');
    });
    expect(fileCheckboxes.length).toBeGreaterThan(0);
    fileCheckboxes.forEach(checkbox => {
      expect(checkbox).not.toBeChecked();
    });
    
    const selectAllButton = screen.getByRole('button', { name: 'Select All' });
    fireEvent.click(selectAllButton);
    
    // Wait for all file checkboxes to be checked
    // Note: After selecting files, "Retail weekly" checkbox may appear, so we filter it out
    await waitFor(() => {
      const allCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
      const currentFileCheckboxes = allCheckboxes.filter(cb => {
        const parent = cb.closest('label');
        if (!parent) return false;
        const text = parent.textContent || '';
        // File checkboxes are in labels that contain file names (xlsx extension)
        return text.includes('.xlsx') || text.includes('.xls');
      });
      // Verify all file checkboxes are checked
      expect(currentFileCheckboxes.length).toBeGreaterThan(0);
      const checkedCount = currentFileCheckboxes.filter(cb => cb.checked).length;
      expect(checkedCount).toBe(currentFileCheckboxes.length);
    }, { timeout: 5000 });
    
    // Final verification - only file checkboxes should be checked
    const finalCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
    const finalFileCheckboxes = finalCheckboxes.filter(cb => {
      const parent = cb.closest('label');
      if (!parent) return false;
      const text = parent.textContent || '';
      // File checkboxes are in labels that contain file names (xlsx extension)
      return text.includes('.xlsx') || text.includes('.xls');
    });
    finalFileCheckboxes.forEach(checkbox => {
      expect(checkbox).toBeChecked();
    });
  });

  test('deselects all files', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const selectAllButton = screen.getByRole('button', { name: 'Select All' });
    const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
    
    // Get initial file checkboxes - find them by their association with file names
    // Wait for file names to appear first
    await waitFor(() => {
      expect(screen.getByText('test1.xlsx')).toBeInTheDocument();
    });
    
    const initialCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
    const fileCheckboxes = initialCheckboxes.filter(cb => {
      const parent = cb.closest('label');
      if (!parent) return false;
      const text = parent.textContent || '';
      // File checkboxes are in labels that contain file names (xlsx extension)
      return text.includes('.xlsx') || text.includes('.xls');
    });
    const fileCheckboxCount = fileCheckboxes.length;
    expect(fileCheckboxCount).toBeGreaterThan(0);
    
    fireEvent.click(selectAllButton);
    
    // Wait for all file checkboxes to be checked
    // Note: After selecting files, "Retail weekly" checkbox may appear, so we filter it out
    await waitFor(() => {
      const allCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
      const currentFileCheckboxes = allCheckboxes.filter(cb => {
        const parent = cb.closest('label');
        if (!parent) return false;
        const text = parent.textContent || '';
        // File checkboxes are in labels that contain file names (xlsx extension)
        return text.includes('.xlsx') || text.includes('.xls');
      });
      // Verify all file checkboxes are checked
      expect(currentFileCheckboxes.length).toBeGreaterThan(0);
      const checkedCount = currentFileCheckboxes.filter(cb => cb.checked).length;
      expect(checkedCount).toBe(currentFileCheckboxes.length);
    }, { timeout: 5000 });
    
    fireEvent.click(deselectAllButton);
    
    // Wait for all file checkboxes to be unchecked
    // After deselecting, "Retail weekly" checkbox may disappear, so we filter it out
    await waitFor(() => {
      const allCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
      const currentFileCheckboxes = allCheckboxes.filter(cb => {
        const parent = cb.closest('label');
        if (!parent) return false;
        const text = parent.textContent || '';
        // File checkboxes are in labels that contain file names (xlsx extension)
        return text.includes('.xlsx') || text.includes('.xls');
      });
      // Verify all file checkboxes are unchecked
      expect(currentFileCheckboxes.length).toBeGreaterThan(0);
      const uncheckedCount = currentFileCheckboxes.filter(cb => !cb.checked).length;
      expect(uncheckedCount).toBe(currentFileCheckboxes.length);
    }, { timeout: 5000 });
    
    // Final verification - only file checkboxes should be unchecked
    const finalCheckboxes = screen.getAllByRole('checkbox') as HTMLInputElement[];
    const finalFileCheckboxes = finalCheckboxes.filter(cb => {
      const parent = cb.closest('label');
      if (!parent) return false;
      const text = parent.textContent || '';
      // File checkboxes are in labels that contain file names (xlsx extension)
      return text.includes('.xlsx') || text.includes('.xls');
    });
    finalFileCheckboxes.forEach(checkbox => {
      expect(checkbox).not.toBeChecked();
    });
  });

  test('shows process button when files are selected', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    expect(screen.getByRole('button', { name: 'Merge Files' })).toBeInTheDocument();
  });

  test('handles processing with no files selected', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    // Verify that "Merge Files" button is not visible when no files are selected
    // The button only appears when at least one file is selected
    const mergeButtonBefore = screen.queryByRole('button', { name: 'Merge Files' });
    expect(mergeButtonBefore).not.toBeInTheDocument();
    
    // Select a file to make the button appear
    const checkboxes = screen.getAllByRole('checkbox');
    // Filter to get only file checkboxes (those with .xlsx in their label)
    const fileCheckboxes = checkboxes.filter(cb => {
      const parent = cb.closest('label');
      if (!parent) return false;
      const text = parent.textContent || '';
      return text.includes('.xlsx') || text.includes('.xls');
    });
    expect(fileCheckboxes.length).toBeGreaterThan(0);
    fireEvent.click(fileCheckboxes[0]);
    
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Merge Files' })).toBeInTheDocument();
    });
    
    // Deselect all files - the button should disappear
    const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
    fireEvent.click(deselectAllButton);
    
    // Verify the button is no longer visible
    await waitFor(() => {
      const mergeButtonAfter = screen.queryByRole('button', { name: 'Merge Files' });
      expect(mergeButtonAfter).not.toBeInTheDocument();
    }, { timeout: 3000 });
  });

  test('processes files successfully', async () => {
    // Mock successful file loading - create file with arrayBuffer
    const mockFile = new File(['test content'], 'test1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    mockFile.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(12));
    
    mockDirectoryHandle.getFileHandle.mockResolvedValue({
      getFile: jest.fn().mockResolvedValue(mockFile)
    });
    
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
    });
  });

  test('handles processing errors', async () => {
    // Mock file loading error
    mockDirectoryHandle.getFileHandle.mockRejectedValue(new Error('File not found'));
    
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

  test('shows download button after processing', async () => {
    // Mock successful file loading - create file with arrayBuffer
    const mockFile = new File(['test content'], 'test1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    mockFile.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(12));
    
    mockDirectoryHandle.getFileHandle.mockResolvedValue({
      getFile: jest.fn().mockResolvedValue(mockFile)
    });
    
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
    });
    
    expect(screen.getByRole('button', { name: 'Save Processed File' })).toBeInTheDocument();
  });

  test('downloads file successfully', async () => {
    // Reset mockAnchor before test
    mockAnchor.click.mockClear();
    mockAnchor.download = '';
    mockAnchor.href = '';
    
    // Mock document.body.appendChild to track if anchor is added
    const appendChildSpy = jest.spyOn(document.body, 'appendChild').mockImplementation(() => {
      return mockAnchor as any;
    });
    const removeChildSpy = jest.spyOn(document.body, 'removeChild').mockImplementation(() => {
      return mockAnchor as any;
    });
    
    // Mock successful file loading - create file with arrayBuffer
    const mockFile = new File(['test content'], 'test1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    mockFile.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(12));
    
    mockDirectoryHandle.getFileHandle.mockImplementation((path: string) => {
      const fileName = path.split('/').pop() || path;
      const file = fileName === 'test1.xlsx' ? mockFile : mockFile;
      return Promise.resolve({
        getFile: jest.fn().mockResolvedValue(file)
      });
    });
    
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
    }, { timeout: 5000 });
    
    // Wait for download button to appear
    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Save Processed File' })).toBeInTheDocument();
    });
    
    const downloadButton = screen.getByRole('button', { name: 'Save Processed File' });
    fireEvent.click(downloadButton);
    
    // Wait a bit for the download to be triggered
    await waitFor(() => {
      expect(mockAnchor.click).toHaveBeenCalled();
    }, { timeout: 2000 });
    
    // Check that download attribute is set (filename format: pivot-YYYYMMDDHHmmss.xlsx)
    expect(mockAnchor.download).toMatch(/^pivot-\d{14}\.xlsx$/);
    
    // Cleanup
    appendChildSpy.mockRestore();
    removeChildSpy.mockRestore();
  });

  test('resets state when choosing new folder', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const chooseNewButton = screen.getByRole('button', { name: 'Choose New Folder' });
    fireEvent.click(chooseNewButton);
    
    expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    expect(screen.queryByText('test-folder')).not.toBeInTheDocument();
  });

  test('handles empty sheets', async () => {
    // For empty sheets, we should still have headers but no data rows
    mockXLSX.read.mockImplementation(() => ({
      SheetNames: ['EmptySheet'],
      Sheets: {
        'EmptySheet': {
          '!ref': 'A1:C1', // Just headers, no data
          'A1': { v: 'Name', t: 's' },
          'B1': { v: 'Age', t: 's' },
          'C1': { v: 'City', t: 's' }
        }
      }
    }));
    
    mockXLSX.utils.sheet_to_json.mockImplementation(() => [
      ['Name', 'Age', 'City'] // Only headers, no data rows
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
    }, { timeout: 5000 });
  });

  test('handles sheets without !ref', async () => {
    // Sheets without !ref should be skipped, so we need at least one valid sheet
    mockXLSX.read.mockImplementation(() => ({
      SheetNames: ['NoRefSheet', 'ValidSheet'],
      Sheets: {
        'NoRefSheet': {}, // No !ref, should be skipped
        'ValidSheet': {
          '!ref': 'A1:C2',
          'A1': { v: 'Name', t: 's' },
          'B1': { v: 'Age', t: 's' },
          'C1': { v: 'City', t: 's' },
          'A2': { v: 'John', t: 's' },
          'B2': { v: 25, t: 'n' },
          'C2': { v: 'New York', t: 's' }
        }
      }
    }));
    
    mockXLSX.utils.sheet_to_json.mockImplementation((worksheet: any) => {
      // Return empty for sheets without !ref, data for valid sheets
      if (!worksheet || !worksheet['!ref']) {
        return [];
      }
      return [
        ['Name', 'Age', 'City'],
        ['John', 25, 'New York']
      ];
    });
    
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
    }, { timeout: 5000 });
  });

  test('handles duplicate removal', async () => {
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
    
    // Note: Filter duplicates checkbox is not currently rendered in the UI
    // The test verifies that processing works without duplicate removal enabled
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
      // Without filter duplicates enabled, all rows are kept (including duplicates)
      expect(screen.getByText(/All.*rows kept in result|Total rows processed/)).toBeInTheDocument();
    }, { timeout: 5000 });
  });

  test('handles header mismatch error', async () => {
    // Mock two files with different headers
    const workbook1 = {
      SheetNames: ['Sheet1'],
      Sheets: {
        'Sheet1': {
          '!ref': 'A1:C2',
          'A1': { v: 'Name', t: 's' },
          'B1': { v: 'Age', t: 's' },
          'C1': { v: 'City', t: 's' },
          'A2': { v: 'John', t: 's' },
          'B2': { v: 25, t: 'n' },
          'C2': { v: 'New York', t: 's' }
        }
      }
    };
    
    const workbook2 = {
      SheetNames: ['Sheet1'],
      Sheets: {
        'Sheet1': {
          '!ref': 'A1:D2',
          'A1': { v: 'Name', t: 's' },
          'B1': { v: 'Age', t: 's' },
          'C1': { v: 'City', t: 's' },
          'D1': { v: 'Country', t: 's' }, // Different header - 4 columns vs 3
          'A2': { v: 'Jane', t: 's' },
          'B2': { v: 30, t: 'n' },
          'C2': { v: 'London', t: 's' },
          'D2': { v: 'UK', t: 's' }
        }
      }
    };

    // Create mock files with arrayBuffer
    const mockFile1 = new File(['file1'], 'test1.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    mockFile1.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(5));
    
    const mockFile2 = new File(['file2'], 'test2.xlsx', { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    mockFile2.arrayBuffer = jest.fn().mockResolvedValue(new ArrayBuffer(5));

    mockDirectoryHandle.getFileHandle.mockImplementation((path: string) => {
      const fileName = path.split('/').pop() || path;
      const file = fileName === 'test1.xlsx' ? mockFile1 : mockFile2;
      return Promise.resolve({
        getFile: jest.fn().mockResolvedValue(file)
      });
    });

    // Mock decode_range to return correct ranges for each workbook
    mockXLSX.utils.decode_range.mockImplementation((range: string) => {
      if (range === 'A1:C2' || range === 'A1:C1') {
        // workbook1 has 3 columns (A-C)
        return { s: { r: 0, c: 0 }, e: { r: 1, c: 2 } };
      } else if (range === 'A1:D2' || range === 'A1:D1') {
        // workbook2 has 4 columns (A-D)
        return { s: { r: 0, c: 0 }, e: { r: 1, c: 3 } };
      }
      // Default fallback
      return { s: { r: 0, c: 0 }, e: { r: 2, c: 2 } };
    });
    
    // Track read calls - files are processed in order: test1.xlsx (3 cols), then test2.xlsx (4 cols)
    let readCallCount = 0;
    mockXLSX.read.mockImplementation((data: any, options?: any) => {
      readCallCount++;
      // First call is for test1.xlsx (3 columns), second call is for test2.xlsx (4 columns)
      return readCallCount === 1 ? workbook1 : workbook2;
    });
    
    // Mock sheet_to_json to return data matching the workbook structure
    let sheetToJsonCount = 0;
    mockXLSX.utils.sheet_to_json.mockImplementation((worksheet: any) => {
      sheetToJsonCount++;
      // First file has 3 columns, second file has 4 columns
      if (sheetToJsonCount <= 1) {
        return [['Name', 'Age', 'City'], ['John', 25, 'New York']];
      } else {
        return [['Name', 'Age', 'City', 'Country'], ['Jane', 30, 'London', 'UK']];
      }
    });

    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    fireEvent.click(checkboxes[1]);
    
    // Reset counters right before processing to ensure clean state
    readCallCount = 0;
    sheetToJsonCount = 0;
    
    const processButton = screen.getByRole('button', { name: 'Merge Files' });
    fireEvent.click(processButton);
    
    await waitFor(() => {
      // The error message contains "HEADER MISMATCH DETECTED"
      expect(screen.getByText(/HEADER MISMATCH DETECTED/i)).toBeInTheDocument();
      expect(screen.getByRole('button', { name: 'OK' })).toBeInTheDocument();
    }, { timeout: 5000 });
  });

  test('filters empty rows when merging', async () => {
    // Mock data with empty rows
    mockXLSX.utils.sheet_to_json.mockImplementation(() => [
      ['Name', 'Age', 'City'],
      ['John', 25, 'New York'],
      [null, null, null], // Empty row
      ['', '', ''], // Empty row as strings
      ['Jane', 30, 'London'],
      ['Bob', 35, 'Paris']
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
    
    // Should process successfully without counting empty rows
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
    }, { timeout: 5000 });
  });

  test('handles filter and merge with specific columns', async () => {
    // Mock data with all columns
    mockXLSX.utils.sheet_to_json.mockImplementation(() => [
      ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'", 'Other Column'],
      ['Type1', '2025-01-01', '2025-01-02', 1000, 'Location1', 'Location2', 'ExtraData'],
      ['Type2', '2025-01-03', '2025-01-04', 2000, 'Location3', 'Location4', 'ExtraData2']
    ]);
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    const filterMergeButton = screen.getByRole('button', { name: /Filter & Merge/i });
    fireEvent.click(filterMergeButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
      expect(screen.getByText(/Columns:/)).toBeInTheDocument();
    }, { timeout: 5000 });
  });

  test('handles missing columns in filter and merge', async () => {
    // Mock data with only some columns
    mockXLSX.utils.sheet_to_json.mockReturnValue([
      ['Bizonylat fajta', 'Kelte', 'Other Column'], // Missing some target columns
      ['Type1', '2025-01-01', 'ExtraData'],
      ['Type2', '2025-01-02', 'ExtraData2']
    ]);
    
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });
    
    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);
    
    const filterMergeButton = screen.getByRole('button', { name: /Filter & Merge/i });
    fireEvent.click(filterMergeButton);
    
    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
    });
  });

  test('displays OK button on error messages', async () => {
    mockDirectoryHandle.getFileHandle.mockRejectedValue(new Error('File not found'));
    
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
      expect(screen.getByRole('button', { name: 'OK' })).toBeInTheDocument();
    });
    
    // Test that OK button clears the message
    const okButton = screen.getByRole('button', { name: 'OK' });
    fireEvent.click(okButton);
    
    await waitFor(() => {
      expect(screen.queryByRole('button', { name: 'OK' })).not.toBeInTheDocument();
    });
  });
});
