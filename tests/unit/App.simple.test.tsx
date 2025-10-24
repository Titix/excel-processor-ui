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

describe('App Component - Simple Tests', () => {
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

  test('renders app without crashing', () => {
    renderApp();
    expect(screen.getByText('Excel File Processor')).toBeInTheDocument();
  });

  test('shows initial folder selection UI', () => {
    renderApp();
    expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    expect(screen.getByText('Choose a folder to scan for Excel files')).toBeInTheDocument();
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
      expect(screen.getByText('2 Excel files found')).toBeInTheDocument();
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
      expect(screen.getByText('Browser not supported')).toBeInTheDocument();
    });
  });

  test('shows file list after folder selection', async () => {
    renderApp();
    
    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);
    
    await waitFor(() => {
      expect(screen.getByText('Excel Files Found')).toBeInTheDocument();
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
    
    const selectAllButton = screen.getByRole('button', { name: 'Select All' });
    fireEvent.click(selectAllButton);
    
    const checkboxes = screen.getAllByRole('checkbox');
    checkboxes.forEach(checkbox => {
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
    
    fireEvent.click(selectAllButton);
    fireEvent.click(deselectAllButton);
    
    const checkboxes = screen.getAllByRole('checkbox');
    checkboxes.forEach(checkbox => {
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

  test('processes files successfully', async () => {
    // Mock successful file loading
    mockDirectoryHandle.getFileHandle.mockResolvedValue({
      getFile: jest.fn().mockResolvedValue(new File(['test content'], 'test1.xlsx'))
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
    // Mock successful file loading
    mockDirectoryHandle.getFileHandle.mockResolvedValue({
      getFile: jest.fn().mockResolvedValue(new File(['test content'], 'test1.xlsx'))
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
    // Mock successful file loading
    mockDirectoryHandle.getFileHandle.mockResolvedValue({
      getFile: jest.fn().mockResolvedValue(new File(['test content'], 'test1.xlsx'))
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
    
    const downloadButton = screen.getByRole('button', { name: 'Save Processed File' });
    fireEvent.click(downloadButton);
    
    await waitFor(() => {
      expect(mockAnchor.click).toHaveBeenCalled();
      expect(mockAnchor.download).toMatch(/^proc_\d{14}\.xlsx$/);
    });
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
    mockXLSX.read.mockReturnValue({
      Sheets: {
        'EmptySheet': {
          '!ref': undefined
        }
      },
      SheetNames: ['EmptySheet']
    });
    
    mockXLSX.utils.sheet_to_json.mockReturnValue([]);
    
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

  test('handles sheets without !ref', async () => {
    mockXLSX.read.mockReturnValue({
      Sheets: {
        'NoRefSheet': {}
      },
      SheetNames: ['NoRefSheet']
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

  test('handles duplicate removal', async () => {
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
      expect(screen.getByText(/Duplicates removed: 1/)).toBeInTheDocument();
    });
  });
});
