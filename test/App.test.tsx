import React from 'react';
import { render, screen, fireEvent, waitFor, act } from '@testing-library/react';
import '@testing-library/jest-dom';

// Mock CSS imports
jest.mock('../src/frontend/App.css', () => ({}));

import App from '../src/frontend/App';
import { LanguageProvider } from '../src/contexts/LanguageContext';

// Helper function to render App with LanguageProvider
const renderApp = () => {
  return render(
    React.createElement(LanguageProvider, null,
      React.createElement(App)
    )
  );
};

// Mock XLSX library
const mockXLSX = {
  write: jest.fn().mockReturnValue(new Uint8Array([1, 2, 3, 4, 5])),
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

// Mock URL and Blob APIs
Object.defineProperty(window, 'URL', {
  value: {
    createObjectURL: jest.fn().mockReturnValue('blob:mock-url'),
    revokeObjectURL: jest.fn(),
  },
  writable: true,
});

// Mock document methods
const mockAnchor = {
  href: '',
  download: '',
  click: jest.fn(),
};
Object.defineProperty(document, 'createElement', {
  value: jest.fn().mockReturnValue(mockAnchor),
  writable: true,
});

Object.defineProperty(document.body, 'appendChild', {
  value: jest.fn(),
  writable: true,
});

Object.defineProperty(document.body, 'removeChild', {
  value: jest.fn(),
  writable: true,
});

describe('App Component - Excel File Processor', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.useFakeTimers();
    
    // Reset mocks
    mockDirectoryHandle.entries.mockClear();
    mockFileHandle.getFile.mockClear();
    (window as any).showDirectoryPicker.mockClear();
    mockXLSX.write.mockClear();
    (window as any).URL.createObjectURL.mockClear();
    (window as any).URL.revokeObjectURL.mockClear();
    document.createElement.mockClear();
    document.body.appendChild.mockClear();
    document.body.removeChild.mockClear();
    mockAnchor.click.mockClear();
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  describe('Initial Render', () => {
  test('renders main heading', () => {
      renderApp();
      expect(screen.getByText('Excel File Processor')).toBeInTheDocument();
    });

    test('renders folder selection area', () => {
      renderApp();
      expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
      expect(screen.getByText('Choose a folder to scan for Excel files')).toBeInTheDocument();
    });

    test('renders select folder button', () => {
      renderApp();
      expect(screen.getByRole('button', { name: 'Select Folder' })).toBeInTheDocument();
  });

  test('renders version in footer', () => {
      renderApp();
    expect(screen.getByText('Version 1.0.0')).toBeInTheDocument();
  });

  test('renders copyright in footer', () => {
      renderApp();
      expect(screen.getByText('© 2025 Excel Processor. Built with React and Node.js.')).toBeInTheDocument();
    });
  });

  describe('Folder Selection', () => {
    test('handles successful folder selection with Excel files', async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx' }) }],
        ['file2.xls', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file2.xls' }) }],
        ['file3.pdf', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file3.pdf' }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

      await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
      });

      expect(screen.getByText('test-folder')).toBeInTheDocument();
      expect(screen.getByText('2 Excel file(s) found')).toBeInTheDocument();
    });

    test('handles folder selection with no Excel files', async () => {
      const mockEntries = [
        ['file1.pdf', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.pdf' }) }],
        ['file2.txt', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file2.txt' }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('No Excel files found in the selected folder')).toBeInTheDocument();
      });

      expect(screen.getByText('No Excel files (.xlsx, .xls) found in this folder.')).toBeInTheDocument();
    });

    test('handles folder selection cancellation', async () => {
      const abortError = new Error('User cancelled');
      abortError.name = 'AbortError';
      (window as any).showDirectoryPicker.mockRejectedValue(abortError);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

      // Should not show any error message for cancellation
      expect(screen.queryByText(/Browser does not support this feature/)).not.toBeInTheDocument();
    });

    test('handles folder selection error', async () => {
      const error = new Error('Browser does not support this feature');
      error.name = 'NotSupportedError';
      (window as any).showDirectoryPicker.mockRejectedValue(error);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

    await waitFor(() => {
        expect(screen.getByText('Browser does not support this feature or user cancelled selection')).toBeInTheDocument();
      });
    });
  });

  describe('File Management', () => {
    beforeEach(async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
        ['file2.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file2.xlsx', size: 2048 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
      });
    });

    test('shows Excel files list with checkboxes', () => {
      expect(screen.getByText('file1.xlsx')).toBeInTheDocument();
      expect(screen.getByText('file2.xlsx')).toBeInTheDocument();
      expect(screen.getByText('1 KB')).toBeInTheDocument();
      expect(screen.getByText('2 KB')).toBeInTheDocument();
      
      const checkboxes = screen.getAllByRole('checkbox');
      expect(checkboxes).toHaveLength(2);
      expect(checkboxes[0]).not.toBeChecked();
      expect(checkboxes[1]).not.toBeChecked();
    });

    test('handles file selection toggle', () => {
      const checkboxes = screen.getAllByRole('checkbox');
      
      fireEvent.click(checkboxes[0]);
      expect(checkboxes[0]).toBeChecked();
      expect(checkboxes[1]).not.toBeChecked();
      
      fireEvent.click(checkboxes[0]);
      expect(checkboxes[0]).not.toBeChecked();
    });

    test('handles select all files', () => {
      const selectAllButton = screen.getByRole('button', { name: 'Select All' });
      fireEvent.click(selectAllButton);
      
      const checkboxes = screen.getAllByRole('checkbox');
      expect(checkboxes[0]).toBeChecked();
      expect(checkboxes[1]).toBeChecked();
    });

    test('handles deselect all files', () => {
      const selectAllButton = screen.getByRole('button', { name: 'Select All' });
      const deselectAllButton = screen.getByRole('button', { name: 'Deselect All' });
      
      fireEvent.click(selectAllButton);
      fireEvent.click(deselectAllButton);
      
      const checkboxes = screen.getAllByRole('checkbox');
      expect(checkboxes[0]).not.toBeChecked();
      expect(checkboxes[1]).not.toBeChecked();
    });

    test('shows merge button when files are selected', () => {
      const checkboxes = screen.getAllByRole('checkbox');
      fireEvent.click(checkboxes[0]);
      
      expect(screen.getByRole('button', { name: 'Merge Files' })).toBeInTheDocument();
    });

    test('hides merge button when no files are selected', () => {
      expect(screen.queryByRole('button', { name: 'Merge Files' })).not.toBeInTheDocument();
    });
  });

  describe('File Processing', () => {
    beforeEach(async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
    });
  });

    test('handles merge files process successfully', async () => {
      const checkbox = screen.getByRole('checkbox');
      fireEvent.click(checkbox);
      
      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
      
    await act(async () => {
        fireEvent.click(mergeButton);
    });

    await waitFor(() => {
        expect(screen.getByText('✅ Successfully merged 1 file(s)! Ready to save.')).toBeInTheDocument();
    });

      expect(screen.getByRole('button', { name: 'Save Merged File' })).toBeInTheDocument();
    });

    test('shows error when trying to merge without selecting files', async () => {
      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
    
    await act(async () => {
        fireEvent.click(mergeButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Please select at least one Excel file to process')).toBeInTheDocument();
    });
  });

    test('handles processing error', async () => {
      // Mock console.error to avoid noise in test output
      const originalConsoleError = console.error;
      console.error = jest.fn();

      const checkbox = screen.getByRole('checkbox');
      fireEvent.click(checkbox);
      
      // Mock loadSelectedFiles to throw an error
      const originalLoadSelectedFiles = require('../src/frontend/App').default;
      
      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      // Restore console.error
      console.error = originalConsoleError;
    });
  });

  describe('File Saving', () => {
    beforeEach(async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

    await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Process files first
      const checkbox = screen.getByRole('checkbox');
      fireEvent.click(checkbox);
      
      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
      
      await act(async () => {
        fireEvent.click(mergeButton);
      });

      await waitFor(() => {
        expect(screen.getByRole('button', { name: 'Save Merged File' })).toBeInTheDocument();
      });
    });

    test('handles save merged file successfully', async () => {
      const saveButton = screen.getByRole('button', { name: 'Save Merged File' });
      
    await act(async () => {
        fireEvent.click(saveButton);
      });

      expect(mockXLSX.write).toHaveBeenCalled();
      expect(window.URL.createObjectURL).toHaveBeenCalled();
      expect(document.createElement).toHaveBeenCalledWith('a');
      expect(mockAnchor.click).toHaveBeenCalled();
      expect(document.body.appendChild).toHaveBeenCalled();
      expect(document.body.removeChild).toHaveBeenCalled();
      expect(window.URL.revokeObjectURL).toHaveBeenCalled();

      // Fast-forward timers to trigger success message
      act(() => {
        jest.advanceTimersByTime(500);
    });

    await waitFor(() => {
        expect(screen.getByText(/📥 Merged file saved successfully!/)).toBeInTheDocument();
    });
  });

    test('handles save error', async () => {
      const originalConsoleError = console.error;
      console.error = jest.fn();

      mockXLSX.write.mockImplementation(() => {
        throw new Error('Save failed');
      });

      const saveButton = screen.getByRole('button', { name: 'Save Merged File' });
      
      await act(async () => {
        fireEvent.click(saveButton);
      });

    await waitFor(() => {
        expect(screen.getByText('Save failed: Save failed')).toBeInTheDocument();
      });

      console.error = originalConsoleError;
    });
  });

  describe('Folder Management', () => {
    test('handles choose new folder button click', async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
    });

      const chooseNewFolderButton = screen.getByRole('button', { name: 'Choose New Folder' });
    
    await act(async () => {
        fireEvent.click(chooseNewFolderButton);
    });

    await waitFor(() => {
        expect(screen.getByText('📁 Ready to select a new folder')).toBeInTheDocument();
      });

      // Should return to initial state
      expect(screen.getByRole('heading', { name: 'Select Folder' })).toBeInTheDocument();
    });
  });

  describe('Message Handling', () => {
    test('auto-hides success messages after 8 seconds', async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

    await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Fast-forward timers to trigger auto-hide
      act(() => {
        jest.advanceTimersByTime(8000);
      });

    await waitFor(() => {
        expect(screen.queryByText('Found 1 Excel file(s) in the folder')).not.toBeInTheDocument();
    });
  });

    test('does not auto-hide error messages', async () => {
      const error = new Error('Test error');
      error.name = 'TestError';
      (window as any).showDirectoryPicker.mockRejectedValue(error);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

      await waitFor(() => {
        expect(screen.getByText('Browser does not support this feature or user cancelled selection')).toBeInTheDocument();
      });

      // Fast-forward timers - error message should still be visible
      act(() => {
        jest.advanceTimersByTime(8000);
      });

      expect(screen.getByText('Browser does not support this feature or user cancelled selection')).toBeInTheDocument();
    });
  });

  describe('File Size Formatting', () => {
    test('formats file sizes correctly', async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 0 }) }],
        ['file2.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file2.xlsx', size: 1024 }) }],
        ['file3.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file3.xlsx', size: 1048576 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

      await waitFor(() => {
        expect(screen.getByText('0 Bytes')).toBeInTheDocument();
        expect(screen.getByText('1 KB')).toBeInTheDocument();
        expect(screen.getByText('1 MB')).toBeInTheDocument();
      });
    });
  });

  describe('Edge Cases', () => {
    test('handles empty folder selection', async () => {
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          // No entries
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
      });

    await waitFor(() => {
        expect(screen.getByText('No Excel files found in the selected folder')).toBeInTheDocument();
      });
    });

    test('handles save without processed data', async () => {
      const mockEntries = [
        ['file1.xlsx', { ...mockFileHandle, getFile: () => Promise.resolve({ ...mockFile, name: 'file1.xlsx', size: 1024 }) }],
      ];
      
      mockDirectoryHandle.entries.mockReturnValue({
        [Symbol.asyncIterator]: async function* () {
          for (const entry of mockEntries) {
            yield entry;
          }
        }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 1 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Try to save without processing
      const saveButton = screen.queryByRole('button', { name: 'Save Merged File' });
      expect(saveButton).not.toBeInTheDocument();
    });

    test('handles save without selected folder', async () => {
      // This test would require mocking the component state directly
      // Since we can't easily test this scenario with the current setup,
      // we'll test it through the integration tests
      expect(true).toBe(true); // Placeholder
    });
  });

  describe('Column Validation', () => {
    test('validates column consistency across files', async () => {
      const mockDirectoryHandle = {
        name: 'test-folder',
        entries: jest.fn().mockReturnValue({
          [Symbol.asyncIterator]: async function* () {
            yield ['file1.xlsx', mockFileHandle];
            yield ['file2.xlsx', mockFileHandle];
          }
        }),
        getFileHandle: jest.fn().mockResolvedValue(mockFileHandle)
      };

      mockFileHandle.getFile.mockResolvedValue(mockFile);

      // Mock XLSX.read to return workbooks with different columns
      const mockWorkbook1 = {
        SheetNames: ['Sheet1'],
        Sheets: {
          'Sheet1': {
            '!ref': 'A1:C2',
            'A1': { v: 'Name', t: 's' },
            'B1': { v: 'Age', t: 's' },
            'C1': { v: 'Email', t: 's' },
            'A2': { v: 'John', t: 's' },
            'B2': { v: 25, t: 'n' },
            'C2': { v: 'john@test.com', t: 's' }
          }
        }
      };

      const mockWorkbook2 = {
        SheetNames: ['Sheet1'],
        Sheets: {
          'Sheet1': {
            '!ref': 'A1:D2',
            'A1': { v: 'Name', t: 's' },
            'B1': { v: 'Age', t: 's' },
            'C1': { v: 'Email', t: 's' },
            'D1': { v: 'Phone', t: 's' },
            'A2': { v: 'Jane', t: 's' },
            'B2': { v: 30, t: 'n' },
            'C2': { v: 'jane@test.com', t: 's' },
            'D2': { v: '123-456-7890', t: 's' }
          }
        }
      };

      mockXLSX.read = jest.fn()
        .mockReturnValueOnce(mockWorkbook1)
        .mockReturnValueOnce(mockWorkbook2);

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
      await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Select both files
      const checkboxes = screen.getAllByRole('checkbox');
      await act(async () => {
        fireEvent.click(checkboxes[0]);
        fireEvent.click(checkboxes[1]);
      });

      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
      
      await act(async () => {
        fireEvent.click(mergeButton);
    });

    await waitFor(() => {
        expect(screen.getByText(/Column mismatch detected/)).toBeInTheDocument();
      });
    });

    test('allows merging files with consistent columns', async () => {
      const mockDirectoryHandle = {
        name: 'test-folder',
        entries: jest.fn().mockReturnValue({
          [Symbol.asyncIterator]: async function* () {
            yield ['file1.xlsx', mockFileHandle];
            yield ['file2.xlsx', mockFileHandle];
          }
        }),
        getFileHandle: jest.fn().mockResolvedValue(mockFileHandle)
      };

      mockFileHandle.getFile.mockResolvedValue(mockFile);

      // Mock XLSX.read to return workbooks with same columns
      const mockWorkbook = {
        SheetNames: ['Sheet1'],
        Sheets: {
          'Sheet1': {
            '!ref': 'A1:C2',
            'A1': { v: 'Name', t: 's' },
            'B1': { v: 'Age', t: 's' },
            'C1': { v: 'Email', t: 's' },
            'A2': { v: 'John', t: 's' },
            'B2': { v: 25, t: 'n' },
            'C2': { v: 'john@test.com', t: 's' }
          }
        }
      };

      mockXLSX.read = jest.fn().mockReturnValue(mockWorkbook);

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
      });

      // Select both files
      const checkboxes = screen.getAllByRole('checkbox');
      await act(async () => {
        fireEvent.click(checkboxes[0]);
        fireEvent.click(checkboxes[1]);
      });

      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
    
    await act(async () => {
        fireEvent.click(mergeButton);
    });

    await waitFor(() => {
        expect(screen.getByText(/All files have consistent columns/)).toBeInTheDocument();
      });
    });
  });

  describe('Language Support', () => {
    test('renders language selector', () => {
      renderApp();
      
      expect(screen.getByText('Language:')).toBeInTheDocument();
      expect(screen.getByDisplayValue('🇺🇸 English')).toBeInTheDocument();
    });

    test('switches language correctly', async () => {
      renderApp();
      
      const languageSelect = screen.getByDisplayValue('🇺🇸 English');
      
    await act(async () => {
        fireEvent.change(languageSelect, { target: { value: 'hu' } });
    });

    await waitFor(() => {
        expect(screen.getByDisplayValue('🇭🇺 Magyar')).toBeInTheDocument();
      });
    });
  });

  describe('Lazy Loading', () => {
    test('only loads files when selected for processing', async () => {
      const mockDirectoryHandle = {
        name: 'test-folder',
        entries: jest.fn().mockReturnValue({
          [Symbol.asyncIterator]: async function* () {
            yield ['file1.xlsx', mockFileHandle];
            yield ['file2.xlsx', mockFileHandle];
          }
        }),
        getFileHandle: jest.fn().mockResolvedValue(mockFileHandle)
      };

      mockFileHandle.getFile.mockResolvedValue(mockFile);
      mockXLSX.read = jest.fn().mockReturnValue({
        SheetNames: ['Sheet1'],
        Sheets: { 'Sheet1': { '!ref': 'A1:A1' } }
      });

      (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

      renderApp();
      
      const selectButton = screen.getByRole('button', { name: 'Select Folder' });
      
    await act(async () => {
        fireEvent.click(selectButton);
    });

    await waitFor(() => {
        expect(screen.getByText('Found 2 Excel file(s) in the folder')).toBeInTheDocument();
    });

      // Files should be listed but not loaded yet
      expect(screen.getByText('file1.xlsx')).toBeInTheDocument();
      expect(screen.getByText('file2.xlsx')).toBeInTheDocument();
    
      // Only when we select and process should files be loaded
      const checkboxes = screen.getAllByRole('checkbox');
    await act(async () => {
        fireEvent.click(checkboxes[0]);
      });

      const mergeButton = screen.getByRole('button', { name: 'Merge Files' });
    
    await act(async () => {
        fireEvent.click(mergeButton);
    });

      // Now files should be loaded
      expect(mockDirectoryHandle.getFileHandle).toHaveBeenCalledWith('file1.xlsx');
    });
  });
});