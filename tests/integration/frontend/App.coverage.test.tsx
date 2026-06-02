import React from 'react';
import { render, screen, fireEvent, waitFor, act } from '@testing-library/react';
import '@testing-library/jest-dom';
import { jest } from '@jest/globals';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

import App from '@/frontend/App';
import { LanguageProvider } from '@/contexts/LanguageContext';

// Mock XLSX library
const mockXLSX = {
  read: jest.fn(),
  write: jest.fn().mockReturnValue(new Uint8Array([1, 2, 3, 4, 5])),
  utils: {
    sheet_to_json: jest.fn(),
    aoa_to_sheet: jest.fn().mockReturnValue({ SheetNames: [], Sheets: {} }),
    book_new: jest.fn().mockReturnValue({ SheetNames: [], Sheets: {} }),
    book_append_sheet: jest.fn(),
    encode_cell: jest.fn().mockImplementation((cell: any) => {
      const col = String.fromCharCode(65 + cell.c);
      return `${col}${cell.r + 1}`;
    }),
    decode_range: jest.fn().mockReturnValue({ s: { r: 0, c: 0 }, e: { r: 2, c: 2 } })
  }
};

Object.defineProperty(window, 'XLSX', {
  value: mockXLSX,
  writable: true,
});

// Mock File System Access API
const createMockFile = (name: string, content: string = 'test') => {
  const blob = new Blob([content], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  // Mock arrayBuffer method that XLSX.read needs
  (file as any).arrayBuffer = jest.fn<Promise<ArrayBuffer>, []>().mockResolvedValue(new ArrayBuffer(content.length));
  return file;
};

const mockDirectoryHandle = {
  name: 'test-folder',
  entries: jest.fn(),
  getFileHandle: jest.fn(),
};

Object.defineProperty(window, 'showDirectoryPicker', {
  value: jest.fn(),
  writable: true,
});

Object.defineProperty(window, 'URL', {
  value: {
    createObjectURL: jest.fn().mockReturnValue('mock-url'),
    revokeObjectURL: jest.fn()
  },
  writable: true
});

const mockAnchor = {
  href: '',
  download: '',
  click: jest.fn(),
  remove: jest.fn(),
};

const originalCreateElement = document.createElement.bind(document);
document.createElement = jest.fn((tagName: string) => {
  if (tagName === 'a') {
    return mockAnchor as any;
  }
  return originalCreateElement(tagName);
});

describe('App Component - Coverage Tests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.useFakeTimers();
    localStorage.clear();

    const existingRoot = document.getElementById('root');
    if (existingRoot) {
      existingRoot.remove();
    }
    
    const div = document.createElement('div');
    div.id = 'root';
    document.body.appendChild(div);

    const mockFile1 = createMockFile('test1.xlsx');
    const mockFile2 = createMockFile('test2.xlsx');
    
    // Create file handles that can be used both in entries() and getFileHandle()
    const createFileHandle = (file: File) => ({
        kind: 'file',
      name: file.name,
      getFile: jest.fn<Promise<File>, []>().mockResolvedValue(file)
    });
    
    const fileHandle1 = createFileHandle(mockFile1);
    const fileHandle2 = createFileHandle(mockFile2);
    
    // Mock entries() to return an async iterator
    (mockDirectoryHandle.entries as jest.Mock).mockReturnValue([
      ['test1.xlsx', fileHandle1],
      ['test2.xlsx', fileHandle2]
    ]);
    
    // Mock getFileHandle() to return the appropriate file handle
    (mockDirectoryHandle.getFileHandle as jest.Mock).mockImplementation((path: string) => {
      const fileName = path.split('/').pop() || path;
      if (fileName === 'test1.xlsx') {
        return Promise.resolve(fileHandle1);
      } else if (fileName === 'test2.xlsx') {
        return Promise.resolve(fileHandle2);
      }
      return Promise.reject(new Error('File not found'));
    });
    
    (window as any).showDirectoryPicker.mockResolvedValue(mockDirectoryHandle);

    // Reset XLSX mocks with default implementation
    mockXLSX.read.mockImplementation((data: any) => {
      return {
      SheetNames: ['Sheet1'],
      Sheets: {
        'Sheet1': {
          '!ref': 'A1:F3',
          'A1': { v: 'Bizonylat fajta', t: 's' },
          'B1': { v: 'Kelte', t: 's' },
          'C1': { v: 'Teljesítés', t: 's' },
          'D1': { v: 'Bruttó érték (HUF)', t: 's' },
          'E1': { v: "'Hol'", t: 's' },
          'F1': { v: "'Hol'", t: 's' },
        }
      }
      };
    });

    // Mock sheet_to_json to handle both regular calls and filterAndMergeSelectedColumns calls
    mockXLSX.utils.sheet_to_json.mockImplementation((worksheet: any, options?: any) => {
      if (options && options.header === 1) {
        // Return array of arrays format for filterAndMergeSelectedColumns
        return [
      ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"],
      ['Invoice', '2025-01-15', '2025-01-20', 1000, 'bolt', 'Üres'],
      ['Invoice', '2025-01-16', '2025-01-21', 2000, 'Üres', 'web'],
        ];
      }
      // Default return for processFiles (regular format)
      return [
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"],
        ['Invoice', '2025-01-15', '2025-01-20', 1000, 'bolt', 'Üres'],
        ['Invoice', '2025-01-16', '2025-01-21', 2000, 'web', 'Üres'],
      ];
    });
  });

  afterEach(() => {
    jest.useRealTimers();
    
    // Clear and restore all mocks/spies to ensure clean state between tests
    jest.clearAllMocks();
    jest.restoreAllMocks();
    
    // Ensure document.createElement is restored (in case it was mocked)
    if (document.createElement !== Document.prototype.createElement) {
      document.createElement = Document.prototype.createElement.bind(document);
    }
    
    const root = document.getElementById('root');
    if (root) {
      root.remove();
    }
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

  // Keep a small set of heavy-coverage tests here; other flows are covered in app-integration.test.tsx
  test('showMessage auto-hides success messages after 8 seconds', async () => {
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

    act(() => {
      jest.advanceTimersByTime(8000);
    });

    await waitFor(() => {
      const messageArea = document.querySelector('.message-area');
      if (messageArea) {
        expect(messageArea.textContent).toBe('');
      } else {
        expect(messageArea).not.toBeInTheDocument();
      }
    });
  });

  test('KiskerHeti calculation creates workbook with totals', async () => {
    renderApp();

    const selectButton = screen.getByRole('button', { name: 'Select Folder' });
    fireEvent.click(selectButton);

    await waitFor(() => {
      expect(screen.getByText('test-folder')).toBeInTheDocument();
    });

    // Provide data that triggers KiskerHeti flow via retail weekly branch
    mockXLSX.utils.sheet_to_json.mockImplementation((worksheet: any, options?: any) => {
      if (options && options.header === 1) {
        return [
          ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"],
          ['Invoice', '2025-01-15', '2025-01-20', 1000, 'bolt', 'Üres'],
          ['Invoice', '2025-01-16', '2025-01-21', 2000, 'bolt', 'Üres'],
          ['Invoice', '2025-01-17', '2025-01-22', 1500, 'web', 'Üres'],
        ];
      }
      return [
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"],
        ['Invoice', '2025-01-15', '2025-01-20', 1000, 'bolt', 'Üres'],
      ];
    });

    const checkboxes = screen.getAllByRole('checkbox');
    fireEvent.click(checkboxes[0]);

    const retailWeeklyCheckbox = await waitFor(() => screen.getByLabelText(/Retail Weekly/i));
    fireEvent.click(retailWeeklyCheckbox);

    const filterMergeButton = screen.getByRole('button', { name: /Filter.*Merge/i });
    fireEvent.click(filterMergeButton);

    await waitFor(() => {
      expect(screen.getByText(/Files processed successfully/)).toBeInTheDocument();
    }, { timeout: 5000 });
  });
});

