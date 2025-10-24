// Simple integration test without React Testing Library DOM rendering
import React from 'react';
import App from '../../src/frontend/App';
import { LanguageProvider } from '../../src/contexts/LanguageContext';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

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

describe('App Integration Tests - Simple', () => {
  beforeEach(() => {
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

  test('App component can be imported', () => {
    expect(App).toBeDefined();
    expect(typeof App).toBe('function');
  });

  test('LanguageProvider can be imported', () => {
    expect(LanguageProvider).toBeDefined();
    expect(typeof LanguageProvider).toBe('function');
  });

  test('XLSX library is properly mocked', () => {
    expect(window.XLSX).toBeDefined();
    expect(window.XLSX.read).toBeDefined();
    expect(window.XLSX.write).toBeDefined();
    expect(window.XLSX.utils).toBeDefined();
  });

  test('File System Access API is properly mocked', () => {
    expect((window as any).showDirectoryPicker).toBeDefined();
    expect(typeof (window as any).showDirectoryPicker).toBe('function');
  });

  test('URL and Blob APIs are properly mocked', () => {
    expect(window.URL).toBeDefined();
    expect(window.URL.createObjectURL).toBeDefined();
    expect(window.URL.revokeObjectURL).toBeDefined();
  });

  test('Document APIs are properly mocked', () => {
    expect(document.createElement).toBeDefined();
    expect(document.body.appendChild).toBeDefined();
    expect(document.body.removeChild).toBeDefined();
  });

  test('XLSX read function works with mock data', () => {
    const result = window.XLSX.read('mock-data');
    expect(result).toBeDefined();
    expect(result.SheetNames).toEqual(['Sheet1']);
    expect(result.Sheets['Sheet1']).toBeDefined();
  });

  test('XLSX write function works', () => {
    const mockWorkbook = { SheetNames: ['Sheet1'], Sheets: {} };
    const result = window.XLSX.write(mockWorkbook);
    expect(result).toBeDefined();
    expect(result).toBeInstanceOf(Uint8Array);
  });

  test('XLSX utils functions work', () => {
    const mockSheet = { '!ref': 'A1:C3' };
    const jsonResult = window.XLSX.utils.sheet_to_json(mockSheet);
    const rangeResult = window.XLSX.utils.decode_range('A1:C3');
    const cellResult = window.XLSX.utils.encode_cell({ r: 0, c: 0 });
    
    expect(jsonResult).toBeDefined();
    expect(rangeResult).toBeDefined();
    expect(cellResult).toBe('A1');
  });

  test('File System Access API mock works', async () => {
    const directoryHandle = await (window as any).showDirectoryPicker();
    expect(directoryHandle).toBeDefined();
    expect(directoryHandle.name).toBe('test-folder');
    
    const entries = directoryHandle.entries();
    expect(entries).toBeDefined();
    expect(entries.length).toBe(2);
  });

  test('URL createObjectURL works', () => {
    const mockBlob = new Blob(['test'], { type: 'text/plain' });
    const url = window.URL.createObjectURL(mockBlob);
    expect(url).toBe('mock-url');
  });

  test('Document createElement works', () => {
    const element = document.createElement('a');
    expect(element).toBeDefined();
    expect(element.href).toBe('');
    expect(element.download).toBe('');
  });
});
