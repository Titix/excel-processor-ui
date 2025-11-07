// Integration test setup - simplified without MSW for now

// Mock File System Access API globally
const mockDirectoryHandle = {
  name: 'test-folder',
  entries: () => [
    ['file1.xlsx', { 
      kind: 'file', 
      getFile: () => Promise.resolve(new File(['mock content 1'], 'file1.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      }))
    }],
    ['file2.xlsx', { 
      kind: 'file', 
      getFile: () => Promise.resolve(new File(['mock content 2'], 'file2.xlsx', { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      }))
    }]
  ]
};

// Mock showDirectoryPicker
Object.defineProperty(window, 'showDirectoryPicker', {
  value: jest.fn().mockResolvedValue(mockDirectoryHandle),
  writable: true,
});

// Mock XLSX library
Object.defineProperty(window, 'XLSX', {
  value: {
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
  },
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

// Export empty object to make this a module
export {};
