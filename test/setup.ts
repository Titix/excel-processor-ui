/**
 * Test Setup File
 * This file runs before each test file
 */

import '@testing-library/jest-dom';

// Mock console methods to reduce noise in tests
const originalConsoleError = console.error;
const originalConsoleWarn = console.warn;

beforeAll(() => {
  console.error = (...args: any[]) => {
    if (
      typeof args[0] === 'string' &&
      args[0].includes('Warning: ReactDOM.render is deprecated')
    ) {
      return;
    }
    originalConsoleError.call(console, ...args);
  };

  console.warn = (...args: any[]) => {
    if (
      typeof args[0] === 'string' &&
      (args[0].includes('componentWillReceiveProps') ||
       args[0].includes('componentWillMount'))
    ) {
      return;
    }
    originalConsoleWarn.call(console, ...args);
  };
});

afterAll(() => {
  console.error = originalConsoleError;
  console.warn = originalConsoleWarn;
});

// Mock IntersectionObserver
global.IntersectionObserver = class IntersectionObserver {
  constructor() {}
  disconnect() {}
  observe() {}
  unobserve() {}
};

// Mock ResizeObserver
global.ResizeObserver = class ResizeObserver {
  constructor() {}
  disconnect() {}
  observe() {}
  unobserve() {}
};

// Mock matchMedia
Object.defineProperty(window, 'matchMedia', {
  writable: true,
  value: jest.fn().mockImplementation(query => ({
    matches: false,
    media: query,
    onchange: null,
    addListener: jest.fn(), // deprecated
    removeListener: jest.fn(), // deprecated
    addEventListener: jest.fn(),
    removeEventListener: jest.fn(),
    dispatchEvent: jest.fn(),
  })),
});

// Mock scrollTo
Object.defineProperty(window, 'scrollTo', {
  writable: true,
  value: jest.fn(),
});

// Mock getComputedStyle
Object.defineProperty(window, 'getComputedStyle', {
  value: () => ({
    getPropertyValue: () => '',
  }),
});

// Mock HTMLElement.prototype.scrollIntoView
HTMLElement.prototype.scrollIntoView = jest.fn();

// Mock HTMLElement.prototype.getBoundingClientRect
HTMLElement.prototype.getBoundingClientRect = jest.fn(() => ({
  width: 120,
  height: 120,
  top: 0,
  left: 0,
  bottom: 0,
  right: 0,
  x: 0,
  y: 0,
  toJSON: jest.fn(),
}));

// Mock File and FileReader
global.File = class File {
  constructor(chunks: any[], filename: string, options: any = {}) {
    this.name = filename;
    this.size = chunks.reduce((acc, chunk) => acc + chunk.length, 0);
    this.type = options.type || '';
    this.lastModified = options.lastModified || Date.now();
  }
  
  name: string;
  size: number;
  type: string;
  lastModified: number;
};

global.FileReader = class FileReader {
  onload: ((event: any) => void) | null = null;
  onerror: ((event: any) => void) | null = null;
  result: any = null;
  readyState: number = 0;
  error: any = null;

  readAsArrayBuffer(file: File) {
    setTimeout(() => {
      this.result = new ArrayBuffer(file.size);
      this.readyState = 2; // DONE
      if (this.onload) {
        this.onload({ target: this });
      }
    }, 0);
  }

  readAsText(file: File) {
    setTimeout(() => {
      this.result = 'mock file content';
      this.readyState = 2; // DONE
      if (this.onload) {
        this.onload({ target: this });
      }
    }, 0);
  }
};

// Mock URL.createObjectURL and URL.revokeObjectURL
Object.defineProperty(URL, 'createObjectURL', {
  writable: true,
  value: jest.fn(() => 'mock-object-url'),
});

Object.defineProperty(URL, 'revokeObjectURL', {
  writable: true,
  value: jest.fn(),
});

// Mock Blob
global.Blob = class Blob {
  constructor(parts: any[] = [], options: any = {}) {
    this.size = parts.reduce((acc, part) => acc + part.length, 0);
    this.type = options.type || '';
  }
  
  size: number;
  type: string;
};

// Mock ArrayBuffer
if (!global.ArrayBuffer) {
  global.ArrayBuffer = class ArrayBuffer {
    constructor(length: number) {
      this.byteLength = length;
    }
    
    byteLength: number;
  };
}

// Mock Uint8Array
if (!global.Uint8Array) {
  global.Uint8Array = class Uint8Array extends Array {
    constructor(buffer: ArrayBuffer | number) {
      if (typeof buffer === 'number') {
        super(buffer);
      } else {
        super(buffer.byteLength);
      }
    }
  };
}

// Setup DOM container for each test
beforeEach(() => {
  // Create a div element for React to render into
  const div = document.createElement('div');
  div.id = 'root';
  document.body.appendChild(div);
});

// Clean up after each test
afterEach(() => {
  // Clear all timers
  jest.clearAllTimers();
  
  // Clear all mocks
  jest.clearAllMocks();
  
  // Reset DOM but keep the root div
  const root = document.getElementById('root');
  if (root) {
    root.innerHTML = '';
  }
  
  // Clear localStorage and sessionStorage
  localStorage.clear();
  sessionStorage.clear();
});