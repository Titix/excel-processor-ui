// Main Jest setup file for all tests

// Mock DOM polyfills to fix animation issues
Object.defineProperty(window, 'CSS', {
  value: {
    supports: () => false,
    escape: (str: string) => str
  },
  writable: true,
  configurable: true
});

// Ensure document.documentElement.style exists and is a proper CSSStyleDeclaration
if (!document.documentElement.style) {
  // Create a mock style object that behaves like CSSStyleDeclaration
  const mockStyle = new Proxy({}, {
    get(target, prop) {
      // Return empty string for all style properties
      if (typeof prop === 'string') {
        return '';
      }
      return target[prop as keyof typeof target];
    },
    has(target, prop) {
      // Always return true for 'in' operator checks
      return true;
    },
    ownKeys() {
      return ['animation', 'webkitAnimation', 'mozAnimation', 'msAnimation', 'transition', 'webkitTransition'];
    }
  });
  
  Object.defineProperty(document.documentElement, 'style', {
    value: mockStyle,
    writable: true,
    configurable: true
  });
}

// Ensure CSSStyleDeclaration.prototype has animation property and supports 'in' operator
// React DOM checks for animation using 'in' operator on the prototype
try {
  if (!('animation' in CSSStyleDeclaration.prototype)) {
    Object.defineProperty(CSSStyleDeclaration.prototype, 'animation', {
      value: '',
      writable: true,
      configurable: true,
      enumerable: true
    });
  }
} catch (e) {
  // If CSSStyleDeclaration is not available, create a mock
  (global as any).CSSStyleDeclaration = class CSSStyleDeclaration {
    animation = '';
    webkitAnimation = '';
    mozAnimation = '';
    msAnimation = '';
  };
}

// Mock all animation-related properties on document.documentElement.style
if (document.documentElement.style) {
  Object.defineProperty(document.documentElement.style, 'animation', {
    value: '',
    writable: true,
    configurable: true,
    enumerable: true
  });

  Object.defineProperty(document.documentElement.style, 'webkitAnimation', {
    value: '',
    writable: true,
    configurable: true,
    enumerable: true
  });

  Object.defineProperty(document.documentElement.style, 'mozAnimation', {
    value: '',
    writable: true,
    configurable: true,
    enumerable: true
  });

  Object.defineProperty(document.documentElement.style, 'msAnimation', {
    value: '',
    writable: true,
    configurable: true,
    enumerable: true
  });
}

// Mock WebkitAnimationEvent
Object.defineProperty(window, 'WebkitAnimationEvent', {
  value: class WebkitAnimationEvent {
    constructor() {}
  },
  writable: true,
  configurable: true
});

// Mock AnimationEvent
Object.defineProperty(window, 'AnimationEvent', {
  value: class AnimationEvent {
    constructor() {}
  },
  writable: true,
  configurable: true
});

// Mock IntersectionObserver
(global as any).IntersectionObserver = class IntersectionObserver {
  constructor() {}
  disconnect() {}
  observe() {}
  unobserve() {}
};

// Mock File and FileReader
(global as any).File = class File {
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

(global as any).FileReader = class FileReader {
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

// Mock Blob
(global as any).Blob = class Blob {
  constructor(parts: any[] = [], options: any = {}) {
    this.size = parts.reduce((acc, part) => acc + part.length, 0);
    this.type = options.type || '';
  }
  
  size: number;
  type: string;
};

// Mock ArrayBuffer
if (!(global as any).ArrayBuffer) {
  (global as any).ArrayBuffer = class ArrayBuffer {
    constructor(length: number) {
      this.byteLength = length;
    }
    
    byteLength: number;
  };
}

// Mock Uint8Array
if (!(global as any).Uint8Array) {
  (global as any).Uint8Array = class Uint8Array extends Array {
    constructor(buffer: ArrayBuffer | number) {
      if (typeof buffer === 'number') {
        super(buffer);
      } else {
        super(buffer.byteLength);
      }
    }
  };
}

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

// Don't override appendChild/removeChild - let React use real ones
// Only mock if we need to for specific test scenarios

// Export empty object to make this a module
export {};
