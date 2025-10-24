// Main Jest setup file for all tests

// Mock DOM polyfills to fix animation issues
Object.defineProperty(window, 'CSS', {
  value: {
    supports: () => false,
    escape: (str: string) => str
  }
});

// Mock all animation-related properties
Object.defineProperty(document.documentElement.style, 'animation', {
  value: '',
  writable: true
});

Object.defineProperty(document.documentElement.style, 'webkitAnimation', {
  value: '',
  writable: true
});

Object.defineProperty(document.documentElement.style, 'mozAnimation', {
  value: '',
  writable: true
});

Object.defineProperty(document.documentElement.style, 'msAnimation', {
  value: '',
  writable: true
});

// Mock WebkitAnimationEvent
Object.defineProperty(window, 'WebkitAnimationEvent', {
  value: class WebkitAnimationEvent {
    constructor() {}
  }
});

// Mock AnimationEvent
Object.defineProperty(window, 'AnimationEvent', {
  value: class AnimationEvent {
    constructor() {}
  }
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

// Export empty object to make this a module
export {};
