/**
 * Global Test Setup
 * This file runs once before all tests
 */

// Set up test environment
process.env.NODE_ENV = 'test';

// Increase timeout for integration tests
jest.setTimeout(30000);

// Mock environment variables
process.env.REACT_APP_VERSION = '1.0.0';
process.env.REACT_APP_NAME = 'Excel Processor';

// Global test utilities
global.testUtils = {
  // Helper to create mock files
  createMockFile: (name: string, type: string, size: number = 1024) => {
    const file = new File(['mock content'], name, { type });
    Object.defineProperty(file, 'size', { value: size });
    return file;
  },
  
  // Helper to create mock Excel files
  createMockExcelFile: (name: string = 'test.xlsx', size: number = 1024) => {
    return global.testUtils.createMockFile(
      name,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      size
    );
  },
  
  // Helper to wait for async operations
  waitFor: (callback: () => void, timeout: number = 1000) => {
    return new Promise((resolve, reject) => {
      const startTime = Date.now();
      
      const check = () => {
        try {
          callback();
          resolve(undefined);
        } catch (error) {
          if (Date.now() - startTime > timeout) {
            reject(error);
          } else {
            setTimeout(check, 10);
          }
        }
      };
      
      check();
    });
  },
  
  // Helper to create mock workbook
  createMockWorkbook: (sheetName: string = 'Sheet1', data: any = {}) => {
    return {
      Sheets: {
        [sheetName]: data
      }
    };
  },
  
  // Helper to create mock sheet data
  createMockSheetData: (rows: number = 10, cols: number = 5) => {
    const sheetData: any = {};
    
    for (let row = 1; row <= rows; row++) {
      for (let col = 1; col <= cols; col++) {
        const cellRef = String.fromCharCode(64 + col) + row;
        sheetData[cellRef] = {
          v: `Value ${row}-${col}`,
          t: 's'
        };
      }
    }
    
    return sheetData;
  }
};

// Global error handler for unhandled promise rejections
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

// Global error handler for uncaught exceptions
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
});

console.log('🧪 Test environment initialized');

