import {
  formatFileSize,
  isValidExcelFile,
  isValidFileSize,
  getSheetNames,
  getSheetData,
  processSheet,
  createDownloadLink,
  getMimeType,
  generateProcessedFileName,
  messageTypes,
  isValidMessageType,
  getMessageIcon,
  formatMessage,
  type MessageType
} from '@/frontend/utils';

describe('Utils', () => {
  describe('formatFileSize', () => {
    test('formats bytes correctly', () => {
      expect(formatFileSize(0)).toBe('0 Bytes');
      expect(formatFileSize(1024)).toBe('1 KB');
      expect(formatFileSize(1024 * 1024)).toBe('1 MB');
      expect(formatFileSize(1024 * 1024 * 1024)).toBe('1 GB');
      
      // Test specific values to ensure we hit the main logic
      expect(formatFileSize(512)).toBe('512 Bytes');
      expect(formatFileSize(1536)).toBe('1.5 KB');
      expect(formatFileSize(1536 * 1024)).toBe('1.5 MB');
      expect(formatFileSize(1536 * 1024 * 1024)).toBe('1.5 GB');
    });

    test('handles negative values', () => {
      expect(formatFileSize(-100)).toBe('0 Bytes');
    });

    test('handles large values', () => {
      // Test values that exceed the sizes array (TB and beyond)
      expect(formatFileSize(1024 * 1024 * 1024 * 1024)).toBe('0 Bytes');
      expect(formatFileSize(1024 * 1024 * 1024 * 1024 * 1024)).toBe('0 Bytes');
    });

    test('formats decimal values correctly', () => {
      expect(formatFileSize(1536)).toBe('1.5 KB');
      expect(formatFileSize(1536 * 1024)).toBe('1.5 MB');
    });
  });

  describe('isValidExcelFile', () => {
    test('validates Excel file extensions', () => {
      expect(isValidExcelFile('test.xlsx')).toBe(true);
      expect(isValidExcelFile('test.xls')).toBe(true);
      expect(isValidExcelFile('TEST.XLSX')).toBe(true);
      expect(isValidExcelFile('test.XLS')).toBe(true);
    });

    test('rejects non-Excel files', () => {
      expect(isValidExcelFile('test.txt')).toBe(false);
      expect(isValidExcelFile('test.pdf')).toBe(false);
      expect(isValidExcelFile('test')).toBe(false);
      expect(isValidExcelFile('')).toBe(false);
    });
  });

  describe('isValidFileSize', () => {
    test('validates file sizes', () => {
      expect(isValidFileSize(1024)).toBe(true);
      expect(isValidFileSize(50 * 1024 * 1024)).toBe(true);
      expect(isValidFileSize(0)).toBe(false);
      expect(isValidFileSize(-100)).toBe(false);
    });

    test('respects custom max size', () => {
      expect(isValidFileSize(100 * 1024 * 1024, 100)).toBe(true);
      expect(isValidFileSize(101 * 1024 * 1024, 100)).toBe(false);
    });
  });

  describe('getSheetNames', () => {
    test('extracts sheet names from workbook', () => {
      const workbook = {
        Sheets: {
          'Sheet1': {},
          'Sheet2': {},
          'Data': {}
        }
      };
      expect(getSheetNames(workbook)).toEqual(['Sheet1', 'Sheet2', 'Data']);
    });

    test('handles empty workbook', () => {
      expect(getSheetNames({})).toEqual([]);
      expect(getSheetNames(null)).toEqual([]);
      expect(getSheetNames(undefined)).toEqual([]);
    });
  });

  describe('getSheetData', () => {
    test('extracts sheet data', () => {
      const workbook = {
        Sheets: {
          'Sheet1': { A1: { v: 'test' } },
          'Sheet2': { B1: { v: 'data' } }
        }
      };
      expect(getSheetData(workbook, 'Sheet1')).toEqual({ A1: { v: 'test' } });
      expect(getSheetData(workbook, 'Sheet2')).toEqual({ B1: { v: 'data' } });
    });

    test('handles missing sheet', () => {
      const workbook = { Sheets: {} };
      expect(getSheetData(workbook, 'Missing')).toBeUndefined();
    });

    test('handles invalid workbook', () => {
      expect(getSheetData(null, 'Sheet1')).toBeUndefined();
      expect(getSheetData({}, 'Sheet1')).toBeUndefined();
    });
  });

  describe('processSheet', () => {
    test('processes sheet data correctly', () => {
      const sheetData = {
        A1: { v: 'hello' },
        B1: { v: 'world' },
        C1: { v: 123 }
      };
      const result = processSheet(sheetData);
      expect(result.A1.v).toBe('HELLO');
      expect(result.B1.v).toBe('WORLD');
      expect(result.C1.v).toBe(123);
    });

    test('handles empty sheet', () => {
      expect(processSheet({})).toEqual({});
      expect(processSheet(null)).toEqual({});
    });
  });

  describe('createDownloadLink', () => {
    beforeEach(() => {
      // Mock URL.createObjectURL
      global.URL.createObjectURL = jest.fn(() => 'mock-url');
    });

    test('creates download link', () => {
      const data = new ArrayBuffer(8);
      const link = createDownloadLink(data, 'test.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      
      expect(link).toBeDefined();
      expect(link.href).toContain('mock-url');
      expect(link.download).toBe('test.xlsx');
    });
  });

  describe('getMimeType', () => {
    test('returns correct MIME types', () => {
      expect(getMimeType('xls')).toBe('application/vnd.ms-excel');
      expect(getMimeType('xlsx')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      expect(getMimeType('XLS')).toBe('application/vnd.ms-excel');
      expect(getMimeType('XLSX')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    });

    test('returns default MIME type for unknown extensions', () => {
      expect(getMimeType('txt')).toBe('application/octet-stream');
      expect(getMimeType('pdf')).toBe('application/octet-stream');
    });
  });

  describe('generateProcessedFileName', () => {
    test('generates processed file names', () => {
      expect(generateProcessedFileName('test.xlsx')).toBe('test_processed.xlsx');
      expect(generateProcessedFileName('data.xls')).toBe('data_processed.xls');
      expect(generateProcessedFileName('TEST.XLSX')).toBe('TEST_processed.XLSX');
    });

    test('handles files without extension', () => {
      expect(generateProcessedFileName('test')).toBe('test');
    });
  });

  describe('messageTypes', () => {
    test('contains expected message types', () => {
      expect(messageTypes).toEqual(['success', 'error', 'info']);
    });
  });

  describe('isValidMessageType', () => {
    test('validates message types', () => {
      expect(isValidMessageType('success')).toBe(true);
      expect(isValidMessageType('error')).toBe(true);
      expect(isValidMessageType('info')).toBe(true);
      expect(isValidMessageType('warning')).toBe(false);
      expect(isValidMessageType('')).toBe(false);
    });
  });

  describe('getMessageIcon', () => {
    test('returns correct icons', () => {
      expect(getMessageIcon('success')).toBe('✅');
      expect(getMessageIcon('error')).toBe('❌');
      expect(getMessageIcon('info')).toBe('ℹ️');
    });
  });

  describe('formatMessage', () => {
    test('formats messages with icons', () => {
      expect(formatMessage('Test message', 'success')).toBe('✅ Test message');
      expect(formatMessage('Error occurred', 'error')).toBe('❌ Error occurred');
      expect(formatMessage('Information', 'info')).toBe('ℹ️ Information');
    });
  });
});