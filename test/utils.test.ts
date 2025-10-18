/**
 * Utility functions tests
 */

describe('File Size Formatting', () => {
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  test('formats zero bytes correctly', () => {
    expect(formatFileSize(0)).toBe('0 Bytes');
  });

  test('formats bytes correctly', () => {
    expect(formatFileSize(500)).toBe('500 Bytes');
  });

  test('formats kilobytes correctly', () => {
    expect(formatFileSize(1024)).toBe('1 KB');
    expect(formatFileSize(1536)).toBe('1.5 KB');
  });

  test('formats megabytes correctly', () => {
    expect(formatFileSize(1024 * 1024)).toBe('1 MB');
    expect(formatFileSize(1024 * 1024 * 2.5)).toBe('2.5 MB');
  });

  test('formats gigabytes correctly', () => {
    expect(formatFileSize(1024 * 1024 * 1024)).toBe('1 GB');
    expect(formatFileSize(1024 * 1024 * 1024 * 1.5)).toBe('1.5 GB');
  });

  test('handles large file sizes', () => {
    expect(formatFileSize(50 * 1024 * 1024)).toBe('50 MB');
    expect(formatFileSize(1024 * 1024 * 1024 * 5)).toBe('5 GB');
  });
});

describe('File Validation', () => {
  const isValidExcelFile = (fileName: string): boolean => {
    return /\.(xlsx?)$/i.test(fileName);
  };

  const isValidFileSize = (fileSize: number, maxSizeMB: number = 50): boolean => {
    return fileSize <= maxSizeMB * 1024 * 1024;
  };

  test('validates Excel file extensions', () => {
    expect(isValidExcelFile('test.xlsx')).toBe(true);
    expect(isValidExcelFile('test.XLSX')).toBe(true);
    expect(isValidExcelFile('test.xls')).toBe(true);
    expect(isValidExcelFile('test.XLS')).toBe(true);
    expect(isValidExcelFile('test.txt')).toBe(false);
    expect(isValidExcelFile('test.pdf')).toBe(false);
    expect(isValidExcelFile('test')).toBe(false);
  });

  test('validates file sizes', () => {
    expect(isValidFileSize(1024)).toBe(true);
    expect(isValidFileSize(50 * 1024 * 1024)).toBe(true);
    expect(isValidFileSize(51 * 1024 * 1024)).toBe(false);
    expect(isValidFileSize(100 * 1024 * 1024)).toBe(false);
  });

  test('validates file sizes with custom max size', () => {
    expect(isValidFileSize(10 * 1024 * 1024, 10)).toBe(true);
    expect(isValidFileSize(11 * 1024 * 1024, 10)).toBe(false);
  });
});

describe('Excel Processing Utilities', () => {
  const mockWorkbook = {
    Sheets: {
      'Sheet1': {
        'A1': { v: 'Name' },
        'B1': { v: 'Age' },
        'A2': { v: 'John' },
        'B2': { v: 25 },
        'A3': { v: 'Jane' },
        'B3': { v: 30 }
      }
    }
  };

  const getSheetNames = (workbook: any): string[] => {
    return Object.keys(workbook.Sheets);
  };

  const getSheetData = (workbook: any, sheetName: string): any => {
    return workbook.Sheets[sheetName];
  };

  const processSheet = (sheetData: any): any => {
    // Simple processing: convert all values to uppercase
    const processedSheet = { ...sheetData };
    Object.keys(processedSheet).forEach(key => {
      if (processedSheet[key].v && typeof processedSheet[key].v === 'string') {
        processedSheet[key].v = processedSheet[key].v.toUpperCase();
      }
    });
    return processedSheet;
  };

  test('gets sheet names from workbook', () => {
    expect(getSheetNames(mockWorkbook)).toEqual(['Sheet1']);
  });

  test('gets sheet data by name', () => {
    const sheetData = getSheetData(mockWorkbook, 'Sheet1');
    expect(sheetData).toBeDefined();
    expect(sheetData['A1'].v).toBe('Name');
    expect(sheetData['B1'].v).toBe('Age');
  });

  test('processes sheet data correctly', () => {
    const sheetData = getSheetData(mockWorkbook, 'Sheet1');
    const processedData = processSheet(sheetData);
    
    expect(processedData['A1'].v).toBe('NAME');
    expect(processedData['B1'].v).toBe('AGE');
    expect(processedData['A2'].v).toBe('JOHN');
    expect(processedData['B2'].v).toBe(25); // Numbers should remain unchanged
  });

  test('handles empty workbook', () => {
    const emptyWorkbook = { Sheets: {} };
    expect(getSheetNames(emptyWorkbook)).toEqual([]);
  });

  test('handles non-existent sheet', () => {
    const sheetData = getSheetData(mockWorkbook, 'NonExistentSheet');
    expect(sheetData).toBeUndefined();
  });
});

describe('Download Utilities', () => {
  const createDownloadLink = (data: ArrayBuffer, fileName: string, mimeType: string): HTMLAnchorElement => {
    const blob = new Blob([data], { type: mimeType });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    return a;
  };

  const getMimeType = (fileExtension: string): string => {
    switch (fileExtension.toLowerCase()) {
      case 'xls':
        return 'application/vnd.ms-excel';
      case 'xlsx':
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      default:
        return 'application/octet-stream';
    }
  };

  const generateProcessedFileName = (originalFileName: string): string => {
    return originalFileName.replace(/\.(xlsx?)$/i, '_processed.$1');
  };

  test('creates download link with correct properties', () => {
    const data = new ArrayBuffer(8);
    const fileName = 'test.xlsx';
    const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    
    const link = createDownloadLink(data, fileName, mimeType);
    
    expect(link.download).toBe(fileName);
    expect(link.href).toBe('mock-object-url');
  });

  test('gets correct MIME types', () => {
    expect(getMimeType('xls')).toBe('application/vnd.ms-excel');
    expect(getMimeType('xlsx')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(getMimeType('XLS')).toBe('application/vnd.ms-excel');
    expect(getMimeType('XLSX')).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    expect(getMimeType('txt')).toBe('application/octet-stream');
  });

  test('generates processed file names correctly', () => {
    expect(generateProcessedFileName('test.xlsx')).toBe('test_processed.xlsx');
    expect(generateProcessedFileName('test.xls')).toBe('test_processed.xls');
    expect(generateProcessedFileName('TEST.XLSX')).toBe('TEST_processed.XLSX');
    expect(generateProcessedFileName('my-file.xlsx')).toBe('my-file_processed.xlsx');
  });
});

describe('Message Display Utilities', () => {
  const messageTypes = ['success', 'error', 'info'] as const;
  type MessageType = typeof messageTypes[number];

  const isValidMessageType = (type: string): type is MessageType => {
    return messageTypes.includes(type as MessageType);
  };

  const getMessageIcon = (type: MessageType): string => {
    switch (type) {
      case 'success':
        return '✅';
      case 'error':
        return '❌';
      case 'info':
        return 'ℹ️';
      default:
        return 'ℹ️';
    }
  };

  const formatMessage = (message: string, type: MessageType): string => {
    const icon = getMessageIcon(type);
    return `${icon} ${message}`;
  };

  test('validates message types', () => {
    expect(isValidMessageType('success')).toBe(true);
    expect(isValidMessageType('error')).toBe(true);
    expect(isValidMessageType('info')).toBe(true);
    expect(isValidMessageType('warning')).toBe(false);
    expect(isValidMessageType('')).toBe(false);
  });

  test('gets correct message icons', () => {
    expect(getMessageIcon('success')).toBe('✅');
    expect(getMessageIcon('error')).toBe('❌');
    expect(getMessageIcon('info')).toBe('ℹ️');
  });

  test('formats messages correctly', () => {
    expect(formatMessage('File uploaded successfully', 'success')).toBe('✅ File uploaded successfully');
    expect(formatMessage('Error occurred', 'error')).toBe('❌ Error occurred');
    expect(formatMessage('Information', 'info')).toBe('ℹ️ Information');
  });
});
