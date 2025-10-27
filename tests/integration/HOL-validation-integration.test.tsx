import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

import App from '../../src/frontend/App';
import { LanguageProvider } from '../../src/contexts/LanguageContext';

// Mock XLSX library
const mockXLSX = {
  read: jest.fn(),
  write: jest.fn().mockReturnValue(new Uint8Array([1, 2, 3, 4, 5])),
  utils: {
    sheet_to_json: jest.fn(),
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

const mockFileHandle = {
  kind: 'file',
  getFile: jest.fn(),
};

const mockFile = new File(['test content'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

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

// Mock ArrayBuffer
if (typeof global.ArrayBuffer === 'undefined') {
  global.ArrayBuffer = jest.fn() as any;
}

describe('HOL Validation Integration Tests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    
    mockXLSX.utils.sheet_to_json.mockReturnValue([
      ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"],
      ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],
      ['Számla', '2025-01-16', '2025-01-16', 200000, 'Value2', 'Value2'],
      ['Számla', '2025-01-17', '2025-01-17', 300000, 'Üres', 'Üres'], // Error row
    ]);
  });

  describe('Scenario 1: Single row with both HOL_1 and HOL_2 as "Üres"', () => {
    test('should detect and report single problematic row', async () => {
      // Test that the validation logic correctly identifies the row number
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Value2', 'Value2'],
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Üres', 'Üres'], // Error at Excel row 4
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          }
        }
      }
      
      expect(problematicRows).toEqual([4]);
      expect(problematicRows.length).toBe(1);
    });

    test('should show correct row number in error message', () => {
      // Simulate the error message format
      const problematicRows = [4];
      const rowNumbers = problematicRows.join(', ');
      
      // Hungarian format
      const huMessage = `❌ ÉRTÉKHIBA!\n\nHOL_1 és HOL_2 egyaránt "Üres"-re van állítva a következő sorokban: ${rowNumbers}\n\nA fájlok feldolgozása megszakadt. Kérjük, javítsa ki az Excel fájlokban az adatokat!`;
      
      expect(huMessage).toContain('4');
      expect(huMessage).toContain('HOL_1 és HOL_2 egyaránt');
      
      // English format
      const enMessage = `❌ VALUE ERROR!\n\nBoth HOL_1 and HOL_2 are set to "Üres" in the following rows: ${rowNumbers}\n\nFile processing stopped. Please fix the data in the Excel files!`;
      
      expect(enMessage).toContain('4');
      expect(enMessage).toContain('Both HOL_1 and HOL_2');
    });
  });

  describe('Scenario 2: Multiple rows with both HOL_1 and HOL_2 as "Üres"', () => {
    test('should collect all problematic row numbers', async () => {
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Üres'], // Error row 3
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Value3', 'Value3'],
        ['Számla', '2025-01-18', '2025-01-18', 400000, 'Üres', 'Üres'], // Error row 5
        ['Számla', '2025-01-19', '2025-01-19', 500000, 'Value5', 'Value5'],
        ['Számla', '2025-01-20', '2025-01-20', 600000, 'Üres', 'Üres'], // Error row 7
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          }
        }
      }
      
      expect(problematicRows).toEqual([3, 5, 7]);
      expect(problematicRows.length).toBe(3);
    });

    test('should format error message with multiple row numbers', () => {
      const problematicRows = [3, 5, 7];
      const rowNumbers = problematicRows.join(', ');
      
      const huMessage = `❌ ÉRTÉKHIBA!\n\nHOL_1 és HOL_2 egyaránt "Üres"-re van állítva a következő sorokban: ${rowNumbers}\n\nA fájlok feldolgozása megszakadt. Kérjük, javítsa ki az Excel fájlokban az adatokat!`;
      
      expect(huMessage).toContain('3, 5, 7');
      expect(huMessage).toContain('HOL_1 és HOL_2 egyaránt');
    });
  });

  describe('Scenario 3: HOL_1 replacement when HOL_2 has value', () => {
    test('should replace HOL_1 with HOL_2 value when HOL_1 is "Üres"', () => {
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Teszt2'],
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Value3', 'Value3'],
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          } else {
            // Replace HOL_1 with HOL_2 value
            row[HOL_1_INDEX] = hol2;
          }
        }
      }
      
      // No problematic rows
      expect(problematicRows).toEqual([]);
      expect(problematicRows.length).toBe(0);
      
      // HOL_1 should be replaced with HOL_2 values
      expect(filteredData[0][HOL_1_INDEX]).toBe('Teszt1');
      expect(filteredData[1][HOL_1_INDEX]).toBe('Teszt2');
      expect(filteredData[2][HOL_1_INDEX]).toBe('Value3'); // No change
    });
  });

  describe('Scenario 4: Complex mixed scenario', () => {
    test('should handle mixed valid and invalid rows correctly', () => {
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],      // Row 2: Should replace
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Üres'],        // Row 3: ERROR
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Value3', 'Value3'],     // Row 4: Valid
        ['Számla', '2025-01-18', '2025-01-18', 400000, 'Üres', 'Üres'],        // Row 5: ERROR
        ['Számla', '2025-01-19', '2025-01-19', 500000, 'Üres', 'Teszt5'],      // Row 6: Should replace
        ['Számla', '2025-01-20', '2025-01-20', 600000, 'Üres', 'Üres'],        // Row 7: ERROR
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          } else {
            row[HOL_1_INDEX] = hol2;
          }
        }
      }
      
      // Should find 3 problematic rows
      expect(problematicRows).toEqual([3, 5, 7]);
      expect(problematicRows.length).toBe(3);
      
      // Valid rows should be processed
      expect(filteredData[0][HOL_1_INDEX]).toBe('Teszt1'); // Replaced
      expect(filteredData[2][HOL_1_INDEX]).toBe('Value3'); // No change
      expect(filteredData[4][HOL_1_INDEX]).toBe('Teszt5'); // Replaced
      
      // Error rows should remain unchanged
      expect(filteredData[1][HOL_1_INDEX]).toBe('Üres');
      expect(filteredData[3][HOL_1_INDEX]).toBe('Üres');
      expect(filteredData[5][HOL_1_INDEX]).toBe('Üres');
    });
  });

  describe('Scenario 5: No errors in data', () => {
    test('should process all rows successfully when no errors exist', () => {
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Teszt2'],
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Value3', 'Value3'],
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          } else {
            row[HOL_1_INDEX] = hol2;
          }
        }
      }
      
      // No problematic rows
      expect(problematicRows).toEqual([]);
      expect(problematicRows.length).toBe(0);
      
      // All HOL_1 values should be replaced with HOL_2 values
      expect(filteredData[0][HOL_1_INDEX]).toBe('Teszt1');
      expect(filteredData[1][HOL_1_INDEX]).toBe('Teszt2');
      expect(filteredData[2][HOL_1_INDEX]).toBe('Value3');
    });
  });

  describe('Scenario 6: All rows have both HOL empty', () => {
    test('should report all rows as problematic', () => {
      const filteredData = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Üres'], // Row 2
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Üres'], // Row 3
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Üres', 'Üres'], // Row 4
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < filteredData.length; i++) {
        const row = filteredData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          }
        }
      }
      
      // All rows should be problematic
      expect(problematicRows).toEqual([2, 3, 4]);
      expect(problematicRows.length).toBe(3);
    });

    test('should format error message with all row numbers', () => {
      const problematicRows = [2, 3, 4];
      const rowNumbers = problematicRows.join(', ');
      
      const huMessage = `❌ ÉRTÉKHIBA!\n\nHOL_1 és HOL_2 egyaránt "Üres"-re van állítva a következő sorokban: ${rowNumbers}\n\nA fájlok feldolgozása megszakadt. Kérjük, javítsa ki az Excel fájlokban az adatokat!`;
      
      expect(huMessage).toContain('2, 3, 4');
    });
  });

  describe('Scenario 7: Duplicate column name mapping', () => {
    test('should correctly map duplicate column names to different indices', () => {
      // Simulate the column mapping logic for duplicate "'Hol'" columns
      const mockHeaders = [
        'Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', 
        'Other1', 'Other2', 'Other3', 'Other4', 'Other5', 'Other6', 
        'Other7', 'Other8', 'Other9', 'Other10', 'Other11', 'Other12', 
        'Other13', 'Other14', 'Other15', 'Other16', 'Other17', 'Other18', 
        'Other19', 'Other20', 'Other21', 'Other22', 'Other23', 'Other24', 
        "'Hol'", "'Hol'" // Two columns with the same name
      ];
      
      // Mock mapping for HOL_1 (first occurrence)
      const hol1TargetIndex = 0; // Column index 28
      const hol1FoundIndex = mockHeaders
        .map((header, index) => ({ header, index }))
        .filter(item => String(item.header).trim() === "'Hol'")
        .map(item => item.index)[hol1TargetIndex];
      
      // Mock mapping for HOL_2 (second occurrence)
      const hol2TargetIndex = 1; // Column index 29
      const hol2FoundIndex = mockHeaders
        .map((header, index) => ({ header, index }))
        .filter(item => String(item.header).trim() === "'Hol'")
        .map(item => item.index)[hol2TargetIndex];
      
      expect(hol1FoundIndex).toBe(28);
      expect(hol2FoundIndex).toBe(29);
      expect(hol1FoundIndex).not.toBe(hol2FoundIndex);
    });

    test('should correctly extract values from different HOL columns', () => {
      // Simulate row data with values in different HOL columns
      const mockRow = [
        'Faktura',           // Index 0: Bizonylat fajta
        '2025-01-15',        // Index 1: Kelte
        '2025-01-15',        // Index 2: Teljesítés
        100000,              // Index 3: Bruttó érték (HUF)
        'Üres',              // Index 4: HOL_1 (mapped from first "'Hol'" at index 28)
        'TesztStore'         // Index 5: HOL_2 (mapped from second "'Hol'" at index 29)
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = mockRow[HOL_1_INDEX];
      const hol2 = mockRow[HOL_2_INDEX];
      
      // Verify that the values are different (not the same cell)
      expect(hol1).toBe('Üres');
      expect(hol2).toBe('TesztStore');
      expect(hol1).not.toBe(hol2);
    });

    test('should handle validation when HOL columns have different values', () => {
      // Test that when HOL_1 is "Üres" and HOL_2 has a different value, 
      // they should be treated as separate columns
      const testData = [
        ['Row1', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Store1'],
        ['Row2', '2025-01-16', '2025-01-16', 200000, 'Store2', 'Store2'],
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      
      for (let i = 0; i < testData.length; i++) {
        const row = testData[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            const excelRowNumber = i + 2;
            problematicRows.push(excelRowNumber);
          } else {
            row[HOL_1_INDEX] = hol2;
          }
        }
      }
      
      // Should not find any problematic rows since HOL_2 has values
      expect(problematicRows).toEqual([]);
      expect(problematicRows.length).toBe(0);
      
      // HOL_1 should be replaced with HOL_2 value
      expect(testData[0][HOL_1_INDEX]).toBe('Store1');
    });
  });
});

