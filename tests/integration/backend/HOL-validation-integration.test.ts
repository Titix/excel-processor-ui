/**
 * Integration tests for HOL validation using physical Excel files
 * These tests read from actual Excel files in test-data directory
 */

import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

describe('HOL Validation - Integration Tests with Physical Excel Files', () => {
  const testDataDir = path.join(__dirname, 'test-data');
  const HOL_1_INDEX = 4;
  const HOL_2_INDEX = 5;

  /**
   * Helper function to read Excel file and convert to array format
   */
  function readExcelFile(filename: string): { headers: string[], rows: any[][] } {
    const filePath = path.join(testDataDir, filename);
    
    if (!fs.existsSync(filePath)) {
      throw new Error(`Test file not found: ${filePath}`);
    }

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to array of arrays
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];
    
    if (data.length === 0) {
      throw new Error(`Excel file is empty: ${filename}`);
    }

    const headers = data[0] as string[];
    const rows = data.slice(1);

    return { headers, rows };
  }

  /**
   * Apply HOL validation rules to rows
   */
  function validateHOLRows(rows: any[][]): { processedRows: any[][], errorRows: number[] } {
    const processedRows: any[][] = [];
    const errorRows: number[] = [];
    let rowNumber = 2; // Excel rows start at 2 (1 is header)

    for (const row of rows) {
      const hol1 = String(row[HOL_1_INDEX] || '').trim();
      const hol2 = String(row[HOL_2_INDEX] || '').trim();

      // Rule 1: Replace HOL_1 with HOL_2 when HOL_1 is "Üres"
      if (hol1 === 'Üres') {
        if (hol2 === 'Üres') {
          // Rule 2: Both are "Üres" - error
          errorRows.push(rowNumber);
        } else {
          // Replace HOL_1 with HOL_2
          row[HOL_1_INDEX] = hol2;
          processedRows.push(row);
        }
      } else {
        // HOL_1 is not "Üres", keep as is
        processedRows.push(row);
      }

      rowNumber++;
    }

    return { processedRows, errorRows };
  }

  describe('Rule 1: Replace HOL_1 with HOL_2 when HOL_1 is "Üres"', () => {
    it('should replace HOL_1 with HOL_2 value when HOL_1 is "Üres" from Excel file', () => {
      const { rows } = readExcelFile('hol-validation-replace.xlsx');

      const { processedRows, errorRows } = validateHOLRows(rows);

      expect(errorRows).toHaveLength(0);
      expect(processedRows).toHaveLength(3);
      
      // First row: HOL_1 was "Üres", should be replaced with "Teszt1"
      expect(processedRows[0][HOL_1_INDEX]).toBe('Teszt1');
      expect(processedRows[0][HOL_2_INDEX]).toBe('Teszt1');
      
      // Second row: HOL_1 was "Üres", should be replaced with "Teszt2"
      expect(processedRows[1][HOL_1_INDEX]).toBe('Teszt2');
      expect(processedRows[1][HOL_2_INDEX]).toBe('Teszt2');
      
      // Third row: HOL_1 was "Original", should remain unchanged
      expect(processedRows[2][HOL_1_INDEX]).toBe('Original');
      expect(processedRows[2][HOL_2_INDEX]).toBe('Teszt3');
    });
  });

  describe('Rule 2: Error when both HOL_1 and HOL_2 are "Üres"', () => {
    it('should detect single error row from Excel file', () => {
      const { rows } = readExcelFile('hol-validation-error.xlsx');

      const { processedRows, errorRows } = validateHOLRows(rows);

      expect(errorRows).toHaveLength(1);
      expect(errorRows[0]).toBe(3); // Row 3 has both HOL_1 and HOL_2 as "Üres"
      
      // Should process 2 rows (first and third)
      expect(processedRows).toHaveLength(2);
      
      // First row should be processed (HOL_1 replaced)
      expect(processedRows[0][HOL_1_INDEX]).toBe('Teszt1');
      
      // Third row should be processed (HOL_1 unchanged)
      expect(processedRows[1][HOL_1_INDEX]).toBe('Original');
    });

    it('should detect multiple error rows from Excel file', () => {
      const { rows } = readExcelFile('hol-validation-multiple-errors.xlsx');

      const { processedRows, errorRows } = validateHOLRows(rows);

      expect(errorRows).toHaveLength(3);
      expect(errorRows).toEqual([3, 5, 7]); // Rows 3, 5, and 7 have errors
      
      // Should process 3 rows (rows 1, 3, and 5)
      expect(processedRows).toHaveLength(3);
      
      // First row: HOL_1 replaced with "Teszt1"
      expect(processedRows[0][HOL_1_INDEX]).toBe('Teszt1');
      
      // Third row: HOL_1 unchanged
      expect(processedRows[1][HOL_1_INDEX]).toBe('Original');
      
      // Fifth row: HOL_1 replaced with "Teszt5"
      expect(processedRows[2][HOL_1_INDEX]).toBe('Teszt5');
    });
  });

  describe('File reading validation', () => {
    it('should throw error if file does not exist', () => {
      expect(() => {
        readExcelFile('non-existent-file.xlsx');
      }).toThrow('Test file not found');
    });

    it('should correctly read headers from Excel file', () => {
      const { headers } = readExcelFile('hol-validation-replace.xlsx');
      
      expect(headers).toHaveLength(9);
      expect(headers[HOL_1_INDEX]).toBe("'Hol'");
      expect(headers[HOL_2_INDEX]).toBe("'Hol'");
    });

    it('should correctly read data rows from Excel file', () => {
      const { headers, rows } = readExcelFile('hol-validation-replace.xlsx');
      
      expect(rows.length).toBe(3);
      expect(rows[0][0]).toBe('Faktura'); // First data row
      expect(rows[0][HOL_1_INDEX]).toBe('Üres');
      expect(rows[0][HOL_2_INDEX]).toBe('Teszt1');
    });
  });
});

