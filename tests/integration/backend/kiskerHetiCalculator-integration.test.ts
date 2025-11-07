/**
 * Integration tests for kiskerHetiCalculator using physical Excel files
 * These tests read from actual Excel files in test-data directory
 */

import { calculateKiskerHeti, KiskerHetiRow } from '@/utils/kiskerHetiCalculator';
import { OUTPUT_COLUMN_NAMES } from '@/constants/excelColumns';
import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

describe('KiskerHeti Calculator - Integration Tests with Physical Excel Files', () => {
  const testDataDir = path.join(__dirname, 'test-data');

  /**
   * Helper function to read Excel file and convert to array format
   */
  function readExcelFile(filename: string): { headers: string[], data: any[][] } {
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

    return { headers, data: [headers, ...rows] };
  }

  describe('Basic grouping and summing', () => {
    it('should group by "Hét Részletesen" and sum bolt and web correctly from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-basic.xlsx');

      const result = calculateKiskerHeti(headers, data);

      expect(result).toHaveLength(2);
      
      // Week 1 (Jan 1-5): bolt = 1000 + 2000 = 3000, web = 1500, grandTotal = 4500
      expect(result[0]).toEqual({
        hetReszletesen: '1. Január 1-5',
        bolt: 3000,
        web: 1500,
        grandTotal: 4500
      });

      // Week 2 (Jan 6-12): bolt = 3000, web = 2500, grandTotal = 5500
      expect(result[1]).toEqual({
        hetReszletesen: '2. Január 6-12',
        bolt: 3000,
        web: 2500,
        grandTotal: 5500
      });
    });

    it('should handle case-insensitive HOL_1 values from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-case-insensitive.xlsx');

      const result = calculateKiskerHeti(headers, data);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(3000); // 1000 + 2000
      expect(result[0].web).toBe(4000); // 1500 + 2500
      expect(result[0].grandTotal).toBe(7000);
    });

    it('should ignore rows with HOL_1 values other than bolt/web from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-mixed-weeks.xlsx');

      const result = calculateKiskerHeti(headers, data);

      expect(result).toHaveLength(2);
      // Week 1: only bolt and web rows counted (other row ignored)
      expect(result[0].bolt).toBe(1000);
      expect(result[0].web).toBe(1500);
      expect(result[0].grandTotal).toBe(2500);
      
      // Week 2: only bolt row counted
      expect(result[1].bolt).toBe(3000);
      expect(result[1].web).toBe(0);
      expect(result[1].grandTotal).toBe(3000);
    });

    it('should skip rows without "Hét Részletesen" value from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-empty-week.xlsx');

      const result = calculateKiskerHeti(headers, data);

      expect(result).toHaveLength(1);
      // Only the first row with valid "Hét Részletesen" should be counted
      expect(result[0].bolt).toBe(1000);
      expect(result[0].web).toBe(0);
    });
  });

  describe('Numeric handling', () => {
    it('should handle numeric Bruttó érték values with decimals from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-decimals.xlsx');

      const result = calculateKiskerHeti(headers, data);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBeCloseTo(3001.25, 2); // 1000.5 + 2000.75
      expect(result[0].web).toBeCloseTo(1500.25, 2);
      expect(result[0].grandTotal).toBeCloseTo(4501.5, 2);
    });
  });

  describe('File reading validation', () => {
    it('should throw error if file does not exist', () => {
      expect(() => {
        readExcelFile('non-existent-file.xlsx');
      }).toThrow('Test file not found');
    });

    it('should correctly read headers from Excel file', () => {
      const { headers } = readExcelFile('kisker-heti-basic.xlsx');
      
      expect(headers).toHaveLength(OUTPUT_COLUMN_NAMES.length);
      expect(headers[0]).toBe('Bizonylat fajta');
      expect(headers[3]).toBe('Bruttó érték (HUF)');
      expect(headers[8]).toBe('Hét Részletesen');
    });

    it('should correctly read data rows from Excel file', () => {
      const { headers, data } = readExcelFile('kisker-heti-basic.xlsx');
      
      // Should have headers + 5 data rows
      expect(data.length).toBe(6); // 1 header + 5 rows
      expect(data[0]).toEqual(headers);
      expect(data[1][0]).toBe('Faktura'); // First data row
    });
  });
});

