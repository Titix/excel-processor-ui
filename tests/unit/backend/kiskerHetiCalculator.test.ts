import { calculateKiskerHeti, KiskerHetiRow } from '@/utils/kiskerHetiCalculator';
import { OUTPUT_COLUMN_NAMES } from '@/constants/excelColumns';

describe('calculateKiskerHeti', () => {
  // Helper to create test data
  const createTestData = (rows: any[][]): any[][] => {
    const headers = [...OUTPUT_COLUMN_NAMES] as any[];
    return [headers, ...rows.map(row => [...row])];
  };

  describe('Basic grouping and summing', () => {
    it('should group by "Hét Részletesen" and sum bolt and web correctly', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        // Row 1: Week 1, bolt, 1000
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        // Row 2: Week 1, bolt, 2000
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        // Row 3: Week 1, web, 1500
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        // Row 4: Week 2, bolt, 3000
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 3000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 8-14'],
        // Row 5: Week 2, web, 2500
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2500, 'web', 'HOL_2', 'Január', 2, '2. Január 8-14'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(2);
      
      // Week 1: bolt = 1000 + 2000 = 3000, web = 1500, grandTotal = 4500
      expect(result[0]).toEqual({
        hetReszletesen: '1. Január 1-7',
        bolt: 3000,
        web: 1500,
        grandTotal: 4500
      });

      // Week 2: bolt = 3000, web = 2500, grandTotal = 5500
      expect(result[1]).toEqual({
        hetReszletesen: '2. Január 8-14',
        bolt: 3000,
        web: 2500,
        grandTotal: 5500
      });
    });

    it('should handle case-insensitive HOL_1 values', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'BOLT', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'Bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'WEB', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2500, 'Web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(3000); // 1000 + 2000
      expect(result[0].web).toBe(4000); // 1500 + 2500
      expect(result[0].grandTotal).toBe(7000);
    });

    it('should ignore rows with HOL_1 values other than bolt/web', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'other', 'HOL_2', 'Január', 1, '1. Január 1-7'], // Should be ignored
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(1000); // Only the bolt row counted
      expect(result[0].web).toBe(1500);
      expect(result[0].grandTotal).toBe(2500);
    });

    it('should skip rows without "Hét Részletesen" value', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 1, ''], // Empty Hét Részletesen
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'web', 'HOL_2', 'Január', 1, null], // Null Hét Részletesen
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(1000); // Only the first row counted
      expect(result[0].web).toBe(0);
    });
  });

  describe('Numeric handling', () => {
    it('should handle numeric Bruttó érték values correctly', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000.5, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000.75, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500.25, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBeCloseTo(3001.25, 2); // 1000.5 + 2000.75
      expect(result[0].web).toBeCloseTo(1500.25, 2);
      expect(result[0].grandTotal).toBeCloseTo(4501.5, 2);
    });

    it('should handle string numeric Bruttó érték values', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', '1000', 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', '2000', 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(3000);
    });

    it('should handle invalid/non-numeric Bruttó érték values as 0', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'invalid', 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', null, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(1000); // Only valid numeric value counted
    });
  });

  describe('Week sorting', () => {
    it('should sort results by week number in ascending order', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Március', 10, '10. Március 5-11'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'bolt', 'HOL_2', 'Február', 5, '5. Február 29-4'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 3000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 8-14'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(4);
      expect(result[0].hetReszletesen).toBe('1. Január 1-7');
      expect(result[1].hetReszletesen).toBe('2. Január 8-14');
      expect(result[2].hetReszletesen).toBe('5. Február 29-4');
      expect(result[3].hetReszletesen).toBe('10. Március 5-11');
    });

    it('should handle weeks without leading number as week 0', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, 'Invalid week format'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 8-14'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(2);
      // Week 0 should come first
      expect(result[0].hetReszletesen).toBe('Invalid week format');
      expect(result[1].hetReszletesen).toBe('2. Január 8-14');
    });
  });

  describe('Edge cases', () => {
    it('should return empty array for empty data', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = [[...OUTPUT_COLUMN_NAMES]]; // Only headers, no data rows

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(0);
    });

    it('should handle multiple rows for same week with different HOL_1 values', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 100, 'bolt', 'HOL_去哪里', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 200, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 300, 'bolt', 'HOL Valuation2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 400, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(600); // 100 + 200 + 300
      expect(result[0].web).toBe(900); // 400 + 500
      expect(result[0].grandTotal).toBe(1500);
    });

    it('should handle whitespace in HOL_1 values', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, '  bolt  ', 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, '  web  ', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);

      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(1000);
      expect(result[0].web).toBe(2000);
    });

    it('should throw error if "Hét Részletesen" column is missing', () => {
      const finalHeaders = ['Kelte', 'Bruttó érték (HUF)', 'HOL_1'];
      const finalData = [[new Date(), 1000, 'bolt']];

      expect(() => {
        calculateKiskerHeti(finalHeaders, finalData as any[][]);
      }).toThrow('Column "Hét Részletesen" not found in headers');
    });

    it('should throw error if "Bruttó érték (HUF)" column is missing', () => {
      const finalHeaders = ['Kelte', 'Hét Részletesen', 'HOL_1'];
      const finalData = [[new Date(), '1. Január 1-7', 'bolt']];

      expect(() => {
        calculateKiskerHeti(finalHeaders, finalData as any[][]);
      }).toThrow('Column "Bruttó érték (HUF)" not found in headers');
    });

    it('should throw error if HOL_1 column index is out of bounds', () => {
      const finalHeaders = ['Kelte', 'Bruttó érték (HUF)', 'Hét Részletesen'];
      const finalData = [[new Date(), 1000, '1. Január 1-7']];

      expect(() => {
        calculateKiskerHeti(finalHeaders, finalData as any[][]);
      }).toThrow('HOL_1 column index out of bounds');
    });

    it('should throw error if required columns are missing', () => {
      const invalidHeaders = ['Bizonylat fajta', 'Kelte'];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      expect(() => {
        calculateKiskerHeti(invalidHeaders, finalData);
      }).toThrow('Column "Hét Részletesen" not found in headers');
    });

    it('should handle null/undefined HOL_1 values', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, null, 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, undefined, 'HOL_2', 'Január', 1, '1. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);
      
      // Only the bolt row should be included (null/undefined HOL_1 values should be skipped)
      expect(result).toHaveLength(1);
      expect(result[0].bolt).toBe(1500);
    });

    it('should handle week sorting when week number parsing fails', () => {
      const finalHeaders = [...OUTPUT_COLUMN_NAMES];
      const finalData = createTestData([
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, 'Invalid. Január 1-7'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 8-14'],
        ['Bizonylat fajta', 'Kelte', 'Teljesítés', 1500, 'bolt', 'HOL_2', 'Január', 1, 'No number format'],
      ]);

      const result = calculateKiskerHeti(finalHeaders, finalData as any[][]);
      
      // Should still work, sorting by week 0 for invalid formats
      expect(result.length).toBeGreaterThan(0);
      // Week 0 (invalid formats) should come before week 2
      const invalidWeekIndex = result.findIndex((r: KiskerHetiRow) => r.hetReszletesen.includes('Invalid') || r.hetReszletesen.includes('No number'));
      const week2Index = result.findIndex((r: KiskerHetiRow) => r.hetReszletesen.includes('2.'));
      
      if (invalidWeekIndex !== -1 && week2Index !== -1) {
        expect(invalidWeekIndex).toBeLessThan(week2Index);
      }
    });
  });
});
