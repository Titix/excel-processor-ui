import { calculateKiskerHeti } from '@/utils/kiskerHetiCalculator';
import { OUTPUT_COLUMN_NAMES } from '@/constants/excelColumns';

/**
 * Simple test cases to validate the KiskerHeti calculation logic
 * These tests verify the core requirement: 
 * - Group by "Hét Részletesen"
 * - Sum "Bruttó érték (HUF)" where HOL_1 = "bolt" 
 * - Sum "Bruttó érték (HUF)" where HOL_1 = "web"
 * - Result: ONE row per unique "Hét Részletesen" value
 */

describe('KiskerHeti Calculator - Simple Validation Tests', () => {
  const createTestData = (rows: any[][]): any[][] => {
    const headers = [...OUTPUT_COLUMN_NAMES] as any[];
    return [headers, ...rows.map(row => [...row])];
  };

  test('Simple case: One week with bolt and web values should produce one row', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    const finalData = createTestData([
      // Week 1: 2 bolt rows + 2 web rows = 1 result row
      ['Bizonylat', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    // Should produce exactly ONE row for Week 1
    expect(result.length).toBe(1);
    expect(result[0].hetReszletesen).toBe('1. Január 1-7');
    // bolt = 1000 + 2000 = 3000
    expect(result[0].bolt).toBe(3000);
    // web = 1500 + 500 = 2000
    expect(result[0].web).toBe(2000);
    // grandTotal = 3000 + 2000 = 5000
    expect(result[0].grandTotal).toBe(5000);
  });

  test('Multiple weeks: Each week should produce exactly one row', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    const finalData = createTestData([
      // Week 1: multiple rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 100, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 200, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 50, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      // Week 2: multiple rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 300, 'bolt', 'HOL_2', 'Január', 2, '2. Január 8-14'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 400, 'web', 'HOL_2', 'Január', 2, '2. Január 8-14'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 100, 'web', 'HOL_2', 'Január', 2, '2. Január 8-14'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    // Should produce exactly TWO rows (one per week)
    expect(result.length).toBe(2);
    
    // Week 1: bolt = 100+200=300, web = 50, grandTotal = 350
    expect(result[0].hetReszletesen).toBe('1. Január 1-7');
    expect(result[0].bolt).toBe(300);
    expect(result[0].web).toBe(50);
    expect(result[0].grandTotal).toBe(350);
    
    // Week 2: bolt = 300, web = 400+100=500, grandTotal = 800
    expect(result[1].hetReszletesen).toBe('2. Január 8-14');
    expect(result[1].bolt).toBe(300);
    expect(result[1].web).toBe(500);
    expect(result[1].grandTotal).toBe(800);
  });

  test('Only bolt values: Should still produce one row with web=0', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    const finalData = createTestData([
      // Week 1: only bolt rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 2000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-7'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    expect(result.length).toBe(1);
    expect(result[0].bolt).toBe(3000);
    expect(result[0].web).toBe(0);
    expect(result[0].grandTotal).toBe(3000);
  });

  test('Only web values: Should still produce one row with bolt=0', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    const finalData = createTestData([
      // Week 1: only web rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 2500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-7'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    expect(result.length).toBe(1);
    expect(result[0].bolt).toBe(0);
    expect(result[0].web).toBe(4000);
    expect(result[0].grandTotal).toBe(4000);
  });

  test('Real example: Multiple files with same week should sum correctly', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    // First file: 3 rows for "24. Június 13-19"
    // Second file: 2 rows for "24. Június 13-19"
    // Note: Using numeric values (12300 instead of 12,300 format)
    const finalData = createTestData([
      // First file rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 12300, 'web', 'HOL_2', 'Június', 24, '24. Június 13-19'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 12300, 'web', 'HOL_2', 'Június', 24, '24. Június 13-19'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 12300, 'bolt', 'HOL_2', 'Június', 24, '24. Június 13-19'],
      // Second file rows
      ['Bizonylat', 'Kelte', 'Teljesítés', 3000, 'web', 'HOL_2', 'Június', 24, '24. Június 13-19'],
      ['Bizonylat', 'Kelte', 'Teljesítés', 3000, 'bolt', 'HOL_2', 'Június', 24, '24. Június 13-19'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    // Should produce exactly ONE row for "24. Június 13-19"
    expect(result.length).toBe(1);
    expect(result[0].hetReszletesen).toBe('24. Június 13-19');
    
    // bolt: 12300 + 3000 = 15300
    expect(result[0].bolt).toBe(15300);
    
    // web: 12300 + 12300 + 3000 = 27600
    expect(result[0].web).toBe(27600);
    
    // Grand Total: 15300 + 27600 = 42900
    expect(result[0].grandTotal).toBe(42900);
  });

  test('Real example with thousands separators: Should parse comma-separated numbers correctly', () => {
    const finalHeaders = [...OUTPUT_COLUMN_NAMES];
    // Test with actual string format from Excel: "6,800.33", "13,899.88", etc.
    const finalData = createTestData([
      ['Számla', '2025.06.02.', '2025.06.02.', '6,800.33', 'web', 'Üres', '6', 23, '23. Június 6-12'],
      ['Számla', '2025.06.02.', '2025.06.02.', '13,899.88', 'bolt', 'Üres', '6', 23, '23. Június 6-12'],
      ['Számla', '2025.06.10.', '2025.06.10.', '-12,300.04', 'web', 'Üres', '6', 24, '24. Június 13-19'],
      ['Számla', '2025.06.10.', '2025.06.10.', '12,300.04', 'web', 'Üres', '6', 24, '24. Június 13-19'],
      ['Számla', '2025.06.10.', '2025.06.10.', '12,300.04', 'bolt', 'Üres', '6', 24, '24. Június 13-19'],
    ]);

    const result = calculateKiskerHeti(finalHeaders, finalData);

    expect(result.length).toBe(2);
    
    // Week 23: bolt=13899.88, web=6800.33, total=20700.21
    expect(result[0].hetReszletesen).toBe('23. Június 6-12');
    expect(result[0].bolt).toBeCloseTo(13899.88, 2);
    expect(result[0].web).toBeCloseTo(6800.33, 2);
    expect(result[0].grandTotal).toBeCloseTo(20700.21, 2);
    
    // Week 24: bolt=12300.04, web=0 (because -12300.04 + 12300.04 = 0), total=12300.04
    expect(result[1].hetReszletesen).toBe('24. Június 13-19');
    expect(result[1].bolt).toBeCloseTo(12300.04, 2);
    expect(result[1].web).toBeCloseTo(0, 2); // -12300.04 + 12300.04 = 0
    expect(result[1].grandTotal).toBeCloseTo(12300.04, 2);
  });
});
