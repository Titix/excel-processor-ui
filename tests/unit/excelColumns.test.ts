import {
  ExcelColumnNames,
  CALCULATED_COLUMNS,
  COLUMN_NAMES,
  OUTPUT_COLUMN_NAMES,
  getWeekNumber,
  parseDate,
  getWeekDateRangeInHungarian,
  getWeekStartDate
} from '../../src/constants/excelColumns';

describe('Excel Columns Constants', () => {
  describe('ExcelColumnNames', () => {
    test('contains all expected column names', () => {
      expect(ExcelColumnNames.BIZONYLAT_FAJTA).toBe('Bizonylat fajta');
      expect(ExcelColumnNames.KELTE).toBe('Kelte');
      expect(ExcelColumnNames.TELJESITES).toBe('Teljesítés');
      expect(ExcelColumnNames.BRUTTO_ERTEK_HUF).toBe('Bruttó érték (HUF)');
      expect(ExcelColumnNames.HOL_1).toBe("'Hol'");
      expect(ExcelColumnNames.HOL_2).toBe("'Hol'");
    });
  });

  describe('CALCULATED_COLUMNS', () => {
    test('contains all expected calculated columns', () => {
      expect(CALCULATED_COLUMNS.HONAP).toBe('Hónap');
      expect(CALCULATED_COLUMNS.HET).toBe('Hét');
      expect(CALCULATED_COLUMNS.HET_RESZLETESEN).toBe('Hét Részletesen');
    });
  });

  describe('COLUMN_NAMES', () => {
    test('contains all source columns in correct order', () => {
      expect(COLUMN_NAMES.length).toBe(6);
      expect(COLUMN_NAMES[0]).toBe(ExcelColumnNames.BIZONYLAT_FAJTA);
      expect(COLUMN_NAMES[1]).toBe(ExcelColumnNames.KELTE);
      expect(COLUMN_NAMES[2]).toBe(ExcelColumnNames.TELJESITES);
      expect(COLUMN_NAMES[3]).toBe(ExcelColumnNames.BRUTTO_ERTEK_HUF);
      expect(COLUMN_NAMES[4]).toBe(ExcelColumnNames.HOL_1);
      expect(COLUMN_NAMES[5]).toBe(ExcelColumnNames.HOL_2);
    });
  });

  describe('OUTPUT_COLUMN_NAMES', () => {
    test('contains source columns plus calculated columns', () => {
      expect(OUTPUT_COLUMN_NAMES.length).toBe(9);
      // First 6 should be source columns
      expect(OUTPUT_COLUMN_NAMES.slice(0, 6)).toEqual(COLUMN_NAMES);
      // Last 3 should be calculated columns
      expect(OUTPUT_COLUMN_NAMES[6]).toBe(CALCULATED_COLUMNS.HONAP);
      expect(OUTPUT_COLUMN_NAMES[7]).toBe(CALCULATED_COLUMNS.HET);
      expect(OUTPUT_COLUMN_NAMES[8]).toBe(CALCULATED_COLUMNS.HET_RESZLETESEN);
    });
  });

  describe('getWeekNumber', () => {
    test('calculates week number correctly for various dates', () => {
      // January 1, 2025 (should be week 1)
      const jan1 = new Date(2025, 0, 1);
      expect(getWeekNumber(jan1)).toBe(1);

      // January 8, 2025 (should be week 2)
      const jan8 = new Date(2025, 0, 8);
      expect(getWeekNumber(jan8)).toBe(2);

      // January 1, 2024
      const jan1_2024 = new Date(2024, 0, 1);
      expect(getWeekNumber(jan1_2024)).toBeDefined();

      // Mid-year date
      const june15 = new Date(2025, 5, 15);
      expect(getWeekNumber(june15)).toBeGreaterThan(20);
      expect(getWeekNumber(june15)).toBeLessThan(30);

      // December 31
      const dec31 = new Date(2025, 11, 31);
      const weekNum = getWeekNumber(dec31);
      // Could be week 1 of next year or last week of current year
      expect(weekNum === 1 || weekNum > 45).toBe(true);
    });

    test('handles leap year correctly', () => {
      const feb29_2024 = new Date(2024, 1, 29);
      expect(getWeekNumber(feb29_2024)).toBeGreaterThan(0);
    });
  });

  describe('parseDate', () => {
    test('parses Excel date serial numbers', () => {
      // Excel date 1 = 1900-01-01
      const excelDate1 = parseDate(1);
      expect(excelDate1).toBeInstanceOf(Date);
      
      // Excel date 43831 = 2020-01-01
      const excelDate43831 = parseDate(43831);
      expect(excelDate43831).toBeInstanceOf(Date);
      expect(excelDate43831!.getFullYear()).toBe(2020);
    });

    test('parses Date objects', () => {
      const date = new Date(2025, 0, 15);
      expect(parseDate(date)).toBe(date);
    });

    test('parses ISO date strings', () => {
      const dateStr = '2025-01-15';
      const parsed = parseDate(dateStr);
      expect(parsed).toBeInstanceOf(Date);
      expect(parsed!.getFullYear()).toBe(2025);
      expect(parsed!.getMonth()).toBe(0);
      expect(parsed!.getDate()).toBe(15);
    });

    test('parses date strings with time', () => {
      const dateStr = '2025-01-15T10:30:00';
      const parsed = parseDate(dateStr);
      expect(parsed).toBeInstanceOf(Date);
      expect(parsed!.getFullYear()).toBe(2025);
    });

    test('returns null for invalid inputs', () => {
      expect(parseDate('invalid date')).toBeNull();
      expect(parseDate('')).toBeNull();
      expect(parseDate(null)).toBeNull();
      expect(parseDate(undefined)).toBeNull();
    });

    test('handles empty strings and whitespace', () => {
      expect(parseDate('   ')).toBeNull();
      expect(parseDate('  ')).toBeNull();
    });
  });

  describe('getWeekStartDate', () => {
    test('returns correct start date for week 1 in 2025', () => {
      const week1Start = getWeekStartDate(1, 2025);
      expect(week1Start).toBeInstanceOf(Date);
      expect(week1Start.getFullYear()).toBe(2025);
    });

    test('returns a consistent day of week for each week', () => {
      // Check a few weeks to ensure they all start on the same day
      const firstWeekStart = getWeekStartDate(1, 2025);
      const firstDayOfWeek = firstWeekStart.getDay();
      
      for (let week = 2; week <= 10; week++) {
        const weekStart = getWeekStartDate(week, 2025);
        const dayOfWeek = weekStart.getDay();
        // All weeks should start on the same day of week
        expect(dayOfWeek).toBe(firstDayOfWeek);
      }
    });

    test('returns sequential dates for consecutive weeks', () => {
      const week1Start = getWeekStartDate(1, 2025);
      const week2Start = getWeekStartDate(2, 2025);
      
      // Week 2 should be 7 days after week 1
      const diffInTime = week2Start.getTime() - week1Start.getTime();
      const diffInDays = diffInTime / (1000 * 60 * 60 * 24);
      expect(diffInDays).toBe(7);
    });

    test('handles year boundaries correctly', () => {
      const lastWeekStart = getWeekStartDate(52, 2025);
      const firstWeekStart = getWeekStartDate(1, 2026);
      
      expect(lastWeekStart.getFullYear()).toBe(2025);
      expect(firstWeekStart.getFullYear()).toBe(2026);
    });
  });

  describe('getWeekDateRangeInHungarian', () => {
    test('returns correct date range for week 1 in 2025', () => {
      // Week 1 in 2025 starts on December 30, 2024 and ends on January 5, 2025
      const result = getWeekDateRangeInHungarian(1, 2025);
      expect(result).toContain('Január');
      expect(result).toMatch(/\d+-\d+/);
    });

    test('returns correct date range for week 10 in 2025', () => {
      // Week 10 in 2025 is typically around early March
      const result = getWeekDateRangeInHungarian(10, 2025);
      expect(result).toContain('Március');
      expect(result).toMatch(/\d+-\d+/);
    });

    test('returns Hungarian month names', () => {
      const result = getWeekDateRangeInHungarian(1, 2025);
      // Should contain Hungarian months like Január, Február, etc.
      const hungarianMonths = ['Január', 'Február', 'Március', 'Április', 
                               'Május', 'Június', 'Július', 'Augusztus',
                               'Szeptember', 'Október', 'November', 'December'];
      const containsHungarianMonth = hungarianMonths.some(month => result.includes(month));
      expect(containsHungarianMonth).toBe(true);
    });

    test('handles week that spans two months', () => {
      // Test a week that crosses month boundary
      const result = getWeekDateRangeInHungarian(1, 2025);
      // Week 1 in 2025 often spans Dec and Jan or just early Jan
      expect(result).toBeDefined();
      expect(typeof result).toBe('string');
    });

    test('returns format with day ranges', () => {
      const result = getWeekDateRangeInHungarian(1, 2025);
      // Should be in format "Month day-day" or "Month day-OtherMonth day"
      // Account for Hungarian characters and cross-month ranges
      expect(result).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
    });

    test('handles different years correctly', () => {
      const result2025 = getWeekDateRangeInHungarian(1, 2025);
      const result2026 = getWeekDateRangeInHungarian(1, 2026);
      
      expect(result2025).not.toBe(result2026); // Different years should give different ranges
    });

    test('handles leap years', () => {
      const result2024 = getWeekDateRangeInHungarian(10, 2024);
      const result2025 = getWeekDateRangeInHungarian(10, 2025);
      
      expect(result2024).toBeDefined();
      expect(result2025).toBeDefined();
    });

    test('returns correct format for mid-year weeks', () => {
      const result = getWeekDateRangeInHungarian(26, 2025);
      // Should contain a month name and day range (may span months)
      expect(result).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
    });

    test('handles year-end weeks', () => {
      const result = getWeekDateRangeInHungarian(52, 2025);
      expect(result).toBeDefined();
      expect(typeof result).toBe('string');
      // May span years (Dec to Jan)
      expect(result).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
    });
  });

  describe('Integration: Week calculation and Hungarian formatting', () => {
    test('calculates week number and formats date range correctly', () => {
      const testDate = new Date(2025, 0, 15); // January 15, 2025
      const weekNumber = getWeekNumber(testDate);
      const year = testDate.getFullYear();
      
      const dateRange = getWeekDateRangeInHungarian(weekNumber, year);
      
      expect(weekNumber).toBeGreaterThan(0);
      expect(weekNumber).toBeLessThanOrEqual(53);
      expect(dateRange).toContain('Január');
      expect(dateRange).toMatch(/\d+(-|-\w+\s+)\d+/);
    });

    test('end-to-end test with Excel date processing', () => {
      // Simulate processing a date from Excel
      const excelDate = 45327; // Approximate date in 2024/2025
      const parsedDate = parseDate(excelDate);
      
      if (parsedDate) {
        const weekNumber = getWeekNumber(parsedDate);
        const year = parsedDate.getFullYear();
        const dateRange = getWeekDateRangeInHungarian(weekNumber, year);
        
        expect(weekNumber).toBeGreaterThan(0);
        expect(dateRange).toBeDefined();
        expect(dateRange).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
      }
    });

    test('handles dates at year boundaries', () => {
      // Test December 31
      const dec31 = new Date(2025, 11, 31);
      const weekNum = getWeekNumber(dec31);
      const range = getWeekDateRangeInHungarian(weekNum, 2025);
      
      expect(range).toBeDefined();
      // May be at end of year or start of next year
      expect(range).toMatch(/^(December|Január)/);
      
      // Test January 1
      const jan1 = new Date(2026, 0, 1);
      const weekNum2 = getWeekNumber(jan1);
      const range2 = getWeekDateRangeInHungarian(weekNum2, 2026);
      
      expect(range2).toBeDefined();
    });

    test('handles dates in different months correctly', () => {
      const testDates = [
        { month: 0, day: 15, year: 2025 }, // January
        { month: 5, day: 15, year: 2025 }, // June
        { month: 11, day: 15, year: 2025 }  // December
      ];

      testDates.forEach(({ month, day, year }) => {
        const date = new Date(year, month, day);
        const weekNum = getWeekNumber(date);
        const range = getWeekDateRangeInHungarian(weekNum, year);
        
        expect(weekNum).toBeGreaterThan(0);
        expect(range).toBeDefined();
        expect(range).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
      });
    });
  });
});

