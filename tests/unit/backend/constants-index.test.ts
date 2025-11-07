/**
 * Test file for constants/index.ts
 */

import * as constantsIndex from '@/constants/index';

describe('Constants Index', () => {
  test('exports all expected constants from excelColumns', () => {
    // Verify that the index file re-exports everything from excelColumns
    expect(constantsIndex).toHaveProperty('ExcelColumnNames');
    expect(constantsIndex).toHaveProperty('CALCULATED_COLUMNS');
    expect(constantsIndex).toHaveProperty('COLUMN_NAMES');
    expect(constantsIndex).toHaveProperty('OUTPUT_COLUMN_NAMES');
    expect(constantsIndex).toHaveProperty('getWeekNumber');
    expect(constantsIndex).toHaveProperty('parseDate');
    expect(constantsIndex).toHaveProperty('getWeekDateRangeInHungarian');
    expect(constantsIndex).toHaveProperty('getWeekStartDate');
    expect(constantsIndex).toHaveProperty('getAllColumnNames');
    expect(constantsIndex).toHaveProperty('getColumnName');
    expect(constantsIndex).toHaveProperty('isDefinedColumn');
    expect(constantsIndex).toHaveProperty('getColumnKey');
  });

  test('exports are functions and can be called', () => {
    expect(typeof constantsIndex.getWeekNumber).toBe('function');
    expect(typeof constantsIndex.parseDate).toBe('function');
    expect(typeof constantsIndex.getWeekDateRangeInHungarian).toBe('function');
    expect(typeof constantsIndex.getWeekStartDate).toBe('function');
    expect(typeof constantsIndex.getAllColumnNames).toBe('function');
    expect(typeof constantsIndex.getColumnName).toBe('function');
    expect(typeof constantsIndex.isDefinedColumn).toBe('function');
    expect(typeof constantsIndex.getColumnKey).toBe('function');
  });
});


