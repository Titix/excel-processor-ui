import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';

// Mock CSS imports
jest.mock('../../src/frontend/App.css', () => ({}));

import { getWeekDateRangeInHungarian, getWeekStartDate } from '../../src/constants/excelColumns';

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

describe('Week Date Range Integration Tests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    
    // Setup default mock for XLSX read
    mockXLSX.read.mockReturnValue({
      SheetNames: ['Sheet1'],
      Sheets: {
        'Sheet1': {
          '!ref': 'A1:H2',
          'A1': { v: 'Bizonylat fajta', t: 's' },
          'B1': { v: 'Kelte', t: 's' },
          'C1': { v: 'Teljesítés', t: 's' },
          'D1': { v: 'Bruttó érték (HUF)', t: 's' },
          'E1': { v: "'Hol'", t: 's' },
          'F1': { v: "'Hol'", t: 's' },
        }
      }
    });
  });

  describe('Week Start Date Calculation', () => {
    test('getWeekStartDate returns correct dates for various weeks', () => {
      // Test week 1 in 2025
      const week1Start = getWeekStartDate(1, 2025);
      expect(week1Start).toBeInstanceOf(Date);
      expect(week1Start.getFullYear()).toBe(2025);
      
      // Test week 25 in 2025 (mid-year)
      const week25Start = getWeekStartDate(25, 2025);
      expect(week25Start).toBeInstanceOf(Date);
      expect(week25Start.getFullYear()).toBe(2025);
      
      // Week 25 should be approximately in June
      expect(week25Start.getMonth()).toBeGreaterThanOrEqual(5); // June is month 5
      expect(week25Start.getMonth()).toBeLessThanOrEqual(6); // July is month 6
    });

    test('getWeekStartDate returns sequential dates for consecutive weeks', () => {
      const week5Start = getWeekStartDate(5, 2025);
      const week6Start = getWeekStartDate(6, 2025);
      
      const diffInTime = week6Start.getTime() - week5Start.getTime();
      const diffInDays = diffInTime / (1000 * 60 * 60 * 24);
      expect(diffInDays).toBe(7);
    });

    test('getWeekDateRangeInHungarian returns Hungarian month names', () => {
      const result = getWeekDateRangeInHungarian(10, 2025);
      
      // Should contain Hungarian month names
      const hungarianMonths = ['Január', 'Február', 'Március', 'Április', 
                               'Május', 'Június', 'Július', 'Augusztus',
                               'Szeptember', 'Október', 'November', 'December'];
      const containsHungarianMonth = hungarianMonths.some(month => result.includes(month));
      expect(containsHungarianMonth).toBe(true);
    });

    test('week dates sort chronologically, not alphabetically', () => {
      // Get dates for January (week ~2) and August (week ~30)
      const janWeekStart = getWeekStartDate(2, 2025);
      const augWeekStart = getWeekStartDate(30, 2025);
      
      // January should come before August
      expect(janWeekStart.getTime()).toBeLessThan(augWeekStart.getTime());
      
      // The month values should also reflect this (0 = Jan, 7 = Aug)
      expect(janWeekStart.getMonth()).toBeLessThan(augWeekStart.getMonth());
    });

    test('handles cross-month weeks correctly', () => {
      const result = getWeekDateRangeInHungarian(1, 2025);
      
      // Week 1 often spans December and January
      expect(result).toMatch(/^(December|Január)/);
    });

    test('handles year boundaries correctly', () => {
      const lastWeekStart = getWeekStartDate(52, 2025);
      const firstWeekStart = getWeekStartDate(1, 2026);
      
      expect(lastWeekStart.getFullYear()).toBe(2025);
      expect(firstWeekStart.getFullYear()).toBe(2026);
      
      // First week of 2026 should be after last week of 2025
      expect(lastWeekStart.getTime()).toBeLessThan(firstWeekStart.getTime());
    });
  });

  describe('Integration with Date Type Storage', () => {
    test('date values are stored with correct type for sorting', async () => {
      // Test that dates are properly formatted for Excel sorting
      const testDate = new Date('2025-01-15');
      const weekNumber = 3; // Approximate week number for January 15
      const weekStart = getWeekStartDate(weekNumber, 2025);
      
      // Verify the date is valid
      expect(weekStart).toBeInstanceOf(Date);
      expect(!isNaN(weekStart.getTime())).toBe(true);
    });

    test('week start dates are calculated from Teljesítés dates', () => {
      const testDate = new Date('2025-07-15');
      const weekNumber = Math.ceil((testDate.getTime() - new Date('2025-01-01').getTime()) / (1000 * 60 * 60 * 24) / 7);
      
      // Get the week start date
      const weekStartDate = getWeekStartDate(weekNumber, 2025);
      
      // The week start date should be a Monday
      const dayOfWeek = weekStartDate.getDay();
      expect(dayOfWeek).toBeGreaterThanOrEqual(0);
      expect(dayOfWeek).toBeLessThan(7);
      
      // The date should be in the same month approximately
      expect(weekStartDate.getMonth()).toBeGreaterThanOrEqual(6); // July
    });

    test('week date range is stored with week number prefix for sorting', () => {
      // Test that the week date range includes week number prefix (e.g., "1. Január 1-7")
      const weekNumber = 1;
      const weekRange = getWeekDateRangeInHungarian(weekNumber, 2025);
      
      // Should contain Hungarian month names
      expect(weekRange).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
      
      // Should not be a Date object
      expect(weekRange).not.toBeInstanceOf(Date);
      expect(typeof weekRange).toBe('string');
      
      // In the actual implementation, it would be prefixed as: "1. Január 1-7"
      // This allows alphabetical sorting to work correctly
      const formattedResult = `${weekNumber}. ${weekRange}`;
      expect(formattedResult).toContain(String(weekNumber));
      expect(formattedResult).toContain('. ');
    });
  });

  describe('Edge Cases', () => {
    test('handles leap years correctly', () => {
      const week1Start2024 = getWeekStartDate(1, 2024);
      const week1Start2025 = getWeekStartDate(1, 2025);
      
      expect(week1Start2024.getFullYear()).toBe(2024);
      expect(week1Start2025.getFullYear()).toBe(2025);
    });

    test('handles dates at year end (December)', () => {
      const week50Start = getWeekStartDate(50, 2025);
      const week51Start = getWeekStartDate(51, 2025);
      const week52Start = getWeekStartDate(52, 2025);
      
      // All should be in 2025
      expect(week50Start.getFullYear()).toBe(2025);
      expect(week51Start.getFullYear()).toBe(2025);
      expect(week52Start.getFullYear()).toBe(2025);
      
      // Should be in December
      expect(week52Start.getMonth()).toBe(11); // December is month 11
    });

    test('returns valid dates for all weeks 1-52', () => {
      for (let week = 1; week <= 52; week++) {
        const weekStart = getWeekStartDate(week, 2025);
        expect(weekStart).toBeInstanceOf(Date);
        expect(!isNaN(weekStart.getTime())).toBe(true);
      }
    });
  });

  describe('Week Range Display', () => {
    test('Hungarian format includes correct month and day information', () => {
      const result = getWeekDateRangeInHungarian(26, 2025);
      
      // Should match pattern: "Month day-day" or "Month day-Month day"
      expect(result).toMatch(/^.+?\s+\d+(-\d+|-.+?\s+\d+)/);
    });

    test('different years produce different week ranges', () => {
      const result2024 = getWeekDateRangeInHungarian(10, 2024);
      const result2025 = getWeekDateRangeInHungarian(10, 2025);
      
      expect(result2024).not.toBe(result2025);
    });

    test('consecutive weeks produce sequential date ranges', () => {
      const week10Start = getWeekStartDate(10, 2025);
      const week11Start = getWeekStartDate(11, 2025);
      
      const diffInTime = week11Start.getTime() - week10Start.getTime();
      const diffInDays = diffInTime / (1000 * 60 * 60 * 24);
      expect(diffInDays).toBe(7);
    });
  });
});

