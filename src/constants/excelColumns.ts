/**
 * Excel Column Names Constants
 * Defines the standard column names used in Excel files for filtering and processing
 */

export const ExcelColumnNames = {
  BIZONYLAT_FAJTA: 'Bizonylat fajta',
  KELTE: 'Kelte',
  TELJESITES: 'Teljesítés',
  BRUTTO_ERTEK_HUF: 'Bruttó érték (HUF)',
  HOL_1: "'Hol'",
  HOL_2: "'Hol'"
} as const;

// Calculated columns (added to output, not in source files)
export const CALCULATED_COLUMNS = {
  HONAP: 'Hónap',
  HET: 'Hét',
  HET_RESZLETESEN: 'Hét Részletesen'
} as const;

/**
 * Get all column names as an array
 */
export const getAllColumnNames = (): readonly string[] => {
  return Object.values(ExcelColumnNames);
};

/**
 * Get column name by key
 */
export const getColumnName = (key: keyof typeof ExcelColumnNames): string => {
  return ExcelColumnNames[key];
};

/**
 * Check if a column name exists in our defined columns
 */
export const isDefinedColumn = (columnName: string): boolean => {
  return Object.values(ExcelColumnNames).includes(columnName as any);
};

/**
 * Get the key for a column name (reverse lookup)
 */
export const getColumnKey = (columnName: string): keyof typeof ExcelColumnNames | undefined => {
  const entry = Object.entries(ExcelColumnNames).find(([_, value]) => value === columnName);
  return entry ? entry[0] as keyof typeof ExcelColumnNames : undefined;
};

/**
 * Type for column names
 */
export type ExcelColumnName = typeof ExcelColumnNames[keyof typeof ExcelColumnNames];

/**
 * Array of column names in order (from source files)
 */
export const COLUMN_NAMES = [
  ExcelColumnNames.BIZONYLAT_FAJTA,
  ExcelColumnNames.KELTE,
  ExcelColumnNames.TELJESITES,
  ExcelColumnNames.BRUTTO_ERTEK_HUF,
  ExcelColumnNames.HOL_1,
  ExcelColumnNames.HOL_2
] as const;

/**
 * Final output columns (source columns + calculated columns)
 */
export const OUTPUT_COLUMN_NAMES = [
  ExcelColumnNames.BIZONYLAT_FAJTA,
  ExcelColumnNames.KELTE,
  ExcelColumnNames.TELJESITES,
  ExcelColumnNames.BRUTTO_ERTEK_HUF,
  ExcelColumnNames.HOL_1,
  ExcelColumnNames.HOL_2,
  CALCULATED_COLUMNS.HONAP,
  CALCULATED_COLUMNS.HET,
  CALCULATED_COLUMNS.HET_RESZLETESEN
] as const;

/**
 * Helper function to get the week number of the year
 */
export const getWeekNumber = (date: Date): number => {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
};

/**
 * Helper function to parse Excel date
 */
export const parseDate = (dateValue: any): Date | null => {
  if (!dateValue) return null;
  
  try {
    // If it's already a Date object
    if (dateValue instanceof Date) {
      return dateValue;
    }
    
    // If it's a number (Excel date serial number)
    if (typeof dateValue === 'number') {
      // Excel dates start from 1900-01-01
      const excelEpoch = new Date(1900, 0, 1);
      excelEpoch.setDate(excelEpoch.getDate() + dateValue - 2); // Subtract 2 because Excel incorrectly treats 1900 as a leap year
      return excelEpoch;
    }
    
    // If it's a string, try to parse it
    if (typeof dateValue === 'string') {
      const parsed = new Date(dateValue);
      if (!isNaN(parsed.getTime())) {
        return parsed;
      }
      // Try common date formats
      const dateWithoutTime = dateValue.split(' ')[0];
      const parsedDate = new Date(dateWithoutTime);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }
    }
    
    // For objects, try to convert to string and parse
    if (typeof dateValue === 'object') {
      const stringValue = String(dateValue);
      const parsed = new Date(stringValue);
      if (!isNaN(parsed.getTime())) {
        return parsed;
      }
    }
    
    return null;
  } catch (error) {
    console.warn('Error parsing date:', dateValue, error);
    return null;
  }
};

/**
 * Get the start date of a week for a given week number and year
 * Returns the Monday of that week (ISO 8601 week)
 */
export const getWeekStartDate = (weekNumber: number, year: number): Date => {
  // ISO 8601: Week 1 is the week containing January 4
  // Find the Monday of the week containing January 4
  const jan4 = new Date(year, 0, 4);
  const jan4Day = jan4.getDay() === 0 ? 7 : jan4.getDay(); // Convert Sunday (0) to 7 (ISO: Mon=1, Sun=7)
  
  // Calculate days to subtract to get to Monday (Jan 4's Monday)
  // If Jan 4 is Monday (1), subtract 0 days
  // If Jan 4 is Tuesday (2), subtract 1 day
  // If Jan 4 is Sunday (7), subtract 6 days
  // Formula: daysToMonday = (jan4Day - 1)
  const daysToMonday = jan4Day - 1;
  
  // Calculate week 1 start (Monday of week containing Jan 4)
  const week1Start = new Date(jan4);
  week1Start.setDate(jan4.getDate() - daysToMonday);
  
  // Calculate the actual week start (Monday of the requested week)
  const weekStart = new Date(week1Start);
  weekStart.setDate(week1Start.getDate() + (weekNumber - 1) * 7);
  
  return weekStart;
};

/**
 * Get the date range for a week in Hungarian format
 * Example: week 1 in 2025 = "Január 1-5"
 */
export const getWeekDateRangeInHungarian = (weekNumber: number, year: number): string => {
  const monthNames: { [key: number]: string } = {
    1: 'Január', 2: 'Február', 3: 'Március', 4: 'Április',
    5: 'Május', 6: 'Június', 7: 'Július', 8: 'Augusztus',
    9: 'Szeptember', 10: 'Október', 11: 'November', 12: 'December'
  };

  // Get the week start date
  const weekStart = getWeekStartDate(weekNumber, year);
  
  // Week end is 6 days after week start
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  
  const startMonth = monthNames[weekStart.getMonth() + 1] || 'Ismeretlen';
  const endMonth = monthNames[weekEnd.getMonth() + 1] || 'Ismeretlen';
  const startDay = weekStart.getDate();
  const endDay = weekEnd.getDate();
  
  if (weekStart.getMonth() === weekEnd.getMonth()) {
    return `${startMonth} ${startDay}-${endDay}`;
  } else {
    return `${startMonth} ${startDay}-${endMonth} ${endDay}`;
  }
};
