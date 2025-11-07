/**
 * Script to generate Excel test files for backend integration tests
 * Run with: node tests/integration/backend/generate-test-files.js
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const OUTPUT_COLUMN_NAMES = [
  'Bizonylat fajta',
  'Kelte',
  'Teljesítés',
  'Bruttó érték (HUF)',
  "'Hol'",
  "'Hol'",
  'Hónap',
  'Hét',
  'Hét Részletesen'
];

// Ensure test-data directory exists
const testDataDir = path.join(__dirname, 'test-data');
if (!fs.existsSync(testDataDir)) {
  fs.mkdirSync(testDataDir, { recursive: true });
}

/**
 * Helper function to create an Excel file
 */
function createExcelFile(filename, data) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  
  const filePath = path.join(testDataDir, filename);
  XLSX.writeFile(workbook, filePath);
  console.log(`✓ Created: ${filename}`);
}

// Test File 1: Basic kiskerHeti calculation - Week 1 (Jan 1-5) and Week 2 (Jan 6-12)
// Week 1 of 2025 is truncated: January 1-5 (Wed-Sun)
// Week 2 starts on Monday, January 6
createExcelFile('kisker-heti-basic.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-02'), new Date('2025-01-02'), 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-03'), new Date('2025-01-03'), 2000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-04'), new Date('2025-01-04'), 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-06'), new Date('2025-01-06'), 3000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 6-12'],
  ['Faktura', new Date('2025-01-07'), new Date('2025-01-07'), 2500, 'web', 'HOL_2', 'Január', 2, '2. Január 6-12'],
]);

// Test File 2: Case-insensitive HOL_1 values - Week 1 (Jan 1-5)
createExcelFile('kisker-heti-case-insensitive.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-01'), new Date('2025-01-01'), 1000, 'BOLT', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-02'), new Date('2025-01-02'), 2000, 'Bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-03'), new Date('2025-01-03'), 1500, 'WEB', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-04'), new Date('2025-01-04'), 2500, 'Web', 'HOL_2', 'Január', 1, '1. Január 1-5'],
]);

// Test File 3: HOL validation - HOL_1 is "Üres" (should be replaced with HOL_2) - Week 3 (Jan 13-19)
createExcelFile('hol-validation-replace.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-13'), new Date('2025-01-13'), 100000, 'Üres', 'Teszt1', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-14'), new Date('2025-01-14'), 200000, 'Üres', 'Teszt2', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-15'), new Date('2025-01-15'), 300000, 'Original', 'Teszt3', 'Január', 3, '3. Január 13-19'],
]);

// Test File 4: HOL validation - Both HOL_1 and HOL_2 are "Üres" (error case) - Week 3 (Jan 13-19)
createExcelFile('hol-validation-error.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-13'), new Date('2025-01-13'), 100000, 'Üres', 'Teszt1', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-14'), new Date('2025-01-14'), 200000, 'Üres', 'Üres', 'Január', 3, '3. Január 13-19'], // Error row
  ['Faktura', new Date('2025-01-15'), new Date('2025-01-15'), 300000, 'Original', 'Teszt3', 'Január', 3, '3. Január 13-19'],
]);

// Test File 5: Multiple error rows - Week 3 (Jan 13-19)
createExcelFile('hol-validation-multiple-errors.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-13'), new Date('2025-01-13'), 100000, 'Üres', 'Teszt1', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-14'), new Date('2025-01-14'), 200000, 'Üres', 'Üres', 'Január', 3, '3. Január 13-19'], // Error row 1
  ['Faktura', new Date('2025-01-15'), new Date('2025-01-15'), 300000, 'Original', 'Teszt3', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-16'), new Date('2025-01-16'), 400000, 'Üres', 'Üres', 'Január', 3, '3. Január 13-19'], // Error row 2
  ['Faktura', new Date('2025-01-17'), new Date('2025-01-17'), 500000, 'Üres', 'Teszt5', 'Január', 3, '3. Január 13-19'],
  ['Faktura', new Date('2025-01-18'), new Date('2025-01-18'), 600000, 'Üres', 'Üres', 'Január', 3, '3. Január 13-19'], // Error row 3
]);

// Test File 6: Mixed weeks with different HOL_1 values - Week 1 (Jan 1-5) and Week 2 (Jan 6-12)
createExcelFile('kisker-heti-mixed-weeks.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-02'), new Date('2025-01-02'), 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-03'), new Date('2025-01-03'), 2000, 'other', 'HOL_2', 'Január', 1, '1. Január 1-5'], // Should be ignored
  ['Faktura', new Date('2025-01-04'), new Date('2025-01-04'), 1500, 'web', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-06'), new Date('2025-01-06'), 3000, 'bolt', 'HOL_2', 'Január', 2, '2. Január 6-12'],
]);

// Test File 7: Empty Hét Részletesen (should be skipped) - Week 1 (Jan 1-5)
createExcelFile('kisker-heti-empty-week.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-02'), new Date('2025-01-02'), 1000, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-03'), new Date('2025-01-03'), 2000, 'bolt', 'HOL_2', 'Január', 1, ''], // Empty - should be skipped
  ['Faktura', new Date('2025-01-04'), new Date('2025-01-04'), 1500, 'web', 'HOL_2', 'Január', 1, null], // Null - should be skipped
]);

// Test File 8: Numeric values with decimals - Week 1 (Jan 1-5)
createExcelFile('kisker-heti-decimals.xlsx', [
  OUTPUT_COLUMN_NAMES,
  ['Faktura', new Date('2025-01-02'), new Date('2025-01-02'), 1000.5, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-03'), new Date('2025-01-03'), 2000.75, 'bolt', 'HOL_2', 'Január', 1, '1. Január 1-5'],
  ['Faktura', new Date('2025-01-04'), new Date('2025-01-04'), 1500.25, 'web', 'HOL_2', 'Január', 1, '1. Január 1-5'],
]);

console.log('\n✓ All test Excel files generated successfully!');
console.log(`Files location: ${testDataDir}`);

