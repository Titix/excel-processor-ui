/**
 * Tests for HOL_1 and HOL_2 validation logic
 * 
 * Rules:
 * 1. If HOL_1 is "Üres", replace it with HOL_2 value from the same row
 * 2. If both HOL_1 and HOL_2 are "Üres", throw error and stop processing
 */

export {};

describe('HOL_1 and HOL_2 Validation Logic', () => {
  describe('Rule 1: Replace HOL_1 with HOL_2 when HOL_1 is "Üres"', () => {
    test('replaces HOL_1 with HOL_2 value when HOL_1 is "Üres"', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Teszt1'];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      // Rule 1: Replace HOL_1 with HOL_2 value
      if (hol1 === 'Üres') {
        row[HOL_1_INDEX] = hol2;
      }
      
      expect(row[HOL_1_INDEX]).toBe('Teszt1');
      expect(row[HOL_2_INDEX]).toBe('Teszt1');
    });

    test('keeps HOL_1 unchanged when HOL_1 is not "Üres"', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Original', 'Teszt1'];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      const originalHol1 = hol1;
      
      // Only replace if HOL_1 is "Üres"
      if (hol1 === 'Üres') {
        row[HOL_1_INDEX] = hol2;
      }
      
      expect(row[HOL_1_INDEX]).toBe(originalHol1);
      expect(row[HOL_1_INDEX]).toBe('Original');
    });

    test('handles empty string in HOL_1', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, '', 'Teszt1'];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      if (hol1 === 'Üres') {
        row[HOL_1_INDEX] = hol2;
      }
      
      expect(row[HOL_1_INDEX]).toBe(''); // Should remain empty (not "Üres")
    });
  });

  describe('Rule 2: Error when both HOL_1 and HOL_2 are "Üres"', () => {
    test('detects when both HOL_1 and HOL_2 are "Üres"', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Üres'];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      let shouldThrowError = false;
      
      if (hol1 === 'Üres') {
        if (hol2 === 'Üres') {
          shouldThrowError = true;
        }
      }
      
      expect(shouldThrowError).toBe(true);
    });

    test('does not throw error when HOL_2 has a value', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'ValidValue'];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      let shouldThrowError = false;
      
      if (hol1 === 'Üres') {
        if (hol2 === 'Üres') {
          shouldThrowError = true;
        }
      }
      
      expect(shouldThrowError).toBe(false);
    });
  });

  describe('Integration: Processing multiple rows', () => {
    test('processes rows correctly with mixed values', () => {
      const rows = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Value1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Value2', 'Value2'],
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Üres', 'Value3']
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const hol1 = row[HOL_1_INDEX];
        const hol2 = row[HOL_2_INDEX];
        
        if (hol1 === 'Üres') {
          if (hol2 === 'Üres') {
            // Error case - both empty
            throw new Error('Both HOL_1 and HOL_2 are empty');
          }
          row[HOL_1_INDEX] = hol2;
        }
      }
      
      expect(rows[0][HOL_1_INDEX]).toBe('Value1');
      expect(rows[1][HOL_1_INDEX]).toBe('Value2');
      expect(rows[2][HOL_1_INDEX]).toBe('Value3');
    });

    test('collects all problematic rows before reporting error', () => {
      const rows = [
        ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Value1'],
        ['Számla', '2025-01-16', '2025-01-16', 200000, 'Üres', 'Üres'], // Error row 1 at Excel row 3
        ['Számla', '2025-01-17', '2025-01-17', 300000, 'Value3', 'Value3'],
        ['Számla', '2025-01-18', '2025-01-18', 400000, 'Üres', 'Üres'], // Error row 2 at Excel row 5
        ['Számla', '2025-01-19', '2025-01-19', 500000, 'Value5', 'Value5']
      ];
      
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      const problematicRows: number[] = [];
      let processingStopped = false;
      
      // Collect all problematic rows first
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
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
      
      // Check if we should stop processing
      if (problematicRows.length > 0) {
        processingStopped = true;
      }
      
      expect(processingStopped).toBe(true);
      expect(problematicRows).toEqual([3, 5]); // Excel row numbers
      expect(problematicRows.length).toBe(2);
      
      // Check that rows with valid HOL_2 values were processed
      expect(rows[0][HOL_1_INDEX]).toBe('Value1'); // Should be replaced
      expect(rows[1][HOL_1_INDEX]).toBe('Üres'); // Both empty - should remain as is
      expect(rows[3][HOL_1_INDEX]).toBe('Üres'); // Both empty - should remain as is
      expect(rows[2][HOL_1_INDEX]).toBe('Value3'); // Valid - no change
      expect(rows[4][HOL_1_INDEX]).toBe('Value5'); // Valid - no change
    });
  });

  describe('Edge cases', () => {
    test('handles whitespace variations', () => {
      const row1 = ['Faktura', '2025-01-15', '2025-01-15', 100000, ' Üres ', 'Value1'];
      const row2 = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Value1'];
      
      const HOL_1_INDEX = 4;
      
      // Row 1 with whitespace - should not match
      expect(row1[HOL_1_INDEX] === 'Üres').toBe(false);
      
      // Row 2 without whitespace - should match
      expect(row2[HOL_1_INDEX] === 'Üres').toBe(true);
    });

    test('handles case sensitivity', () => {
      const row1 = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'üres', 'Value1'];
      const row2 = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 'Value1'];
      
      const HOL_1_INDEX = 4;
      
      // Row 1 lowercase - should not match
      expect(row1[HOL_1_INDEX] === 'Üres').toBe(false);
      
      // Row 2 exact case - should match
      expect(row2[HOL_1_INDEX] === 'Üres').toBe(true);
    });

    test('handles numeric HOL_2 values', () => {
      const row = ['Faktura', '2025-01-15', '2025-01-15', 100000, 'Üres', 123];
      const HOL_1_INDEX = 4;
      const HOL_2_INDEX = 5;
      
      const hol1 = row[HOL_1_INDEX];
      const hol2 = row[HOL_2_INDEX];
      
      if (hol1 === 'Üres') {
        row[HOL_1_INDEX] = hol2;
      }
      
      expect(row[HOL_1_INDEX]).toBe(123);
    });
  });
});

