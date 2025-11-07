/**
 * Unit tests for duplicate detection functionality
 */

describe('Duplicate Detection', () => {
  // Mock data for testing
  const mockHeaders = ['Name', 'Age', 'City'];
  const mockData = [
    ['John', 25, 'New York'],
    ['Jane', 30, 'London'],
    ['John', 25, 'New York'], // Duplicate
    ['Bob', 35, 'Paris'],
    ['Jane', 30, 'London'], // Duplicate
    ['Alice', 28, 'Tokyo']
  ];

  test('should group duplicate rows with all row numbers', () => {
    const mockHeaders = ['Name', 'Age', 'City'];
    const mockDataWithInfo = [
      { row: ['John', 25, 'New York'], originalRowNumber: 2, sheetName: 'Sheet1' },
      { row: ['Jane', 30, 'London'], originalRowNumber: 3, sheetName: 'Sheet1' },
      { row: ['John', 25, 'New York'], originalRowNumber: 4, sheetName: 'Sheet1' }, // Duplicate
      { row: ['Bob', 35, 'Paris'], originalRowNumber: 5, sheetName: 'Sheet1' },
      { row: ['Jane', 30, 'London'], originalRowNumber: 6, sheetName: 'Sheet1' }, // Duplicate
      { row: ['Alice', 28, 'Tokyo'], originalRowNumber: 7, sheetName: 'Sheet1' }
    ];

    const seenRows = new Map<string, { row: any[], rowNumbers: number[], sheetNames: string[] }>();

    mockDataWithInfo.forEach((rowInfo) => {
      const rowString = rowInfo.row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        // This is a duplicate - add row number to existing group
        const existingGroup = seenRows.get(rowString)!;
        existingGroup.rowNumbers.push(rowInfo.originalRowNumber);
        existingGroup.sheetNames.push(rowInfo.sheetName);
      } else {
        // First occurrence - create new group
        seenRows.set(rowString, {
          row: rowInfo.row,
          rowNumbers: [rowInfo.originalRowNumber],
          sheetNames: [rowInfo.sheetName]
        });
      }
    });

    // Convert to array and filter only groups with duplicates
    const duplicateGroups: any[] = [];
    seenRows.forEach((group) => {
      if (group.rowNumbers.length > 1) {
        duplicateGroups.push(group);
      }
    });

    // Should find 2 duplicate groups
    expect(duplicateGroups).toHaveLength(2);
    
    // First group: John, 25, New York (rows 2, 4)
    const johnGroup = duplicateGroups.find(g => g.row[0] === 'John');
    expect(johnGroup).toBeDefined();
    expect(johnGroup.rowNumbers).toEqual([2, 4]);
    expect(johnGroup.row).toEqual(['John', 25, 'New York']);
    
    // Second group: Jane, 30, London (rows 3, 6)
    const janeGroup = duplicateGroups.find(g => g.row[0] === 'Jane');
    expect(janeGroup).toBeDefined();
    expect(janeGroup.rowNumbers).toEqual([3, 6]);
    expect(janeGroup.row).toEqual(['Jane', 30, 'London']);
  });

  test('should handle multiple instances of the same duplicate', () => {
    const mockDataWithInfo = [
      { row: ['John', 25, 'New York'], originalRowNumber: 2, sheetName: 'Sheet1' },
      { row: ['John', 25, 'New York'], originalRowNumber: 4, sheetName: 'Sheet1' }, // Duplicate 1
      { row: ['Jane', 30, 'London'], originalRowNumber: 5, sheetName: 'Sheet1' },
      { row: ['John', 25, 'New York'], originalRowNumber: 7, sheetName: 'Sheet1' }, // Duplicate 2
      { row: ['John', 25, 'New York'], originalRowNumber: 9, sheetName: 'Sheet1' }  // Duplicate 3
    ];

    const seenRows = new Map<string, { row: any[], rowNumbers: number[], sheetNames: string[] }>();

    mockDataWithInfo.forEach((rowInfo) => {
      const rowString = rowInfo.row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        const existingGroup = seenRows.get(rowString)!;
        existingGroup.rowNumbers.push(rowInfo.originalRowNumber);
        existingGroup.sheetNames.push(rowInfo.sheetName);
      } else {
        seenRows.set(rowString, {
          row: rowInfo.row,
          rowNumbers: [rowInfo.originalRowNumber],
          sheetNames: [rowInfo.sheetName]
        });
      }
    });

    const duplicateGroups: any[] = [];
    seenRows.forEach((group) => {
      if (group.rowNumbers.length > 1) {
        duplicateGroups.push(group);
      }
    });

    // Should find 1 duplicate group with 4 instances
    expect(duplicateGroups).toHaveLength(1);
    expect(duplicateGroups[0].rowNumbers).toEqual([2, 4, 7, 9]);
    expect(duplicateGroups[0].row).toEqual(['John', 25, 'New York']);
  });

  test('should create proper headers with grouped row numbers', () => {
    const originalHeaders = ['Name', 'Age', 'City'];
    const newHeaders = ['Duplicate Row Numbers', 'Sheets', ...originalHeaders];
    
    expect(newHeaders).toEqual(['Duplicate Row Numbers', 'Sheets', 'Name', 'Age', 'City']);
  });

  test('should format grouped duplicate rows correctly', () => {
    const duplicateGroup = {
      row: ['John', 25, 'New York'],
      rowNumbers: [2, 4, 7],
      sheetNames: ['Sheet1', 'Sheet1', 'Sheet1']
    };
    
    const rowNumbersStr = duplicateGroup.rowNumbers.join(', ');
    const uniqueSheets = Array.from(new Set(duplicateGroup.sheetNames));
    const sheetsStr = uniqueSheets.join(', ');
    const formattedRow = [rowNumbersStr, sheetsStr, ...duplicateGroup.row];
    
    expect(formattedRow).toEqual(['2, 4, 7', 'Sheet1', 'John', 25, 'New York']);
  });

  test('should handle empty data', () => {
    const seenRows = new Set<string>();
    const duplicateRows: any[][] = [];
    const duplicateRowSignatures = new Set<string>();

    const emptyData: any[][] = [];

    emptyData.forEach((row) => {
      const rowString = row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        if (!duplicateRowSignatures.has(rowString)) {
          duplicateRowSignatures.add(rowString);
          duplicateRows.push(row);
        }
      } else {
        seenRows.add(rowString);
      }
    });

    expect(duplicateRows).toHaveLength(0);
  });

  test('should handle data with no duplicates', () => {
    const seenRows = new Set<string>();
    const duplicateRows: any[][] = [];
    const duplicateRowSignatures = new Set<string>();

    const uniqueData = [
      ['John', 25, 'New York'],
      ['Jane', 30, 'London'],
      ['Bob', 35, 'Paris'],
      ['Alice', 28, 'Tokyo']
    ];

    uniqueData.forEach((row) => {
      const rowString = row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        if (!duplicateRowSignatures.has(rowString)) {
          duplicateRowSignatures.add(rowString);
          duplicateRows.push(row);
        }
      } else {
        seenRows.add(rowString);
      }
    });

    expect(duplicateRows).toHaveLength(0);
  });

  test('should handle null and undefined values', () => {
    const seenRows = new Set<string>();
    const duplicateRows: any[][] = [];
    const duplicateRowSignatures = new Set<string>();

    const dataWithNulls = [
      ['John', 25, 'New York'],
      ['Jane', null, 'London'],
      ['John', 25, 'New York'], // Duplicate
      ['Jane', null, 'London'], // Duplicate
      [null, 35, 'Paris']
    ];

    dataWithNulls.forEach((row) => {
      const rowString = row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        if (!duplicateRowSignatures.has(rowString)) {
          duplicateRowSignatures.add(rowString);
          duplicateRows.push(row);
        }
      } else {
        seenRows.add(rowString);
      }
    });

    expect(duplicateRows).toHaveLength(2);
    expect(duplicateRows[0]).toEqual(['John', 25, 'New York']);
    expect(duplicateRows[1]).toEqual(['Jane', null, 'London']);
  });

  test('should handle whitespace differences', () => {
    const seenRows = new Set<string>();
    const duplicateRows: any[][] = [];
    const duplicateRowSignatures = new Set<string>();

    const dataWithWhitespace = [
      ['John', 25, 'New York'],
      ['John ', 25, ' New York'], // Same data with extra spaces
      ['Jane', 30, 'London']
    ];

    dataWithWhitespace.forEach((row) => {
      const rowString = row.map(cell => 
        cell === null || cell === undefined ? '' : String(cell).trim()
      ).join('|');
      
      if (seenRows.has(rowString)) {
        if (!duplicateRowSignatures.has(rowString)) {
          duplicateRowSignatures.add(rowString);
          duplicateRows.push(row);
        }
      } else {
        seenRows.add(rowString);
      }
    });

    expect(duplicateRows).toHaveLength(1);
    expect(duplicateRows[0]).toEqual(['John ', 25, ' New York']);
  });
});

// Export empty object to make this a module
export {};
