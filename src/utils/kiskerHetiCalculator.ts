/**
 * Utility function to calculate KiskerHeti data from pivot excel data
 * Groups by "Hét Részletesen" and sums "Bruttó érték (HUF)" based on HOL_1 values
 */

export interface KiskerHetiRow {
  hetReszletesen: string;
  bolt: number;
  web: number;
  grandTotal: number;
}

/**
 * Calculates KiskerHeti data from pivot excel rows
 * @param finalHeaders - Array of header names from pivot excel (OUTPUT_COLUMN_NAMES)
 * @param finalData - Array of data rows (first row is headers)
 * @returns Array of KiskerHeti rows grouped by "Hét Részletesen"
 */
export function calculateKiskerHeti(
  finalHeaders: readonly string[] | string[],
  finalData: any[][]
): KiskerHetiRow[] {
  // Find column indices dynamically
  // OUTPUT_COLUMN_NAMES order: [Bizonylat fajta, Kelte, Teljesítés, Bruttó érték (HUF), HOL_1, HOL_2, Hónap, Hét, Hét Részletesen]
  const hetReszletesenIndex = finalHeaders.indexOf('Hét Részletesen');
  const bruttoErtekIndex = finalHeaders.indexOf('Bruttó érték (HUF)');
  
  // Find HOL_1 - it's the first occurrence of "'Hol'" in OUTPUT_COLUMN_NAMES
  // Since both HOL_1 and HOL_2 have the same name "'Hol'", we need to find the first one
  // In OUTPUT_COLUMN_NAMES, HOL_1 is at index 4 and HOL_2 is at index 5
  // But to be safe, let's verify by checking the actual position
  let hol1Index = -1;
  const holColumnName = "'Hol'";
  // Find all indices where "'Hol'" appears
  const holIndices: number[] = [];
  for (let i = 0; i < finalHeaders.length; i++) {
    if (String(finalHeaders[i]).trim() === holColumnName) {
      holIndices.push(i);
    }
  }
  // HOL_1 should be the first occurrence
  if (holIndices.length > 0) {
    hol1Index = holIndices[0];
  } else {
    // Fallback: assume it's at index 4 if not found
    hol1Index = 4;
  }

  // Validate indices
  if (hetReszletesenIndex === -1) {
    throw new Error('Column "Hét Részletesen" not found in headers');
  }
  if (bruttoErtekIndex === -1) {
    throw new Error('Column "Bruttó érték (HUF)" not found in headers');
  }
  if (hol1Index >= finalHeaders.length) {
    throw new Error('HOL_1 column index out of bounds');
  }

  // Group data by "Hét Részletesen" and calculate sums
  // For each unique "Hét Részletesen" value, we create ONE row with:
  // - bolt: sum of all "Bruttó érték (HUF)" where HOL_1 = "bolt"
  // - web: sum of all "Bruttó érték (HUF)" where HOL_1 = "web"
  // - grandTotal: bolt + web
  const groupedData = new Map<string, { bolt: number; web: number }>();

  // Skip header row (index 0) and process data rows
  console.log('calculateKiskerHeti: Processing', finalData.length - 1, 'data rows');
  console.log('Column indices: hetReszletesenIndex=', hetReszletesenIndex, ', bruttoErtekIndex=', bruttoErtekIndex, ', hol1Index=', hol1Index);
  
  for (let i = 1; i < finalData.length; i++) {
    const row = finalData[i];
    const hetReszletesen = row[hetReszletesenIndex];
    const bruttoErtekValue = row[bruttoErtekIndex];
    const hol1 = String(row[hol1Index] || '').trim();
    
    // Parse numeric value
    // Handle numbers with thousands separators (commas) and decimal points
    // Example: "6,800.33" should parse to 6800.33, "13,899.88" to 13899.88
    let bruttoErtek = 0;
    if (bruttoErtekValue !== null && bruttoErtekValue !== undefined && bruttoErtekValue !== '') {
      // Convert to string and remove thousands separators (commas)
      // Format is: "number,number.number" where comma is thousands separator
      let strValue = String(bruttoErtekValue).trim();
      
      // Remove commas (thousands separators) before parsing
      // This handles formats like "6,800.33", "13,899.88", or "-12,300.04"
      strValue = strValue.replace(/,/g, '');
      
      const parsed = parseFloat(strValue);
      bruttoErtek = isNaN(parsed) ? 0 : parsed;
      if (i <= 3) { // Log first 3 rows for debugging
        console.log(`Row ${i}: hetReszletesen="${hetReszletesen}", bruttoErtek=${bruttoErtekValue}->${bruttoErtek}, hol1="${hol1}"`);
      }
    }

    // Skip rows without week info
    if (!hetReszletesen || hetReszletesen === null || hetReszletesen === undefined) {
      continue;
    }

    // Use the week value as the grouping key
    const key = String(hetReszletesen).trim();

    // Initialize group if it doesn't exist
    if (!groupedData.has(key)) {
      groupedData.set(key, { bolt: 0, web: 0 });
    }

    const group = groupedData.get(key)!;

    // Case-insensitive comparison for HOL_1 values and accumulate sums
    const hol1Lower = hol1.toLowerCase();
    if (hol1Lower === 'bolt') {
      group.bolt += bruttoErtek;
    } else if (hol1Lower === 'web') {
      group.web += bruttoErtek;
    }
    // Note: If HOL_1 is neither "bolt" nor "web", the row is ignored (not added to any sum)
  }

  // Convert to array and sort by week number
  const result: KiskerHetiRow[] = [];

  // Sort by week (extract week number from "Hét Részletesen" for sorting)
  const sortedKeys = Array.from(groupedData.keys()).sort((a, b) => {
    // Extract week number from format "1. Január 1-7" or similar
    const weekA = parseInt(String(a).split('.')[0]) || 0;
    const weekB = parseInt(String(b).split('.')[0]) || 0;
    return weekA - weekB;
  });

  sortedKeys.forEach(key => {
    const group = groupedData.get(key)!;
    const grandTotal = group.bolt + group.web;
    result.push({
      hetReszletesen: key,
      bolt: group.bolt,
      web: group.web,
      grandTotal: grandTotal
    });
  });

  return result;
}
