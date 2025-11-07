# Excel Column Names Constants

This module defines standardized column names for Excel files used in the application.

## Usage Examples

### Import the constants

```typescript
import { 
  ExcelColumnNames, 
  COLUMN_NAMES, 
  getAllColumnNames,
  isDefinedColumn,
  getColumnKey 
} from './constants';
```

### Access column names

```typescript
// Get a specific column
const columnName = ExcelColumnNames.BIZONYLAT_FAJTA; // 'Bizonylat fajta'

// Get all columns as array
const allColumns = getAllColumnNames();
// Returns: ['Bizonylat fajta', 'Kelte', 'Teljesítés', 'Bruttó érték (HUF)', "'Hol'", "'Hol'"]

// Check if a column is defined
if (isDefinedColumn('Bizonylat fajta')) {
  // Column exists
}

// Get column key from name (reverse lookup)
const key = getColumnKey('Bizonylat fajta'); // Returns: 'BIZONYLAT_FAJTA'
```

### Type-safe usage

```typescript
import { ExcelColumnName } from './constants';

function processColumn(column: ExcelColumnName) {
  // column is type-safe: only accepts the defined column names
}

processColumn(ExcelColumnNames.KELTE); // OK
processColumn('Some other column'); // TypeScript error
```
