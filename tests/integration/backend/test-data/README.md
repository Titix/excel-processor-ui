# Backend Integration Test Data

This directory contains Excel test files used by backend integration tests.

## Test Files

### KiskerHeti Calculator Tests

- **kisker-heti-basic.xlsx**: Basic test with Week 1 and Week 2 data
  - Week 1 (Jan 1-5, truncated): 2 bolt rows (1000, 2000), 1 web row (1500)
  - Week 2 (Jan 6-12): 1 bolt row (3000), 1 web row (2500)
  - Expected: Week 1 totals (bolt: 3000, web: 1500), Week 2 totals (bolt: 3000, web: 2500)
  - Note: Week 1 of 2025 is truncated (January 1-5) because Jan 1, 2025 is a Wednesday

- **kisker-heti-case-insensitive.xlsx**: Tests case-insensitive HOL_1 values
  - Contains: BOLT, Bolt, WEB, Web
  - Expected: All should be grouped correctly

- **kisker-heti-mixed-weeks.xlsx**: Tests filtering of non-bolt/web values
  - Contains: bolt, other (should be ignored), web, bolt
  - Expected: Only bolt and web rows counted

- **kisker-heti-empty-week.xlsx**: Tests skipping rows with empty "Hét Részletesen"
  - Contains: Valid row, empty row, null row
  - Expected: Only valid row counted

- **kisker-heti-decimals.xlsx**: Tests decimal number handling
  - Contains: Decimal values (1000.5, 2000.75, 1500.25)
  - Expected: Correct decimal summation

### HOL Validation Tests

- **hol-validation-replace.xlsx**: Tests HOL_1 replacement when HOL_1 is "Üres"
  - Row 1: HOL_1="Üres", HOL_2="Teszt1" → Should replace
  - Row 2: HOL_1="Üres", HOL_2="Teszt2" → Should replace
  - Row 3: HOL_1="Original", HOL_2="Teszt3" → Should remain unchanged

- **hol-validation-error.xlsx**: Tests error detection when both HOL_1 and HOL_2 are "Üres"
  - Row 1: HOL_1="Üres", HOL_2="Teszt1" → Should replace
  - Row 2: HOL_1="Üres", HOL_2="Üres" → ERROR
  - Row 3: HOL_1="Original", HOL_2="Teszt3" → Should remain unchanged

- **hol-validation-multiple-errors.xlsx**: Tests multiple error rows
  - Contains 3 error rows (rows 2, 4, 6)
  - Expected: All error rows detected and reported

## Regenerating Test Files

To regenerate all test Excel files, run:

```bash
node tests/integration/backend/generate-test-files.js
```

## Modifying Test Data

To modify test data:

1. Edit the data arrays in `generate-test-files.js`
2. Run the generation script to recreate the Excel files
3. Run the integration tests to verify changes

## File Format

All Excel files use the following column structure (OUTPUT_COLUMN_NAMES):

1. Bizonylat fajta
2. Kelte
3. Teljesítés
4. Bruttó érték (HUF)
5. 'Hol' (HOL_1)
6. 'Hol' (HOL_2)
7. Hónap
8. Hét
9. Hét Részletesen

## Notes

- All dates are in JavaScript Date format
- Numeric values can be integers or decimals
- HOL_1 values should be "bolt", "web", or "Üres" for validation tests
- Empty or null values in "Hét Részletesen" will be skipped in kiskerHeti calculations

## Week Structure for 2025

**Important**: Week 1 of 2025 is a truncated week:
- **Week 1**: January 1-5 (Wed-Sun) - Format: "1. Január 1-5"
- **Week 2**: January 6-12 (Mon-Sun) - Format: "2. Január 6-12"
- **Week 3**: January 13-19 (Mon-Sun) - Format: "3. Január 13-19"
- And so on...

This is because January 1, 2025 falls on a Wednesday, making the first week of 2025 only 5 days long (Jan 1-5). Week 2 starts on Monday, January 6.

