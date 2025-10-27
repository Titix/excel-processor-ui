# Test Coverage for Week Date Range Functionality

## Overview
This document summarizes the test coverage for the "Hét Részletesen" (Week Detailed) feature that stores sortable dates in the Excel output.

## Test Files

### 1. Unit Tests (`tests/unit/excelColumns.test.ts`)
**29 tests covering:**
- Excel column name constants validation
- Calculated column constants (Hónap, Hét, Hét Részletesen)
- Week number calculation (`getWeekNumber`)
- Date parsing from various Excel formats (`parseDate`)
- Week start date calculation (`getWeekStartDate`)
- Hungarian week range formatting (`getWeekDateRangeInHungarian`)
- Integration scenarios for week calculation and formatting

**Key test categories:**
- Column name validation
- Week number accuracy for various dates
- Date parsing (Excel serial numbers, ISO strings, Date objects)
- Hungarian month name verification
- Cross-month and year boundary handling
- Sequential date calculations

### 2. Integration Tests (`tests/integration/week-date-range.test.tsx`)
**15 tests covering:**
- Week start date calculation for various weeks
- Sequential date progression (7 days between weeks)
- Chronological vs alphabetical sorting
- Hungarian month name verification
- Date type storage for Excel compatibility
- Edge cases (leap years, year boundaries)
- Week range display format validation

**Key test categories:**
- Date calculation accuracy
- Sorting functionality
- Hungarian localization
- Edge case handling (leap years, year boundaries)
- Excel date type validation

## Test Results
✅ **Total: 44 tests, all passing**
- Unit tests: 29 passing
- Integration tests: 15 passing

## What Is Tested

### Core Functionality
1. **Week Calculation**: Verifies that weeks are calculated correctly from dates
2. **Date Sorting**: Ensures dates are stored as Date objects for proper Excel sorting
3. **Hungarian Formatting**: Tests that week ranges display in Hungarian format
4. **Edge Cases**: Handles leap years, year boundaries, and cross-month weeks

### Excel Compatibility
1. **Date Type Storage**: Dates are stored with type 'd' for proper Excel handling
2. **Chronological Sorting**: Verifies that January comes before August when sorted
3. **Week Start Dates**: All weeks start on consistent days of the week

### Data Integrity
1. **Sequential Progression**: Consecutive weeks are exactly 7 days apart
2. **Year Boundaries**: Correct handling of weeks spanning year boundaries
3. **Month Boundaries**: Proper formatting for weeks spanning multiple months

## Running the Tests

```bash
# Run all tests for week date range functionality
npm test -- --testPathPattern="excelColumns|week-date-range"

# Run unit tests only
npm test -- tests/unit/excelColumns.test.ts

# Run integration tests only
npm test -- tests/integration/week-date-range.test.tsx

# Run all tests
npm test
```

## Test Coverage Summary

### Week Start Date Calculation (`getWeekStartDate`)
✅ Returns correct dates for various weeks
✅ Returns sequential dates (7 days apart) for consecutive weeks
✅ Handles year boundaries correctly
✅ Returns valid dates for all weeks 1-52
✅ Consistent day of week (Mondays for ISO weeks)

### Hungarian Date Range Formatting (`getWeekDateRangeInHungarian`)
✅ Returns Hungarian month names
✅ Handles cross-month weeks (e.g., "December 26-Január 1")
✅ Different years produce different week ranges
✅ Consistent format: "Month day-day" or "Month day-Month day"

### Date Type Storage
✅ Dates stored with type 'd' for Excel compatibility
✅ Proper sorting (chronological, not alphabetical)
✅ All dates are valid Date objects

### Edge Cases
✅ Leap years handled correctly
✅ Year end dates (December) handled properly
✅ Cross-month weeks formatted correctly
✅ Invalid dates return null

## Notes
- All dates are stored as Excel date type ('d') to enable proper sorting
- The "Hét Részletesen" column shows dates in Hungarian format
- Week start dates ensure January comes before August when sorted
- Comprehensive test coverage includes unit and integration tests
- All 44 tests are passing

