# HOL Validation Test Scenarios

This document describes the test scenarios for HOL_1 and HOL_2 validation logic.

## Validation Rules

1. **Rule 1**: If HOL_1 is "Üres", replace it with HOL_2 value from the same row
2. **Rule 2**: If both HOL_1 and HOL_2 are "Üres", collect all problematic row numbers and show error message, then stop processing

## Test Scenarios

### Scenario 1: Single Error Row
**Input:**
```
Row 2: HOL_1="Üres", HOL_2="Teszt1" → HOL_1 replaced with "Teszt1"
Row 3: HOL_1="Üres", HOL_2="Value2" → HOL_1 replaced with "Value2"
Row 4: HOL_1="Üres", HOL_2="Üres" → ERROR (both empty)
```

**Expected:**
- Error message: "HOL_1 és HOL_2 egyaránt 'Üres'-re van állítva a következő sorokban: 4"
- Processing stops

### Scenario 2: Multiple Error Rows
**Input:**
```
Row 2: HOL_1="Üres", HOL_2="Teszt1" → HOL_1 replaced with "Teszt1"
Row 3: HOL_1="Üres", HOL_2="Üres" → ERROR (both empty)
Row 4: HOL_1="Value3", HOL_2="Value3" → No change
Row 5: HOL_1="Üres", HOL_2="Üres" → ERROR (both empty)
Row 6: HOL_1="Üres", HOL_2="Teszt5" → HOL_1 replaced with "Teszt5"
Row 7: HOL_1="Üres", HOL_2="Üres" → ERROR (both empty)
```

**Expected:**
- Error message: "HOL_1 és HOL_2 egyaránt 'Üres'-re van állítva a következő sorokban: 3, 5, 7"
- Processing stops
- Rows 2 and 6 are processed (HOL_1 replaced with HOL_2)

### Scenario 3: No Errors
**Input:**
```
Row 2: HOL_1="Üres", HOL_2="Teszt1" → HOL_1 replaced with "Teszt1"
Row 3: HOL_1="Üres", HOL_2="Teszt2" → HOL_1 replaced with "Teszt2"
Row 4: HOL_1="Value3", HOL_2="Value3" → No change
```

**Expected:**
- No errors
- All rows processed successfully
- HOL_1 values replaced where applicable

### Scenario 4: All Rows Have Errors
**Input:**
```
Row 2: HOL_1="Üres", HOL_2="Üres" → ERROR
Row 3: HOL_1="Üres", HOL_2="Üres" → ERROR
Row 4: HOL_1="Üres", HOL_2="Üres" → ERROR
```

**Expected:**
- Error message: "HOL_1 és HOL_2 egyaránt 'Üres'-re van állítva a következő sorokban: 2, 3, 4"
- Processing stops

### Scenario 5: Complex Mixed Scenario
**Input:**
```
Row 2: HOL_1="Üres", HOL_2="Value1" → HOL_1 replaced
Row 3: HOL_1="Üres", HOL_2="Üres" → ERROR
Row 4: HOL_1="Value3", HOL_2="Value3" → No change
Row 5: HOL_1="Üres", HOL_2="Üres" → ERROR
Row 6: HOL_1="Üres", HOL_2="Value6" → HOL_1 replaced
Row 7: HOL_1="Üres", HOL_2="Üres" → ERROR
```

**Expected:**
- Error message: "HOL_1 és HOL_2 egyaránt 'Üres'-re van állítva a következő sorokban: 3, 5, 7"
- Processing stops
- Rows 2 and 6 successfully replace HOL_1 with HOL_2
- Rows 3, 5, 7 are problematic
- Row 4 remains unchanged

## Test Coverage

### Unit Tests (`tests/unit/HOL-validation.test.ts`)
- ✅ Replaces HOL_1 with HOL_2 when HOL_1 is "Üres"
- ✅ Keeps HOL_1 unchanged when HOL_1 is not "Üres"
- ✅ Detects when both HOL_1 and HOL_2 are "Üres"
- ✅ Does not throw error when HOL_2 has a value
- ✅ Processes multiple rows correctly with mixed values
- ✅ Stops processing when error is detected
- ✅ Handles whitespace variations
- ✅ Handles case sensitivity
- ✅ Handles numeric HOL_2 values

### Integration Tests (`tests/integration/HOL-validation-integration.test.tsx`)
- ✅ Detects and reports single problematic row
- ✅ Collects all problematic row numbers
- ✅ Formats error message with multiple row numbers
- ✅ Replaces HOL_1 with HOL_2 value when valid
- ✅ Handles mixed valid and invalid rows correctly
- ✅ Processes all rows successfully when no errors exist
- ✅ Reports all rows as problematic when all are errors
- ✅ Formats error message with all row numbers

## Running the Tests

```bash
# Run all HOL validation tests
npm test -- --testPathPattern="HOL"

# Run unit tests only
npm test -- tests/unit/HOL-validation.test.ts

# Run integration tests only
npm test -- tests/integration/HOL-validation-integration.test.tsx
```

## Test Results

✅ **All tests passing**:
- Unit tests: 10 tests
- Integration tests: 9 tests
- **Total: 19 tests passing**

