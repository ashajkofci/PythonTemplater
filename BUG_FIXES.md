# Bug Fixes - Visual Guide

## 1. Remove Button Bug Fix (Commit b68a3f8)

### Before (Bug):
```
Field: {NAME}
  1. [company_name  ▼]
  2. [first_name    ▼] [-]  ← Clicking this...
  3. [last_name     ▼] [-]  
  
Result: Added another field instead of removing!
  1. [company_name  ▼]
  2. [first_name    ▼] [-]
  3. [last_name     ▼] [-]
  4. [              ▼] [-]  ← New field added ❌
```

### After (Fixed):
```
Field: {NAME}
  1. [company_name  ▼]
  2. [first_name    ▼] [-]  ← Clicking this...
  3. [last_name     ▼] [-]  
  
Result: Correctly removes the field!
  1. [company_name  ▼]
  2. [last_name     ▼] [-]  ← Field removed ✅
```

**What was fixed**: 
- The lambda function in the remove button was capturing the wrong index
- Now correctly stores current values, removes the selected one, and recreates with proper indexing

## 2. Dual Filename Fields (Commit b68a3f8)

### Before (Limited):
```
Filename Configuration:
┌─────────────────────────────────┐
│ Name Field:  [customer_name  ▼] │  ← Only 1 field
│ Prefix:      [Invoice_        ] │
│ Suffix:      [_2024           ] │
└─────────────────────────────────┘

Result: Invoice_CustomerName_2024.docx
```

### After (Enhanced):
```
Filename Configuration:
┌─────────────────────────────────┐
│ Name Field 1: [first_name     ▼] │  ← Primary field
│ Name Field 2: [last_name      ▼] │  ← Optional second field
│ Prefix:       [Invoice_        ] │
│ Suffix:       [_2024           ] │
└─────────────────────────────────┘

Result: Invoice_John_Doe_2024.docx
         ↑      ↑    ↑
         prefix field1_field2 suffix
```

**How it works**:
- Select Name Field 1 (required): Primary identifier
- Select Name Field 2 (optional): Additional identifier
- Both fields combined with underscore separator
- If only Field 1 selected: Uses just that field
- If only Field 2 selected: Uses just that field
- If both selected: Combines them with underscore

**Examples**:

| Field 1     | Field 2   | Result Filename       |
|-------------|-----------|----------------------|
| John        | Doe       | John_Doe.docx        |
| Acme Corp   | (empty)   | Acme_Corp.docx       |
| (empty)     | Smith     | Smith.docx           |
| customer_id | region    | C12345_West.docx     |

**Backward Compatibility**:
- Old configs with single "filename_field" still work
- Automatically migrated to "filename_field1" on load

## Test Coverage

New test added: `test_dual_filename_fields()`

```python
# Test with dual filename fields
field_mapping = {
    '{NAME}': 'first_name last_name',
    '{AMOUNT}': 'amount'
}

files, _ = generate_documents(
    ...
    filename_field='first_name last_name',  # Space-separated
    filename_prefix='Invoice_',
    ...
)

# Verifies:
# - 2 files created
# - Filenames contain both first and last names
# - Combined with underscore separator
```

**Test Results**: ✅ All 6 tests passing
- Template Scanning
- Multi-Column Priority
- Column Combination
- **Dual Filename Fields** ← NEW
- Settings Persistence
- Enhanced GUI Import

## Code Changes Summary

**Files Modified**:
1. `templater_gui_enhanced.py`:
   - Fixed `remove_column_selector()` method
   - Added `filename_field2_var` and `filename_field2_combo`
   - Updated save/load config for dual fields
   - Updated generate to combine fields

2. `templater_core.py`:
   - Enhanced filename handling to support space-separated fields
   - Combines fields with underscore for clean filenames

3. `test_enhanced.py`:
   - Added `test_dual_filename_fields()` test case

**Lines Changed**: 
- Added: ~100 lines
- Modified: ~20 lines
- Removed: ~10 lines

## User Impact

**Positive**:
- ✅ Remove button now works correctly
- ✅ More flexible filename generation
- ✅ Better handling of multi-part names
- ✅ Backward compatible with old configs

**No Breaking Changes**:
- Existing configurations still load correctly
- Single field filenames still work as before
- All existing functionality preserved
