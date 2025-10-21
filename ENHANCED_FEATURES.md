# Enhanced GUI Features - Visual Guide

## New UI Elements

### 1. Multi-Column Mapping with Priority

```
Template Placeholder: {NAME}
┌─────────────────────────────────────────────────────────┐
│ 1. [company_name        ▼]                              │
│ 2. [first_name          ▼]                              │
│ 3. [last_name           ▼]                              │
│    [+] button to add more (up to 5 columns)             │
│    [-] button to remove columns                         │
│                                                          │
│ [☐] Combine                                             │
└─────────────────────────────────────────────────────────┘
```

**How it works:**
- System tries columns in order: 1st → 2nd → 3rd → 4th → 5th
- Uses first non-empty value found
- Example: If company_name is empty, uses first_name; if that's empty, uses last_name

### 2. Column Combination

```
Template Placeholder: {FULL_NAME}
┌─────────────────────────────────────────────────────────┐
│ 1. [first_name          ▼]                              │
│ 2. [last_name           ▼]                              │
│                                                          │
│ [☑] Combine  ← Check this box                           │
└─────────────────────────────────────────────────────────┘
```

**Result:** Combines values with spaces
- first_name="John", last_name="Doe" → "John Doe"

### 3. Settings Persistence

```
Field Mapping Section Header:
┌─────────────────────────────────────────────────────────┐
│ Field Mapping (Priority: 1st→2nd→...)                   │
│ [Reset Config] button                                   │
│ Tip: Use '+' to add fallback columns, 'Combine' to...  │
└─────────────────────────────────────────────────────────┘
```

**Features:**
- Auto-saves when you change any mapping
- Auto-loads when you select CSV + template
- "Reset Config" clears all saved settings
- Each CSV+template pair has unique config file

### 4. Drag & Drop

```
File Selection Section:
┌─────────────────────────────────────────────────────────┐
│ File Selection (or drag & drop)                         │
│                                                          │
│ CSV File:    [example.csv              ] [Browse...]    │
│ Template:    [template.docx            ] [Browse...]    │
│ Output Dir:  [/path/to/output          ] [Browse...]    │
└─────────────────────────────────────────────────────────┘
```

**Usage:**
- Drag .csv file anywhere → Auto-loads to CSV slot
- Drag .docx file anywhere → Auto-loads to Template slot
- Works alongside Browse buttons

## Complete Workflow Example

### Step 1: Setup Files
Drag and drop your CSV and DOCX files, or use Browse buttons.

### Step 2: Map Fields with Priority
```
{NAME}:
  1. company_name     (try first)
  2. first_name       (fallback if empty)
  3. last_name        (second fallback)
  [☐] Combine

{ADDRESS}:
  1. street
  2. city  
  3. zip
  [☑] Combine         (merge with spaces)

{PHONE}:
  1. mobile_phone
  2. office_phone     (fallback)
  [☐] Combine
```

### Step 3: Configure Output
```
Filename: [full_name     ▼]
Prefix:   [Invoice_        ]
Suffix:   [_2024           ]

Result: Invoice_John_Doe_2024.docx
```

### Step 4: Generate
Click "Generate Documents" and watch progress bar.

## Config Storage Locations

- **Linux**: `~/.config/csvtemplater/abc123def456.json`
- **macOS**: `~/Library/Application Support/CSVTemplater/abc123def456.json`
- **Windows**: `%APPDATA%/CSVTemplater/abc123def456.json`

Where `abc123def456` is MD5 hash of CSV+template paths.

## Benefits

1. **No Re-mapping**: Settings persist between sessions
2. **Flexible Fallbacks**: Handle incomplete data gracefully
3. **Column Merging**: Create full names, addresses, etc.
4. **User Friendly**: Drag files, auto-save, visual feedback
5. **Per-Project**: Different settings for different file pairs

## Example Use Cases

### Use Case 1: Company or Person
```
{RECIPIENT}:
  1. company_name    ← If this exists, use it
  2. first_name      ← Otherwise, try person's first name
  3. last_name       ← Last resort
```

### Use Case 2: Full Address
```
{ADDRESS}:
  1. street
  2. city
  3. state
  4. zip
  [☑] Combine    ← Merge all with spaces
```
Result: "123 Main St New York NY 10001"

### Use Case 3: Phone Hierarchy
```
{CONTACT}:
  1. mobile_phone
  2. office_phone
  3. home_phone
  4. fax_number
```
Uses first available phone number.

## Testing

All features have comprehensive tests in `test_enhanced.py`:
- ✅ Multi-column priority mapping
- ✅ Column combination
- ✅ Settings persistence
- ✅ Template scanning
- ✅ GUI component initialization

Run tests:
```bash
python test_enhanced.py
```
