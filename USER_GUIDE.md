# CSV Templater - User Guide

## Quick Start

### Running the Application

#### Using Pre-built Binaries
1. Download the appropriate binary for your platform from the [Releases](../../releases) page
2. Extract the archive
3. Run the executable:
   - **Windows**: Double-click `CSVTemplater.exe`
   - **macOS**: Open `CSVTemplater.app`
   - **Linux**: Run `./CSVTemplater` from terminal

#### Running from Source
```bash
python run_gui.py
```
or
```bash
python templater_gui.py
```

## Step-by-Step Guide

### Step 1: Select Your Files

1. **CSV File**: Click "Browse..." next to "CSV File" and select your data file
   - Supported formats: CSV with any delimiter (comma, semicolon, tab)
   - Encoding: UTF-8, Latin1, or Windows-1252 (auto-detected)
   
2. **Template**: Click "Browse..." next to "Template" and select your DOCX template
   - Must be a .docx file (Word 2007 or later)
   - Should contain placeholders in curly braces like `{NAME}`, `{AMOUNT}`, etc.

3. **Output Directory**: Click "Browse..." next to "Output Dir" and choose where to save generated files

### Step 2: Map Template Fields to CSV Columns

After selecting both CSV and Template files, the "Field Mapping" section will display:

1. **Left Column**: Shows placeholders found in your template (e.g., `{NAME}`, `{AMOUNT}`)
2. **Right Column**: Dropdown menus with your CSV column names

**Auto-Mapping**: The application automatically suggests matches based on similar names. Review and adjust as needed.

**Manual Mapping**: 
- Click any dropdown to select the CSV column that should fill that placeholder
- Leave a dropdown empty if you don't want to map that placeholder
- The placeholder will be left as-is in the output if not mapped

### Step 3: Configure Filenames (Optional)

In the "Filename Configuration" section:

1. **Name Field**: Select which CSV column to use for the filename
   - Example: If you select "Name", and the row has "John Doe", the file will be named based on that
   - If not selected, the first non-empty mapped field will be used

2. **Prefix**: Add text before the filename
   - Example: "Invoice_" → `Invoice_JohnDoe.docx`

3. **Suffix**: Add text after the filename
   - Example: "_2024" → `JohnDoe_2024.docx`

**Complete Example**:
- Name Field: "Customer Name" (contains "John Doe")
- Prefix: "Receipt_"
- Suffix: "_Jan"
- Result: `Receipt_John_Doe_Jan.docx`

### Step 4: Set Options

**Create ZIP archive**: Check this box to create a single ZIP file containing all generated documents.

### Step 5: Generate Documents

1. Click "Generate Documents" button
2. Monitor progress in the progress bar
3. When complete, a message box will show:
   - Number of documents generated
   - Location of ZIP archive (if created)

## Tips and Best Practices

### Template Design

**Good placeholders**:
```
Dear {CUSTOMER_NAME},

Thank you for your payment of {AMOUNT} CHF.

Date: {DATE}
```

**Placeholders work in**:
- Regular text paragraphs
- Tables
- Headers and footers

**Placeholder naming**:
- Use descriptive names: `{FULL_NAME}` not `{N}`
- Use UPPERCASE or Title_Case for clarity
- No spaces inside braces: `{FIRST NAME}` won't work, use `{FIRST_NAME}`

### CSV Preparation

**Column Headers**: 
- First row must contain column names
- Make column names descriptive to help auto-mapping

**Data Quality**:
- Remove empty rows (they'll be skipped)
- Ensure required data is present (e.g., if mapping NAME, that column shouldn't be empty)
- Test with a small subset first (e.g., 5 rows)

**Special Characters**:
- The application handles special characters (é, ñ, ü, etc.)
- File names will be "slugified" (special chars removed/replaced)

### Large Datasets

**Performance**:
- 100 documents: ~30 seconds
- 1,000 documents: ~5 minutes
- 10,000 documents: ~50 minutes

**Memory**:
- Each document needs ~1-2 MB of memory during processing
- For very large datasets (>5,000 rows), close other applications

### Troubleshooting

**"No file selected" or grayed out button**:
- Make sure all three files/directories are selected (CSV, Template, Output)

**"Failed to read CSV"**:
- Check that the file is a valid CSV
- Try opening it in Excel or a text editor
- Make sure it's not open in another program

**"No field mappings configured"**:
- You must map at least one placeholder to a CSV column
- Check the Field Mapping section

**Empty output files**:
- Verify your CSV has data in the mapped columns
- Check that rows aren't empty
- Open one of the generated files to see what was filled

**Filename conflicts**:
- The application automatically adds "_1", "_2", etc. to duplicate filenames
- Consider using a CSV column with unique values for filenames

## Advanced Usage

### Command Line Interface

For automation or batch processing, use the CLI version:

```bash
# Basic usage
python attestation.py \
    --csv data.csv \
    --template template.docx \
    --date "2024-01-15" \
    --outdir output/

# With specific column mapping
python attestation.py \
    --csv data.csv \
    --template template.docx \
    --first-col "Prénom" \
    --last-col "Nom" \
    --amount-col "Montant" \
    --outdir output/ \
    --zip
```

### Integration with Other Tools

The core library can be used in your own Python scripts:

```python
from templater_core import generate_documents

field_mapping = {
    '{NAME}': 'customer_name',
    '{AMOUNT}': 'total_amount',
    '{DATE}': 'invoice_date'
}

files, zip_path = generate_documents(
    csv_path='data.csv',
    template_path='template.docx',
    outdir='output/',
    field_mapping=field_mapping,
    filename_field='customer_name',
    filename_prefix='Invoice_',
    make_zip=True
)
```

## Support

For issues, feature requests, or questions:
- Open an issue on [GitHub](../../issues)
- Check existing issues for solutions
- Include error messages and screenshots when reporting bugs

## Examples

The repository includes example files:
- `MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv` - Sample CSV
- `DONATION-AMIS-JENISCH.docx` - Sample template

Try these files to familiarize yourself with the application!
