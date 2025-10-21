# CSV to DOCX Templater

A powerful multiplatform GUI application for generating multiple DOCX documents from CSV data and DOCX templates. Perfect for creating personalized documents like certificates, letters, invoices, or attestations in bulk.

## ✨ Key Features

### Advanced Field Mapping
- **Multi-Column Priority**: Map up to 5 CSV columns per placeholder with fallback priority
  - Example: Use company_name first, fall back to first_name if empty, then last_name
- **Column Combination**: Merge multiple columns with spaces
  - Example: Combine first_name + last_name → "John Doe"
- **Auto-Matching**: Automatically suggests likely column matches for each placeholder

### Smart Persistence
- **Auto-Save**: Configuration automatically saved per CSV+template combination
- **Unique Configs**: Each file pair gets its own settings (using MD5 hash)
- **OS-Appropriate Storage**:
  - Linux: `~/.config/csvtemplater/`
  - macOS: `~/Library/Application Support/CSVTemplater/`
  - Windows: `%APPDATA%/CSVTemplater/`
- **Reset Button**: Clear saved settings anytime

### User-Friendly Interface
- **Drag & Drop**: Drop CSV or DOCX files anywhere on the window
- **File Type Detection**: Automatically recognizes .csv and .docx extensions
- **Dynamic Field Management**: Add/remove column mappings with +/- buttons
- **Progress Tracking**: Real-time progress bar during document generation
- **Responsive UI**: Threaded generation keeps interface responsive

### Flexible Output
- **Custom Filenames**: Use any CSV field as base filename with prefix/suffix
- **Batch Processing**: Generate hundreds of documents automatically
- **ZIP Archive**: Optional ZIP creation of all generated documents
- **Smart CSV Reading**: Auto-detects encoding (UTF-8, Latin1, Windows-1252) and delimiters

## Installation

### From Release (Recommended)

Download the pre-built binary for your platform from the [Releases](../../releases) page:
- **Windows**: `CSVTemplater-Windows.zip`
- **macOS**: `CSVTemplater-macOS.zip`
- **Linux**: `CSVTemplater-Linux.tar.gz`

Extract and run the application.

### From Source

1. Clone the repository:
```bash
git clone https://github.com/ashajkofci/PythonTemplater.git
cd PythonTemplater
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the GUI application:
```bash
python run_gui.py
```

## Usage

### GUI Application

1. **Select Files** (or drag & drop):
   - Choose your CSV file containing the data
   - Choose your DOCX template file with placeholders (e.g., `{NAME}`, `{AMOUNT}`)
   - Choose an output directory for generated files

2. **Map Fields**:
   - Each template placeholder shows in the mapping section
   - Use dropdowns to select CSV columns (with priority order)
   - Click **"+"** to add fallback columns (up to 5 per field)
   - Check **"Combine"** to merge multiple columns with spaces
   - Auto-matching suggests likely matches

3. **Configure Filenames** (Optional):
   - Select which CSV column to use for the filename
   - Add a custom prefix (e.g., "Invoice_")
   - Add a custom suffix (e.g., "_2024")

4. **Generate**:
   - Click "Generate Documents"
   - Monitor progress in the progress bar
   - Find your generated documents in the output directory

### Command Line Interface (Library Usage)

```python
from templater_core import generate_documents

# Map with priority fallback
field_mapping = {
    '{NAME}': 'company_name',  # First priority
    '{NAME}_fallback': ['first_name', 'last_name'],  # Fallback columns
    '{ADDRESS}': 'street city zip'  # Combine columns with space
}

files, zip_path = generate_documents(
    csv_path='data.csv',
    template_path='template.docx',
    outdir='output/',
    field_mapping=field_mapping,
    filename_prefix='Document_',
    make_zip=True
)
```

## Examples

See `example.py` for a simple usage example.

The repository includes sample files:
- `MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv` - Sample CSV data
- `DONATION-AMIS-JENISCH.docx` - Sample template

## Template Format

Your DOCX template should contain placeholders in curly braces:
- `{NAME}` - Will be replaced with data from the mapped CSV column
- `{AMOUNT}` - Any placeholder name you choose
- `{DATE}` - Custom field names are supported

Placeholders work in:
- Regular text paragraphs
- Tables
- Headers and footers

Example template:
```
Dear {NAME},

Thank you for your payment of {AMOUNT} CHF.

Date: {DATE}
```

## CSV Format

Your CSV file should have:
- Headers in the first row
- One row per document to generate
- Any delimiter (comma, semicolon, tab) - auto-detected
- Any encoding (UTF-8, Latin1, etc.) - auto-detected

Example CSV:
```csv
company_name,first_name,last_name,amount
Acme Corp,,,1500
,John,Doe,500
,Jane,Smith,750
```

## Building from Source

To create standalone executables:

```bash
# Install PyInstaller
pip install pyinstaller

# Build
pyinstaller templater.spec

# Find the executable in the dist/ directory
```

## Development

### Project Structure

- `templater_gui_enhanced.py` - Enhanced GUI application with all features
- `templater_gui.py` - Basic GUI (legacy)
- `templater_core.py` - Core document generation logic
- `run_gui.py` - GUI launcher
- `example.py` - Example library usage
- `test_modules.py` - Basic test suite
- `test_enhanced.py` - Enhanced features test suite
- `templater.spec` - PyInstaller configuration
- `.github/workflows/build-release.yml` - CI/CD configuration

### Running Tests

```bash
# Basic tests
python test_modules.py

# Enhanced features tests
python test_enhanced.py
```

## Requirements

- **Python**: 3.11+
- **Platforms**: Windows 10+, macOS 10.14+, Linux (any distribution)
- **Dependencies**: pandas, python-docx, tkinterdnd2 (see `requirements.txt`)

## License

This project is open source and available for use.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues.

## Releases

Releases are automatically created via GitHub Actions when a version tag is pushed:

```bash
git tag v1.0.0
git push origin v1.0.0
```

This triggers a build for all platforms and creates a release with downloadable installers.
