# CSV to DOCX Templater

A powerful multiplatform GUI application for generating multiple DOCX documents from CSV data and DOCX templates. Perfect for creating personalized documents like certificates, letters, invoices, or attestations in bulk.

**Author**: Adrian Shajkofci  
**License**: BSD-3-Clause (see [LICENSE](LICENSE))  
**Version**: 2.0

---

## ✨ Features

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
- **Help Menu**: About and License information accessible from menu bar

### Flexible Output
- **Dual Filename Fields**: Use up to 2 CSV fields combined for filenames
- **Custom Prefixes/Suffixes**: Add text before and after filenames
- **Batch Processing**: Generate hundreds of documents automatically
- **ZIP Archive**: Optional ZIP creation of all generated documents
- **Smart CSV Reading**: Auto-detects encoding (UTF-8, Latin1, Windows-1252) and delimiters

---

## 📦 Installation

### Option 1: From Release (Recommended)

Download the pre-built binary for your platform from the [Releases](../../releases) page:
- **Windows**: `CSVTemplater-Windows.zip`
- **macOS**: `CSVTemplater-macOS.zip`
- **Linux**: `CSVTemplater-Linux.tar.gz`

Extract and run the application.

### Option 2: From Source

1. **Clone the repository**:
```bash
git clone https://github.com/ashajkofci/PythonTemplater.git
cd PythonTemplater
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

3. **Run the GUI application**:
```bash
python run_gui.py
```

---

## 🚀 Quick Start

### Using the GUI

1. **Launch the application**:
   ```bash
   python run_gui.py
   ```

2. **Select or drag & drop files**:
   - CSV file with your data
   - DOCX template with placeholders like `{NAME}`, `{AMOUNT}`, etc.
   - Output directory for generated files

3. **Map fields**:
   - Each template placeholder appears with dropdown menus
   - Select CSV columns for each placeholder
   - Use **+** button to add fallback columns (up to 5)
   - Check **Combine** to merge multiple columns

4. **Configure filename** (optional):
   - Select up to 2 CSV fields for the filename
   - Add custom prefix (e.g., "Invoice_")
   - Add custom suffix (e.g., "_2024")

5. **Generate**:
   - Click "Generate Documents"
   - Monitor progress
   - Find documents in output directory

### Using as a Library

```python
from templater_core import generate_documents

# Map with priority fallback and combination
field_mapping = {
    '{NAME}': 'company_name',  # First priority
    '{NAME}_fallback': ['first_name', 'last_name'],  # Fallbacks
    '{ADDRESS}': 'street city zip'  # Combine columns with space
}

files, zip_path = generate_documents(
    csv_path='data.csv',
    template_path='template.docx',
    outdir='output/',
    field_mapping=field_mapping,
    filename_field='first_name last_name',  # Combine for filename
    filename_prefix='Document_',
    filename_suffix='_2024',
    make_zip=True
)

print(f'Generated {len(files)} documents')
```

---

## 📝 Template Format

Your DOCX template should contain placeholders in curly braces:

```
Dear {NAME},

Thank you for your payment of {AMOUNT} CHF.

Date: {DATE}
Best regards,
{SENDER}
```

**Placeholders work in**:
- Regular text paragraphs
- Tables
- Headers and footers

**Note**: The application handles placeholders even when they're split across text runs in the DOCX file (a common issue in Word documents).

---

## 📊 CSV Format

Your CSV file should have:
- **Headers in the first row**: Column names
- **One row per document**: Each row generates one output document
- **Any delimiter**: Comma, semicolon, tab (auto-detected)
- **Any encoding**: UTF-8, Latin1, Windows-1252 (auto-detected)

**Example CSV**:
```csv
company_name,first_name,last_name,amount,date
Acme Corp,,,1500,2024-01-15
,John,Doe,500,2024-01-16
,Jane,Smith,750,2024-01-17
```

---

## 🎯 Advanced Features

### Multi-Column Priority Mapping

Map one placeholder to multiple CSV columns with fallback:

```
{RECIPIENT}:
  1. company_name    ← Try this first
  2. first_name      ← If empty, try this
  3. last_name       ← If still empty, try this
  Result: Uses first non-empty value
```

### Column Combination

Merge multiple columns into one field:

```
{FULL_NAME}:
  Columns: first_name, last_name
  [✓] Combine
  Result: "John Doe"
```

### Dual Filename Fields

Combine two CSV fields for the output filename:

```
Name Field 1: customer_id
Name Field 2: last_name
Prefix: Invoice_
Suffix: _2024

Result: Invoice_12345_Smith_2024.docx
```

### Settings Persistence

- Configuration auto-saves when you make changes
- Auto-loads when you select the same CSV+template pair
- Each combination has its own unique config file
- Use "Reset Config" button to clear and start fresh

---

## 🔧 Building from Source

To create standalone executables:

```bash
# Install PyInstaller
pip install pyinstaller

# Build
pyinstaller templater.spec

# Find executable in dist/ directory
```

---

## 🧪 Testing

Run the test suites to verify functionality:

```bash
# Basic tests
python test_modules.py

# Enhanced features tests
python test_enhanced.py
```

All tests should pass with ✅ indicators.

---

## 📂 Project Structure

```
PythonTemplater/
├── templater_gui_enhanced.py    # Enhanced GUI (main application)
├── templater_core.py             # Core document generation library
├── run_gui.py                    # GUI launcher script
├── example.py                    # Example usage script
├── test_modules.py               # Basic test suite
├── test_enhanced.py              # Enhanced features tests
├── templater.spec                # PyInstaller configuration
├── requirements.txt              # Python dependencies
├── LICENSE                       # BSD-3-Clause license
├── README.md                     # This file
├── CHANGELOG.md                  # Version history
└── .github/workflows/            # CI/CD configuration
    └── build-release.yml
```

**Deprecated files** (old GUI, kept for compatibility):
- `templater_gui.py` - Basic GUI (superseded by enhanced version)

---

## 🐛 Known Issues & Solutions

### Issue: Placeholders not replaced

**Solution**: This has been fixed in version 2.0. The application now handles placeholders that are split across multiple text runs in the DOCX file (a common issue when editing templates in Word).

### Issue: Remove button adds fields

**Solution**: Fixed in commit b68a3f8. The remove (-) button now correctly removes field mappings.

---

## 💡 Tips & Best Practices

### Template Design
- Use descriptive placeholder names: `{CUSTOMER_NAME}` not `{N}`
- Use UPPERCASE or Title_Case for clarity
- No spaces inside braces: `{FIRST_NAME}` not `{FIRST NAME}`
- Test with a small CSV subset first (e.g., 5 rows)

### CSV Preparation
- Make column names descriptive to help auto-matching
- Remove empty rows (they'll be skipped automatically)
- Ensure required data is present in the columns you map

### Performance
- 100 documents: ~30 seconds
- 1,000 documents: ~5 minutes
- 10,000 documents: ~50 minutes

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues.

### Development Setup

1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Make your changes
4. Run tests: `python test_modules.py && python test_enhanced.py`
5. Submit a pull request

---

## 📜 License

This project is licensed under the BSD 3-Clause License. See the [LICENSE](LICENSE) file for details.

Copyright (c) 2025, Adrian Shajkofci

---

## 🚀 Creating a Release

To create a release with installers for all platforms:

```bash
git tag v2.0.0
git push origin v2.0.0
```

This triggers the GitHub Actions workflow to:
1. Run tests on Ubuntu
2. Build executables for Windows, macOS, and Linux
3. Create a release with downloadable binaries

---

## 📞 Support

For issues, feature requests, or questions:
- Open an issue on [GitHub](../../issues)
- Check existing issues for solutions
- Include error messages and screenshots when reporting bugs

---

## 🎉 Acknowledgments

Special thanks to all contributors and users who provided feedback to make this application better.

---

**Quick Links**:
- [Installation](#-installation)
- [Quick Start](#-quick-start)
- [Template Format](#-template-format)
- [Advanced Features](#-advanced-features)
- [License](LICENSE)
- [Changelog](CHANGELOG.md)
