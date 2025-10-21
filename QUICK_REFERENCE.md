# Quick Reference

## Installation & Running

### GUI Application
```bash
# From source
python run_gui.py

# Or directly
python templater_gui.py
```

### CLI Application
```bash
python attestation.py --csv data.csv --template template.docx --outdir output/
```

## Common Use Cases

### Generate documents with GUI
1. Select CSV file
2. Select DOCX template  
3. Choose output directory
4. Map fields
5. Click "Generate Documents"

### Generate documents with CLI
```bash
python attestation.py \
    --csv "data.csv" \
    --template "template.docx" \
    --date "January 1, 2024" \
    --outdir "output/" \
    --zip
```

### Use as a Python library
```python
from templater_core import generate_documents

files, zip_path = generate_documents(
    csv_path='data.csv',
    template_path='template.docx',
    outdir='output/',
    field_mapping={'{NAME}': 'customer_name'},
    make_zip=True
)
```

## Template Format

Placeholders use curly braces: `{PLACEHOLDER_NAME}`

Example template:
```
Dear {NAME},

Your donation of {AMOUNT} has been received.

Thank you!
Date: {DATE}
```

## CSV Format

- First row: column headers
- Any delimiter (auto-detected)
- Any encoding (auto-detected)

Example:
```csv
name,amount,date
John Doe,100,2024-01-15
Jane Smith,250,2024-01-16
```

## File Locations

- GUI app: `templater_gui.py`
- CLI app: `attestation.py`
- Core library: `templater_core.py`
- Tests: `test_modules.py`
- Documentation: `README.md`, `USER_GUIDE.md`

## Building Executables

```bash
pip install pyinstaller
pyinstaller templater.spec
```

Executables will be in `dist/` folder.

## Testing

```bash
# Run test suite
python test_modules.py

# Test CLI
python attestation.py --help

# Test core module
python -c "import templater_core; print('OK')"
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Module not found | Run `pip install -r requirements.txt` |
| CSV not reading | Check encoding, try different encodings |
| No files generated | Check CSV has data in mapped columns |
| GUI won't start | tkinter required (installed by default on most systems) |
| Build fails | Ensure PyInstaller is installed |

## Getting Help

- User Guide: See `USER_GUIDE.md`
- Issues: https://github.com/ashajkofci/PythonTemplater/issues
- Code: https://github.com/ashajkofci/PythonTemplater
