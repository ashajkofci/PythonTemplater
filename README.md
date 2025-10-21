# CSV to DOCX Templater

A multiplatform GUI application for generating multiple DOCX documents from a CSV file and a DOCX template. Perfect for creating personalized documents like certificates, letters, or attestations in bulk.

## Features

- **GUI Application**: Easy-to-use graphical interface built with tkinter
- **CSV to DOCX**: Process CSV data and fill DOCX templates
- **Field Mapping**: Graphically align template placeholders with CSV columns
- **Custom Filenames**: Configure filename using any CSV field, with custom prefix and suffix
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Batch Processing**: Generate hundreds of documents automatically
- **ZIP Archive**: Optional ZIP creation of all generated documents

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
python templater_gui.py
```

## Usage

### GUI Application

1. **Select Files**:
   - Choose your CSV file containing the data
   - Choose your DOCX template file with placeholders (e.g., `{NAME}`, `{AMOUNT}`)
   - Choose an output directory for generated files

2. **Map Fields**:
   - The application will automatically detect placeholders in your template
   - Map each placeholder to the corresponding CSV column
   - Auto-matching will suggest likely matches

3. **Configure Filenames** (Optional):
   - Select which CSV column to use for the filename
   - Add a custom prefix (e.g., "Invoice_")
   - Add a custom suffix (e.g., "_2024")

4. **Generate**:
   - Click "Generate Documents"
   - Monitor progress in the progress bar
   - Find your generated documents in the output directory

### Command Line Interface

For automation or scripting, use the CLI version:

```bash
python attestation.py \
    --csv "data.csv" \
    --template "template.docx" \
    --date "January 1, 2024" \
    --outdir "output" \
    --zip
```

## Template Format

Your DOCX template should contain placeholders in curly braces:
- `{NAME}` - Will be replaced with data from the mapped CSV column
- `{AMOUNT}` - Any placeholder name you choose
- `{DATE}` - Custom field names are supported

Example template:
```
Dear {NAME},

Thank you for your donation of {AMOUNT} CHF.

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
Name,Amount,Date
John Doe,100,2024-01-15
Jane Smith,250,2024-01-16
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

- `templater_gui.py` - GUI application
- `templater_core.py` - Core document generation logic
- `attestation.py` - CLI interface (legacy compatibility)
- `templater.spec` - PyInstaller configuration
- `.github/workflows/build-release.yml` - CI/CD configuration

### Running Tests

```bash
# Test imports
python -c "import templater_core; print('Core module OK')"

# Test CLI
python attestation.py --help
```

## Examples

The repository includes example files:
- `MJ-FAM -Contacts-MAJ-25-AOUT(Donateurs 2022-2023-2024).csv` - Sample CSV data
- `DONATION-AMIS-JENISCH.docx` - Sample template

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
