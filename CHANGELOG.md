# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added - Enhanced Features (v2.0)

- **Multi-Column Mapping with Priority**:
  - Map up to 5 CSV columns per placeholder
  - Priority fallback system (1st → 2nd → 3rd → 4th → 5th)
  - Dynamic +/- buttons to add/remove columns
  - Uses first non-empty value from priority list

- **Column Combination**:
  - Merge multiple CSV columns with spaces
  - "Combine" checkbox per field
  - Example: first_name + last_name = "John Doe"

- **Settings Persistence**:
  - Auto-save configuration per CSV+template combination
  - Unique config file per file pair (MD5 hash)
  - OS-appropriate storage locations:
    - Linux: `~/.config/csvtemplater/`
    - macOS: `~/Library/Application Support/CSVTemplater/`
    - Windows: `%APPDATA%/CSVTemplater/`
  - "Reset Config" button to clear saved settings
  - Auto-load on file selection

- **Drag & Drop Support**:
  - Drop CSV or DOCX files anywhere on window
  - Auto-detection by file extension
  - Integrated with tkinterdnd2

- **Enhanced Testing**:
  - `test_enhanced.py` with comprehensive tests
  - Tests for multi-column priority
  - Tests for column combination
  - Tests for settings persistence
  - Tests for template scanning
  - All tests passing ✅

### Changed

- **Removed `attestation.py`**: Specific donation attestation CLI removed
- **Added `example.py`**: Generic example showing library usage
- **Enhanced `templater_core.py`**: Support for multi-column priority and combination
- **Updated `run_gui.py`**: Tries enhanced GUI first, falls back to basic
- **Updated `templater.spec`**: Builds enhanced GUI with tkinterdnd2
- **Updated requirements.txt**: Added tkinterdnd2>=0.3.0

### Technical Details

**New Components**:
- `templater_gui_enhanced.py` - Enhanced GUI (650+ lines)
- `FieldMappingRow` class - Manages individual field mappings
- Settings persistence system with JSON storage
- Drag and drop event handling

**Architecture Improvements**:
- Modular field mapping rows
- Clean separation of UI and logic
- Config management utilities
- Enhanced error handling

## [1.0.0] - Initial Release

### Added
- **GUI Application** (`templater_gui.py`):
  - File selection for CSV and DOCX templates
  - Visual field mapping
  - Custom filename configuration
  - Progress tracking
  - ZIP archive creation option

- **Core Library** (`templater_core.py`):
  - Smart CSV reading (auto-detects encoding and delimiter)
  - Template placeholder extraction
  - Flexible field mapping
  - Duplicate filename handling
  - Progress callbacks

- **CLI Interface** (`attestation.py`):
  - Command-line interface for automation
  - Auto-detection of name, amount, civility columns
  - Specific for donation attestations

- **Testing** (`test_modules.py`):
  - Module import tests
  - CSV reading tests
  - Template parsing tests
  - Column detection tests

- **Documentation**:
  - README.md - Project overview
  - USER_GUIDE.md - Detailed guide
  - QUICK_REFERENCE.md - Quick reference
  - CHANGELOG.md - Version history

- **Build System**:
  - PyInstaller configuration
  - GitHub Actions workflow
  - Cross-platform releases (Windows, macOS, Linux)

- **Project Infrastructure**:
  - requirements.txt
  - .gitignore
  - run_gui.py launcher

### Technical Details

**Languages & Frameworks**:
- Python 3.11+
- tkinter for GUI
- pandas for CSV processing
- python-docx for Word documents
- PyInstaller for executables

**Compatibility**:
- Windows 10+
- macOS 10.14+
- Linux (any distribution)
