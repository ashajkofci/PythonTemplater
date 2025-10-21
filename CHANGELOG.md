# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- **GUI Application**: Full-featured multiplatform graphical user interface
  - File selection for CSV and DOCX templates
  - Visual field mapping with auto-suggestions
  - Custom filename configuration (prefix, suffix, CSV field selection)
  - Progress tracking with progress bar
  - ZIP archive creation option
  - Threaded document generation for responsive UI
  
- **Core Library** (`templater_core.py`): Reusable document generation engine
  - CSV reading with automatic encoding and delimiter detection
  - Template placeholder extraction
  - Flexible field mapping
  - Custom filename generation
  - Progress callback support
  - ZIP archive creation
  
- **Test Suite** (`test_modules.py`): Comprehensive testing
  - Module import tests
  - CSV reading tests
  - Template parsing tests
  - Column detection tests
  
- **Documentation**:
  - `README.md` - Project overview and quick start
  - `USER_GUIDE.md` - Detailed step-by-step user guide
  - `QUICK_REFERENCE.md` - Quick reference card
  - Code documentation and docstrings
  
- **Build System**:
  - PyInstaller configuration (`templater.spec`)
  - GitHub Actions workflow for automated builds
  - Cross-platform releases (Windows, macOS, Linux)
  
- **Project Infrastructure**:
  - `requirements.txt` - Dependency management
  - `.gitignore` - Version control configuration
  - `run_gui.py` - Simple GUI launcher

### Changed
- **Refactored** `attestation.py`:
  - Now uses `templater_core` module
  - Maintains backward compatibility with existing CLI interface
  - Fixed amount column detection (prioritizes "montant" over "don")
  - Simplified code by extracting common functionality

### Technical Details

**Languages & Frameworks**:
- Python 3.11+
- tkinter for GUI (built-in)
- pandas for CSV processing
- python-docx for Word document manipulation
- PyInstaller for executable creation

**Architecture**:
- Modular design with separation of concerns
- Core logic separated from UI
- CLI and GUI share the same core functionality
- Threaded operations to keep UI responsive

**Compatibility**:
- Windows 10+
- macOS 10.14+
- Linux (any distribution with Python 3.11+)

### Testing

All tests pass successfully:
- ✓ Core module functionality
- ✓ CLI interface
- ✓ GUI module (with tkinter availability check)
- ✓ CSV reading with multiple encodings
- ✓ Template placeholder detection
- ✓ Document generation (106 files from test data)
- ✓ Custom configuration (249 files with prefix/suffix)

## [1.0.0] - Initial Release (Pending)

First official release with GUI and automated builds.
