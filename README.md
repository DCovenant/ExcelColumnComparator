# Excel Column Comparator

Compare cable identifiers and data across multiple Excel files with ease.

## System Requirements

- **Operating System**: Windows 7 or later (Windows 10/11 recommended)
- **Processor**: 1 GHz or faster
- **RAM**: Minimum 2 GB (4 GB recommended)
- **Disk Space**: 200 MB available
- **Display**: 1024x768 or higher resolution
- **Excel**: Not required (works with standalone .xlsx/.xls files)

## Installation

### Using the Installer

1. Download `ExcelColumnComparator-Setup.exe`
2. Double-click the installer
3. Follow the installation wizard
4. Choose installation directory (default: `C:\Program Files\ExcelComparator`)
5. Click "Install"
6. Optionally create a desktop shortcut
7. Launch the application

### Manual Installation (Portable)

1. Download `union_gui.exe`
2. Place in desired location
3. Double-click to run

## Usage Guide

### Basic Workflow

1. **Launch Application**: Click "Excel Column Comparator" from Start Menu or desktop shortcut

2. **Select Files**:
   - Click the file selection dialog
   - Select 2 or more Excel files to compare
   - Click "Open"

3. **Create Template (Optional)**:
   - If comparing multiple similar files, you'll be asked: "Do you want to create a template?"
   - **Yes**: Select header row and columns from first file only (applied to all)
   - **No**: Configure header and columns for each file individually

4. **Configure Each File**:
   - **Select Sheet**: Choose which sheet to use
   - **Select Header Row**: Click the row containing column headers
   - **Select Columns**: Check columns you want to compare

5. **Map Columns**:
   - Map columns from File #1 to other files
   - Same column names auto-match
   - Use "-- skip --" to exclude columns

6. **View Results**:
   - **Common Values**: Cables/identifiers present in all files
   - **Only In [File]**: Unique to specific file
   - Row numbers show Excel row location

7. **Interact with Results**:
   - **Single Click**: Copy value to clipboard
   - **Double Click**: Open actual Excel file showing that row

### Features

- Support for hidden Excel sheets
- Auto-detection of cable/circuit columns
- Template mode for batch processing identical file structures
- Direct Excel file integration
- Row-level tracking for verification

## Troubleshooting

**"File not found" error**:
- Ensure file path doesn't contain special characters
- File must be .xlsx or .xls format
- Close the file in Excel before comparing

**Empty results**:
- Verify header row is correctly selected
- Check that column names match expected pattern
- Ensure data exists below header row

**Slow performance with large files**:
- File size impact: Processing slower with 10,000+ rows
- Close unnecessary applications to free RAM
- Consider splitting very large files

## System Behavior

- Temporary data loaded into memory (files not modified)
- Excel files must be closed before opening via double-click
- Windows file associations used for opening Excel files
- No internet connection required

## Support & Contact

**Documentation**: See this README

**Issue Reporting**: Check file format and system requirements first

**System Information for Troubleshooting**:
- Windows version (Settings → System → About)
- Excel version (if installed)
- File size and row count of problematic files

## Version

Excel Column Comparator v1.0.0

---

**Last Updated**: 2026-02-13
