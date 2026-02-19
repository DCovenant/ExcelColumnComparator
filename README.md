# Excel Column Comparator

A powerful desktop application for comparing cable identifiers and data across multiple Excel files with an intuitive graphical interface.

## Overview

**Excel Column Comparator** (also known as **Union**) is a Python-based GUI tool designed to help engineers and data analysts quickly identify common values, differences, and discrepancies across multiple Excel spreadsheets. It's particularly useful for:

- Comparing cable lists across multiple project files
- Identifying duplicate identifiers or missing entries
- Cross-referencing data across spreadsheets
- Validating data consistency in bulk
- Quality assurance in technical documentation

The application provides a step-by-step wizard interface that guides users through file selection, column mapping, and displays comprehensive comparison results.

## Key Features

✨ **Multi-file Comparison**
- Compare 2 or more Excel files simultaneously
- Support for both `.xlsx` and `.xls` formats

🔄 **Flexible Column Mapping**
- Select which columns to compare
- Automatically match column names across files
- Manually map columns with different names
- Skip columns that don't need comparison

📋 **Template Mode**
- Define column structure once for multiple similar files
- Batch process files with identical layouts
- Significantly faster when comparing large file sets

📊 **Detailed Results**
- Shows common values found in all files
- Lists values unique to each file
- Displays Excel row numbers for easy reference
- Expandable results with show-more functionality

🖱️ **Convenient Interactions**
- **Single click** on any value to copy to clipboard
- **Double click** (or "Open Excel" button) to jump directly to that row in Excel
- Direct integration with Excel file associations

✅ **Additional Capabilities**
- Support for hidden Excel sheets
- Preserves cell coloring from original files
- Preview of up to 200 rows when selecting header row
- Cross-platform support (Windows, macOS, Linux)
- No Excel installation required (standalone operation)

## System Requirements

### Minimum Requirements
- **OS**: Windows 7+, macOS 10.12+, or Linux (any modern distribution)
- **Processor**: 1 GHz dual-core or faster
- **RAM**: 2 GB minimum (4 GB recommended)
- **Disk Space**: 200 MB free space
- **Display**: 1024×768 minimum resolution

### Recommended
- **OS**: Windows 10/11, macOS Sonoma+, or modern Ubuntu/Debian
- **RAM**: 4+ GB
- **Display**: 1920×1080 or higher

**Note:** Excel installation is NOT required; the application works with standalone Excel files.

## Installation

### For Windows (Executable)

**Option 1: Using the Installer**

1. Download `ExcelColumnComparator-Setup.exe` from releases
2. Double-click to run the installer
3. Follow the installation wizard
4. Choose your installation directory (default: `C:\Program Files\ExcelComparator`)
5. Click "Install"
6. Launch from Start Menu or desktop shortcut

**Option 2: Portable Executable**

1. Download `ExcelColumnComparator.exe`
2. Place it anywhere on your computer
3. Double-click to run (no installation needed)
4. Application runs directly without leaving registry entries

### For Development / Running from Source

**Prerequisites:**
- Python 3.8 or higher
- pip (Python package manager)
- Git (optional, for cloning)

**Installation Steps:**

```bash
# Clone the repository (or download as ZIP)
git clone https://github.com/yourusername/union.git
cd union

# Create a virtual environment (recommended)
python -m venv .venv

# Activate virtual environment
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python test/union_gui.py

# Or run the monolithic version:
python ExelColumnComparator.py
```

## Usage Guide

### Step-by-Step Workflow

#### 1. **Launch the Application**
- Windows: Click "Excel Column Comparator" from Start Menu or desktop shortcut
- From source: Run `python test/union_gui.py` or `python ExelColumnComparator.py`

#### 2. **Select Files**
- Click on the file selection dialog
- Choose 2 or more Excel files (`.xlsx` or `.xls`)
- Click "Open"
- *Tip: You can compare 2 files for a quick comparison or 3+ files to find common values across all*

#### 3. **Choose Template Mode** (Optional)
If comparing 2+ files, you'll be asked:
> "Do you want to create a template?"

- **Yes**: Define column structure once (1st file), apply to all others automatically
- **No**: Configure each file individually (more flexible, takes longer)

Choose **Yes** if your files have identical structures; choose **No** if they differ.

#### 4. **Select Sheet and Header Row**
For each file:
1. Select which sheet to use from the left panel
   - Hidden sheets are marked with `[hidden]`
2. Click the row containing your column headers
3. The row preview appears in the status bar
4. Click "Next →"

#### 5. **Select Columns to Compare**
For each file:
1. Check the boxes for columns you want to compare
2. You must select at least one column
3. Common column names (case-insensitive matching) are usually auto-detected
4. Click "Next →"

#### 6. **Map Columns Across Files**
The final configuration step:
1. Each row shows a column from File #1 (left side)
2. For each other file, choose which column maps to it:
   - Same-named columns auto-match
   - Use `-- skip --` to exclude columns
3. Click "Compare →" to run the analysis

#### 7. **Review Comparison Results**
Results are displayed with:
- **Common Values** (green) - appear in all files
- **Only in [File]** (blue/orange) - unique to specific files
- Row numbers showing exact Excel location
- Summary statistics at the bottom

### Interactive Features

**Copy to Clipboard:**
- Single-click any value in the results
- Value is copied to clipboard
- Confirmation dialog appears

**View in Excel:**
- Click "Open Excel" button next to results section
- Excel file opens automatically with data visible
- Unique values are highlighted in red
- Click a row to navigate to that location

**Show More Items:**
- Results preview up to 20 items
- Click "Show all N items ↓" to expand and see everything

**Start New Comparison:**
- Click "New Comparison" at the bottom
- Returns to file selection step
- Previous selections are cleared

## Project Architecture

The project has two versions:

### Version 1: Monolithic (`ExelColumnComparator.py`)
Single-file implementation with complete functionality. Good for:
- Learning how the system works
- Quick deployment as standalone executable
- Understanding the full workflow in one place

### Version 2: Modular (`test/` directory) - **Recommended**
Refactored with separated concerns:

```
test/
├── app.py                    # Main controller and state management
├── union_gui.py             # Entry point script
├── comparison_engine.py      # Core comparison logic and data models
├── theme.py                 # UI theme configuration
├── widgets.py               # Custom UI widgets
├── screen_sheet.py          # Sheet & header row selection UI
├── screen_columns.py        # Column selection UI
├── screen_mapping.py        # Column mapping UI
├── screen_results.py        # Results display UI
└── test_comparison_engine.py # Unit tests for comparison logic
```

**Key Classes & Data Structures:**

- **`FileConfig`** (dataclass): Represents a file with its sheet, header row, and selected columns
- **`ComparisonResult`** (dataclass): Holds results of comparing two files
- **`App`**: Main application controller, manages workflow and state
- **Screen Classes**: Each handles one step of the workflow

**Core Functions:**

- `compare_value_sets()`: Finds common/unique values between two datasets
- `column_values()`: Extracts and maps column data with row numbers
- `auto_match_columns()`: Intelligently matches columns by name
- `run_comparison()`: Orchestrates the full comparison process

## Development & Contributing

### Running Tests

```bash
cd test
pytest test_comparison_engine.py -v
```

Tests cover:
- Value set comparison logic
- Column extraction and mapping
- Header detection
- Edge cases (empty files, special characters, etc.)

### Code Style

- Python 3.8+ compatible
- Uses tkinter for UI (built-in, cross-platform)
- Data handling via pandas and openpyxl
- Modular screen-based architecture

### Adding Features

The modular version is structured for easy extension:

1. **New comparison logic**: Add functions to `comparison_engine.py`
2. **New UI screen**: Create `screen_newfeature.py`, add to `app.py`
3. **UI improvements**: Modify theme colors in `theme.py`
4. **Custom widgets**: Add to `widgets.py`

### Building Executables

To create `.exe` files for distribution:

```bash
pip install pyinstaller

# Create single-file executable
pyinstaller --onefile --windowed test/union_gui.py

# Output: dist/union_gui.exe
```

See `installer.iss` for Inno Setup configuration to create installer executable.

## Dependencies

- **tkinter**: GUI framework (built-in with Python)
- **pandas**: Data manipulation and Excel reading
- **openpyxl**: Low-level Excel file operations
- **pytest** (dev): Unit testing framework

Full dependency list: See `requirements.txt`

```
pandas==2.3.3
openpyxl==3.1.5
pyinstaller==6.18.0
pytest==9.0.2
```

## Troubleshooting

### "File not found" Error
- Ensure file path doesn't contain special characters
- File must be `.xlsx` or `.xls` format
- Close the file in Excel before comparing
- Try using absolute paths instead of relative paths

### Empty Results
- Verify header row is correctly selected
- Check that column names match expected pattern
- Ensure data exists below the header row
- Try selecting a different header row

### Application Crashes on Large Files
- Processing 10,000+ rows can be slow
- Close unnecessary applications to free RAM
- Consider splitting very large files into smaller chunks
- Monitor system memory usage

### Column Names Not Matching
- Check for leading/trailing whitespace in header cells
- Verify exact spelling (comparison is case-insensitive but spell-sensitive)
- Use manual column mapping if auto-match fails

### Excel File Won't Open from Results
- Ensure Excel is installed on your system
- Close the file in Excel first, then try opening from results
- On Linux/macOS, check if your default spreadsheet application is configured
- Try opening the file manually in Excel/LibreOffice

### GUI Looks Broken/Unresponsive
- Update tkinter: `pip install --upgrade tk`
- On Linux: `sudo apt install python3-tk`
- On macOS: `brew install python-tk@3.x`
- Try resizing the window to force redraw

## System Behavior

- **Memory**: Temporary data loaded into RAM (files not modified)
- **File Access**: Files must be closed before opening via double-click
- **File Associations**: Uses Windows file associations to open Excel files
- **Internet**: No internet connection required; runs offline
- **Performance**:
  - Small files (< 1,000 rows): < 1 second
  - Medium files (1,000 - 10,000 rows): 1-5 seconds
  - Large files (10,000+ rows): 5-30 seconds depending on RAM

## FAQ

**Q: Can I compare more than 2 files?**
A: Yes! Compare 2 or more files. File #1 becomes the reference, and it's compared against all others.

**Q: Does the application modify my Excel files?**
A: No. The application only reads data; it never modifies or saves to your files.

**Q: Can I use this with non-ASCII characters (accents, Unicode, etc.)?**
A: Yes, full Unicode support is included.

**Q: How large can my Excel files be?**
A: Theoretically unlimited, but performance degrades above 50,000 rows. For very large files, consider splitting into chunks.

**Q: Can I compare CSV files instead of Excel?**
A: Not directly, but you can convert CSV to Excel (open in Excel and "Save As" .xlsx) and then use this tool.

**Q: Is there a command-line version?**
A: Not currently, but you can import `comparison_engine.py` in your own Python scripts for programmatic use.

## License

This project is provided as-is for use in engineering and data analysis workflows.

## Credits & Attribution

- **GUI Framework**: tkinter (Python standard library)
- **Data Processing**: pandas, openpyxl
- **Packaging**: pyinstaller
- **Testing**: pytest
- **Color Theme**: Custom light theme designed for readability

## Version History

- **v1.0.0** (2026-02-19) - Initial release
  - Multi-file comparison
  - Template mode
  - Full modular architecture with comprehensive UI
  - Excel integration with row-level tracking

## Support & Contact

For issues, feature requests, or questions:

1. Check the **Troubleshooting** section above
2. Review system requirements and verify your setup
3. Check file format and data structure
4. Open an issue on GitHub with:
   - Your OS version
   - Python version (if running from source)
   - Steps to reproduce
   - Any error messages or screenshots

## Getting Started Quickly

```bash
# Clone and setup
git clone <repository-url>
cd union
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt

# Run
python test/union_gui.py

# Or run tests
pytest test/test_comparison_engine.py -v
```

---

**Ready to compare your Excel files?** Download the latest release or run from source today!

Made with ❤️ for data-driven workflows
