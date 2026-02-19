# Excel Column Comparator

A powerful desktop application for comparing cable identifiers and data across multiple Excel files with an intuitive graphical interface.

---

## Overview

**Excel Column Comparator** (aka **Union**) is a Python-based GUI tool designed to help engineers and data analysts quickly identify common values, differences, and discrepancies across multiple Excel spreadsheets. Perfect for:

- 🔌 Comparing cable lists across project files
- 🔍 Identifying duplicate identifiers or missing entries
- ✅ Cross-referencing data across spreadsheets
- 📊 Validating data consistency in bulk
- 📋 Quality assurance in technical documentation

The application provides a step-by-step wizard interface that guides users through file selection, column mapping, and displays comprehensive comparison results.

---

## ✨ Key Features

**Multi-file Comparison**
- Compare 2 or more Excel files simultaneously
- Support for both `.xlsx` and `.xls` formats

**Flexible Column Mapping**
- Select which columns to compare
- Automatically match column names across files
- Manually map columns with different names
- Skip columns that don't need comparison

**Template Mode**
- Define column structure once for multiple similar files
- Batch process files with identical layouts
- Significantly faster when comparing large file sets

**Detailed Results**
- Shows common values found in all files
- Lists values unique to each file
- Displays Excel row numbers for easy reference
- Expandable results with show-more functionality

**Convenient Interactions**
- **Single click** on any value to copy to clipboard
- **Double click** (or "Open Excel" button) to jump directly to that row in Excel
- Direct integration with Excel file associations

**Additional Capabilities**
- Support for hidden Excel sheets
- Preserves cell coloring from original files
- Preview of up to 200 rows when selecting header row
- Cross-platform support (Windows, macOS, Linux)
- No Excel installation required (standalone operation)

---

## 📋 System Requirements

### Minimum
- **OS**: Windows 7+, macOS 10.12+, or Linux (any modern distribution)
- **Processor**: 1 GHz dual-core or faster
- **RAM**: 2 GB minimum (4 GB recommended)
- **Disk Space**: 200 MB free space
- **Display**: 1024×768 minimum resolution

### Recommended
- **OS**: Windows 10/11, macOS Sonoma+, or Ubuntu 20.04+
- **RAM**: 4+ GB
- **Display**: 1920×1080 or higher

**Note:** Excel installation is NOT required; the application works with standalone Excel files.

---

## 🚀 Installation

### From Source (Recommended for Development)

**Prerequisites:**
- Python 3.8 or higher
- pip (Python package manager)

**Steps:**

```bash
# Clone the repository
git clone https://github.com/yourusername/union.git
cd union

# Create a virtual environment
python -m venv .venv

# Activate virtual environment
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python ExelColumnComparator.py
```

### Using Executables (Windows)

**Coming Soon:** Pre-built `.exe` files will be available in the [Releases](../../releases) section.

---

## 📖 Usage Guide

### Step-by-Step Workflow

#### 1. Launch the Application
```bash
python ExelColumnComparator.py
```

#### 2. Select Files
- Click the file selection dialog
- Choose 2 or more Excel files (`.xlsx` or `.xls`)
- Click "Open"
- **Tip:** Compare 2 files for quick comparison, or 3+ files to find common values across all

#### 3. Choose Template Mode (Optional for 2+ Files)
You'll be asked: *"Do you want to create a template?"*

- **Yes**: Define column structure once (1st file), apply to all others automatically
- **No**: Configure each file individually (more flexible, takes longer)

Choose **Yes** if your files have identical structures; choose **No** if they differ.

#### 4. Select Sheet and Header Row
For each file:
1. Select which sheet to use from the left panel
2. Click the row containing your column headers
3. The row preview appears in the status bar
4. Click "Next →"

#### 5. Select Columns to Compare
For each file:
1. Check the boxes for columns you want to compare
2. You must select at least one column
3. Common column names (case-insensitive matching) are usually auto-detected
4. Click "Next →"

#### 6. Map Columns Across Files
The final configuration step:
1. Each row shows a column from File #1 (left side)
2. For each other file, choose which column maps to it
3. Use `-- skip --` to exclude columns
4. Click "Compare →" to run the analysis

#### 7. Review Comparison Results
Results are displayed with:
- **Common Values** (green) - appear in all files
- **Only in [File]** (blue/orange) - unique to specific files
- Row numbers showing exact Excel location
- Summary statistics at the bottom

### Interactive Features

**Copy to Clipboard:**
- Single-click any value in the results
- Confirmation dialog appears

**View in Excel:**
- Click "Open Excel" button next to results section
- Excel file opens automatically with data visible
- Unique values are highlighted in red

**Show More Items:**
- Results preview up to 20 items
- Click "Show all N items ↓" to expand and see everything

**Start New Comparison:**
- Click "New Comparison" at the bottom to start over

---

## 🏗️ Project Architecture

The application uses a single-file monolithic architecture for simplicity and easy distribution.

**Main Components:**

- **File Selection**: Multi-file picker with format validation
- **Sheet & Header Detection**: Interactive row selection with data preview
- **Column Mapping**: Intelligent matching with manual override options
- **Comparison Engine**: Fast set operations for finding common/unique values
- **Results Display**: Scrollable, expandable results with Excel integration

**Key Functions:**

- `col_vals()`: Extracts column data with row numbers
- `run_comparison()`: Orchestrates the full comparison process
- `_show_pair_selector()`: Interactive column mapping UI
- `_exp()`: Expandable results display

**Technology Stack:**

- **GUI**: tkinter (Python standard library)
- **Data Processing**: pandas, openpyxl
- **Packaging**: pyinstaller (for executable builds)

---

## 🧪 Development

### Running Tests

Development tests are available in the source repository. To run them:

```bash
pip install pytest
pytest test_comparison_engine.py -v
```

### Code Style

- Python 3.8+ compatible
- Built with tkinter for cross-platform support
- Modular screen-based architecture
- Clear separation of concerns

### Building Executables

To create `.exe` files for Windows distribution:

```bash
pip install pyinstaller

# Create single-file executable
pyinstaller --onefile --windowed ExelColumnComparator.py

# Output: dist/ExelColumnComparator.exe
```

---

## 📦 Dependencies

The application requires minimal dependencies:

```
pandas==2.3.3          # Data manipulation and Excel reading
openpyxl==3.1.5       # Low-level Excel file operations
pyinstaller==6.18.0   # Building standalone executables (optional)
pytest==9.0.2         # Testing framework (development only)
```

See `requirements.txt` for the complete list.

---

## 🔧 Troubleshooting

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

---

## ⚡ Performance

**Memory**: Temporary data loaded into RAM (files not modified)

**File Access**: Files must be closed before opening via double-click

**File Associations**: Uses system file associations to open Excel files

**Internet**: No internet connection required; runs completely offline

**Processing Times:**
- Small files (< 1,000 rows): < 1 second
- Medium files (1,000 - 10,000 rows): 1-5 seconds
- Large files (10,000+ rows): 5-30 seconds depending on RAM

---

## ❓ FAQ

**Q: Can I compare more than 2 files?**
A: Yes! Compare 2 or more files. File #1 becomes the reference, compared against all others.

**Q: Does the application modify my Excel files?**
A: No. The application only reads data; it never modifies or saves to your files.

**Q: Can I use this with non-ASCII characters (accents, Unicode, etc.)?**
A: Yes, full Unicode support is included.

**Q: How large can my Excel files be?**
A: Theoretically unlimited, but performance degrades above 50,000 rows. For very large files, consider splitting into chunks.

**Q: Can I compare CSV files instead of Excel?**
A: Not directly, but you can convert CSV to Excel (open in Excel and "Save As" .xlsx) and then use this tool.

**Q: Is there a command-line version?**
A: Not currently, but you can import the comparison logic in your own Python scripts for programmatic use.

---

## 📄 License

This project is provided as-is for use in engineering and data analysis workflows.

---

## 🙌 Credits

- **GUI Framework**: tkinter (Python standard library)
- **Data Processing**: pandas, openpyxl
- **Packaging**: pyinstaller
- **Testing**: pytest
- **Color Theme**: Custom light theme designed for readability

---

## 📝 Version History

- **v1.0.0** (2026-02-19) - Initial release
  - Multi-file comparison
  - Template mode for batch processing
  - Excel integration with row-level tracking
  - Cross-platform GUI

---

## 💬 Support & Feedback

For issues, feature requests, or questions:

1. Check the **Troubleshooting** section above
2. Review system requirements and verify your setup
3. Check file format and data structure
4. Open an issue on GitHub with:
   - Your OS version
   - Python version (if running from source)
   - Steps to reproduce
   - Any error messages or screenshots

---

## 🚀 Quick Start

```bash
# Clone and setup
git clone <repository-url>
cd union
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt

# Run
python ExelColumnComparator.py
```

---

**Ready to compare your Excel files?** Download and run today!

Made with ❤️ for data-driven workflows
