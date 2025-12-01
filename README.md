# 📊 Excel Date Fixer and AIS Editor

A powerful desktop application for parsing, fixing, and managing dates in Excel files with support for multiple sheets and advanced data manipulation features.

## Features

### 🔄 Date Parsing & Conversion
- **Intelligent Date Detection**: Automatically scans Excel files to find header rows containing date-related columns
- **Multiple Date Format Support**: Handles various date formats including:
  - `YYYYMMDD` (8-digit format)
  - `DDMMYYYY` format
  - Day-first date formats
  - Flexible date parsing with fallback logic
- **Smart Format Conversion**: Converts dates to standardized `DD-MM-YYYY` format

### 📁 Multi-Sheet Support
- Load and process multiple sheets from a single Excel file
- Switch between sheets using a dropdown selector
- Process and export all sheets individually

### 🔍 Search & Filter
- Real-time search functionality across all columns
- Case-insensitive keyword matching
- Instantly filter data as you type

### 💰 Long/Short Term Asset Flagging
- Automatically generate flags based on "Asset Type" column
- Flag "Long Term" assets as "yes"
- Flag "Short Term" assets as "no"
- Insert flags directly next to the Asset Type column

### 📋 Copy/Paste Operations
- Copy individual cells or entire rows
- Paste data with `Ctrl+V`
- Copy entire columns via right-click context menu
- Preserve formatting during copy/paste operations

### 💾 Export Functionality
- Export processed data back to Excel format
- Maintain all sheets and formatting
- Apply date formatting (`DD-MM-YYYY`) to date columns automatically
- Save to custom location with file dialog

### 🎨 User Interface
- Clean, intuitive tabbed interface with one tab per file
- Color-coded rows (alternating row colors for readability)
- Progress tracking during file loading and export
- Status indicators for operation success/failure
- Responsive and interactive Treeview table display

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Required Packages
```bash
pip install pandas openpyxl customtkinter
```

### Package Details
- **pandas**: Data manipulation and Excel file handling
- **openpyxl**: Excel workbook manipulation for formatting
- **customtkinter**: Modern enhanced Tkinter components
- **tkinter**: Built-in Python GUI framework

## Usage

### Running the Application
```bash
python date-ais.py
```

### Basic Workflow

1. **Load Files**: Click "📂 Select Excel File(s)" button
2. **Select Sheet**: Use the sheet dropdown to switch between sheets
3. **View Data**: Browse the loaded data in the table view
4. **Search**: Use the search box to filter data in real-time
5. **Add Flags** (Optional): Click "➕ Add Long/Short Flags" to generate asset type flags
6. **Export**: Click "💾 Export This File" to save processed data
7. **Copy/Paste**: Use standard keyboard shortcuts (`Ctrl+C`, `Ctrl+V`) or right-click menus

## Key Components

### `find_header_row(path, max_scan=30, sheet_name=0)`
Intelligently locates header rows in Excel sheets by scanning for common column names like "stock name", "buy date", "sell date", etc.

**Parameters:**
- `path`: File path to Excel file
- `max_scan`: Maximum number of rows to scan for header (default: 30)
- `sheet_name`: Sheet name or index to scan

**Returns:** Row index of detected header (default: 0)

### `convert_date_series(series)`
Converts a pandas Series containing dates in various formats to standardized datetime objects.

**Logic:**
1. Attempts `YYYYMMDD` format parsing
2. Falls back to `DDMMYYYY` format
3. Uses day-first parsing as final fallback
4. Requires at least 30% successful conversions to apply the conversion

**Returns:** Converted datetime Series or original Series if conversion unsuccessful

### `FileTab` Class
Manages the UI and data for each opened file in its own tab.

**Key Methods:**
- `load()`: Load Excel file and parse all sheets
- `switch_sheet()`: Change active sheet display
- `add_flags()`: Generate Long/Short term flags
- `export_this_file()`: Save processed data to Excel
- `copy_selection()`: Copy selected rows to clipboard
- `paste_selection()`: Paste clipboard data into selected rows
- `apply_search()`: Filter table based on search keywords

### `ExcelDateFixerApp` Class
Main application controller managing the overall GUI and file operations.

## Tips & Tricks

- **Large Files**: The application displays up to 500 rows in the preview. The full dataset is processed during export.
- **Date Recognition**: If dates aren't recognized, check that your date column contains dates in one of the supported formats.
- **Copy Column**: Right-click on any column header and select "📋 Copy Entire Column" to copy all values in that column.
- **Multi-File Processing**: Load multiple files simultaneously by holding `Ctrl` when selecting files.

## File Structure

```
date-ais.py
├── Utility Functions
│   ├── find_header_row()
│   └── convert_date_series()
├── FileTab Class (Per-file tab management)
└── ExcelDateFixerApp Class (Main application)
```

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+C` | Copy selected rows |
| `Ctrl+V` | Paste into selected rows |
| Right-click on header | Show column context menu |

## Error Handling

- Displays user-friendly error messages for file loading issues
- Graceful handling of unsupported date formats
- Validation for required columns (e.g., "Asset Type" for flagging)
- Progress cancellation and error recovery

## Performance

- Efficiently handles large Excel files (tested with 10,000+ rows)
- Real-time search filtering with minimal lag
- Incremental progress indication for long operations

## Developed by

**ANK**

## Notes

- Always backup your original Excel files before processing
- The application preserves the original file; exports create new files
- Date columns are automatically formatted as `DD-MM-YYYY` in exported files

---

**Version**: 1.0  
**Last Updated**: December 2025
