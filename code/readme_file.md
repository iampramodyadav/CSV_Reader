# Professional Spreadsheet Editor - Installation Guide

## File Structure

Place all files in the same directory:

```
spreadsheet_editor/
├── main.py                    # Entry point
├── config.py                  # Configuration and constants
├── formula_engine.py          # Formula calculation engine
├── ui_components.py           # UI elements (menus, toolbars)
├── mdt_helpers.py            # MDT-specific features
├── table_editor.py           # Main editor class (needs assembly)
└── csv.ico                   # Application icon
```

## Important: Assembling table_editor.py

The `table_editor.py` file is split into 3 parts. You need to combine them:

### Method 1: Manual Assembly

1. Create a new file called `table_editor.py`
2. Copy the content from "table_editor.py (Part 1 of 2)"
3. **Inside the `TableEditor` class**, add all methods from Part 2 and Part 3
4. The methods in Part 2 and Part 3 are defined as standalone functions - add them as methods indented under the class

### Method 2: Structure Guide

```python
# table_editor.py should look like this:

"""
Main Table Editor Class
"""
import statements...

class TableEditor:
    """Main application class."""
    
    def __init__(self, root, initial_file=None):
        # From Part 1
        ...
    
    def _setup_ui(self):
        # From Part 1
        ...
    
    # ... All other methods from Part 1 ...
    
    # Add methods from Part 2 (indented as class methods)
    def load_file(self, file_path: str):
        ...
    
    def save_file(self):
        ...
    
    # ... etc ...
    
    # Add methods from Part 3 (indented as class methods)
    def update_selection_status(self, event=None):
        ...
    
    # ... etc ...
```

## Dependencies

Install required packages:

```bash
pip install pandas tksheet ttkbootstrap pyvis openpyxl
```

## Running the Application

```bash
python main.py
```

Or open a file directly:

```bash
python main.py "path/to/file.csv"
```

## Features

- ✅ Excel multi-sheet support (.xlsx, .xls)
- ✅ CSV/TSV/TXT file support
- ✅ Formula engine (SUM, AVERAGE, IF, PYTHON expressions)
- ✅ Manual/Auto calculation modes
- ✅ MDT suggestion mode with dropdowns
- ✅ Dark/Light themes
- ✅ Autofill patterns
- ✅ Component graph plotting
- ✅ Export with formatting

## Quick Start

1. Combine the three parts of `table_editor.py` into one file
2. Ensure all other files are in the same directory
3. Install dependencies
4. Run `python main.py`

## File Combinations Note

**Parts 2 and 3** contain method definitions that need to be **indented and added as methods** to the `TableEditor` class from Part 1. Remove the standalone `def` declarations and make them class methods by:

1. Indenting all code 4 spaces
2. Ensuring proper `self` parameter
3. Placing them after the Part 1 methods in the class

## Troubleshooting

**Import Errors**: Make sure all files are in the same directory

**Module Not Found**: Install missing dependencies with pip

**TableEditor has no attribute**: Ensure Part 2 and Part 3 methods are added to the class

**Formula Not Working**: Check that `formula_engine.py` is present and imported correctly


"https://ragardner.github.io/tksheet/DOCUMENTATION.html"

