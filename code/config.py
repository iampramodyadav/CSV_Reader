"""
Configuration and Constants
"""

# Application Metadata
APP_NAME = "CSVEditor"
APP_VERSION = "v3.1"
APP_TITLE = f"{APP_NAME} {APP_VERSION}"

# Default Settings
DEFAULT_THEME = "flatly"
DEFAULT_DARK_THEME = "darkly"
DEFAULT_SEPARATOR = ","

SHEET_LIGHT_THEME = "light green"
SHEET_DARK_THEME  = "dark"

# File Paths
LOG_DIR = "log"
LOG_FILE = "spreadsheet_editor_log.log"
USER_GUIDE_PATH = "docs/user_guide.pdf"

# URLs
HELP_URL = "https://github.com/iampramodyadav/CSV_Reader"
REPORT_EMAIL = "pkyadav01234@gmail.com"
LEFT_STATUS = "© 2025 iampramodyadav"

# Application Dropdown Options
APP_DROPDOWNS = {
    "Turbine_Comp": {
        "Create": {"yes": "Generated inside FE model", "no": "Gets imported"},
        "RotateSetting": {
            '0': 'No rotation',
            '1': ' rotation',
        },

        "Import": {"yes": "Read from a CAD file", "no": "Generated inside the FE model"},

    },
    "Turbine_Conta": {
        "Type": {
            "6": 'Init-Bonded LC-Friction',
            "5": 'No separation contact',
            "4": 'Bonded MPC',
            "3": 'Bonded',
            "2": 'Rough',
            "1": 'Friction',
            "0": 'Inactive'
        },
        "NLHIST": ["Off", "On"],
        
    },
}

# Graph Plotting
GRAPH_TYPE_COLORS = {
    0: "gray",
    1: "purple",
    2: "blue",
    3: "cyan",
    4: "green",
    5: "red",
    6: "orange",
}

FORMULA_HELP_TEXT = """FORMULA REFERENCE GUIDE
    ═════════════════════════════════════════════════════════════

    📊 MATHEMATICAL FUNCTIONS
    ─────────────────────────────────────────────────────────────

    =SUM(A1:A10)              Sum all values in range
    =AVERAGE(A1:A10)          Calculate average of range
    =MAX(A1:A10)              Find maximum value
    =MIN(A1:A10)              Find minimum value
    =COUNT(A1:A10)            Count non-empty cells
    =SQRT(A1)                 Square root
    =POWER(A1,2)              Raise to power (A1²)
    =ABS(A1)                  Absolute value
    =ROUND(A1,2)              Round to 2 decimal places

    🔤 PYTHON EXPRESSIONS
    ─────────────────────────────────────────────────────────────

    =PYTHON(A1*2.5+B1)        Basic arithmetic
    =PYTHON(sqrt(A1))         Square root
    =PYTHON(A1**2)            Power (A1 squared)
    =PYTHON(sin(A1))          Sine function
    =PYTHON(cos(A1))          Cosine function
    =PYTHON(tan(A1))          Tangent function
    =PYTHON(log(A1))          Natural logarithm
    =PYTHON(log10(A1))        Base-10 logarithm
    =PYTHON(exp(A1))          Exponential (e^A1)
    =PYTHON(pi)               Pi constant (3.14159...)
    =PYTHON(e)                Euler's number (2.71828...)

    🔀 LOGICAL FUNCTIONS
    ─────────────────────────────────────────────────────────────

    =IF(A1>10,"High","Low")   Conditional expression
    =COUNTIF(A1:A10,">5")     Count cells matching criteria
    =SUMIF(A1:A10,">5")       Sum cells matching criteria

    📝 TEXT FUNCTIONS
    ─────────────────────────────────────────────────────────────

    =CONCAT(A1,B1)            Join text from cells
    =UPPER(A1)                Convert to uppercase
    =LOWER(A1)                Convert to lowercase
    =LEN(A1)                  Length of text

    ➕ SIMPLE EXPRESSIONS
    ─────────────────────────────────────────────────────────────

    =A1+B1                    Addition
    =A1-B1                    Subtraction
    =A1*B1                    Multiplication
    =A1/B1                    Division

    💡 TIPS
    ─────────────────────────────────────────────────────────────

    • All formulas start with = sign
    • Cell references are case-insensitive
    • Yellow highlight = calculated formula
    • Blue highlight = formula not calculated (manual mode)
    • Press Ctrl+R for autofill
    • Use 'Auto Calc' toggle or 'Calc' button
    """

ABOUT_TEXT = f"""{APP_NAME}
Version: {APP_VERSION}

Spreadsheet editor with:
- Excel multi-sheet support
- Formula engine (SUM, AVERAGE, IF, etc.)
- CSV editing features
- Modern GUI
- Data Cleaning

Developer: Pramod Kumar Yadav (pkyadav01234@gmail.com)
GitHub: https://github.com/iampramodyadav/CSV_Reader
© 2025 iampramodyadav"""
