"""
Configuration and Constants - Generic Version
"""
import os

# Application Metadata
APP_NAME = "CSV-Spreadsheet-Editor"
APP_VERSION = "v1.0.0"
APP_TITLE = f"{APP_NAME} {APP_VERSION}"

# Default Settings
DEFAULT_THEME = "flatly" 
DEFAULT_DARK_THEME = "darkly"
DEFAULT_SEPARATOR = ","

SHEET_LIGHT_THEME = "light green"
SHEET_DARK_THEME = "dark"

# Generic File Paths (Using relative paths or environment variables)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DROPDOWN_DATA_PATH = os.path.join(BASE_DIR, "data", "dropdown_data.json")
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE = "app_editor.log"
USER_GUIDE_PATH = os.path.join(BASE_DIR, "docs", "user_guide.pdf")
PLUGIN_FILE_PATH = os.path.join(BASE_DIR, "plugins", "user_plugins.py")

# URLs & Contact
REPO_URL = "https://github.com/iampramodyadav/CSV_Reader"
HELP_URL = "https://github.com/iampramodyadav/CSV_Reader/wiki"
REPORT_EMAIL = "pkyadav01234@gmail.com"  
LEFT_STATUS = "© 2025 Pramod Kumar Yadav"

# Default Dropdown Options (Generic Categories)
DATA_DROPDOWNS_DEFAULT = {
    "Component_Settings": {
        "Active": {"yes": "Enabled", "no": "Disabled"},
        "Rotation": {
            '0': 'None',
            '1': 'X-Axis',
            '2': 'Y-Axis',
            '3': 'Z-Axis'
        },
        "Type": ['TYPE_A', 'TYPE_B', 'TYPE_C', 'TYPE_D'],
        "Method": ['LINEAR', 'QUADRATIC', 'EXPONENTIAL'],
        "Material": ['Steel', 'Aluminum', 'Plastic', 'Composite'],
    },
    "Interface_Settings": {
        "Type": {
            "0": "Inactive",
            "1": "Friction",
            "2": "Rough",
            "3": "Bonded"
        },
        "Status": ["Off", "On"],
        "Mode": {
            "0": "Default",
            "1": "Penalty",
            "2": "Pure Lagrange"
        }
    }
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
    =SUM(A1:A10)            Sum all values
    =AVERAGE(A1:A10)        Average
    =MAX(A1:A10)            Maximum
    =MIN(A1:A10)            Minimum
    =SQRT(A1)               Square root
    =ROUND(A1,2)            Round to 2 decimal places

    🔤 PYTHON EXPRESSIONS
    ─────────────────────────────────────────────────────────────
    =PYTHON(A1*2.5)         Basic math
    =PYTHON(sqrt(A1))       Square root
    =PYTHON(pi)             Pi constant

    🔀 LOGICAL FUNCTIONS
    ─────────────────────────────────────────────────────────────
    =IF(A1>10,"High","Low") Conditional logic
    =COUNTIF(A1:A10,">5")   Conditional count

    📝 TEXT FUNCTIONS
    ─────────────────────────────────────────────────────────────
    =CONCAT(A1,B1)          Join text
    =UPPER(A1)              To Uppercase
    =LEN(A1)                Length of text
    """

ABOUT_TEXT = f"""{APP_NAME}
Version: {APP_VERSION}

Features:
• Multi-sheet CSV/Excel support
• Built-in Formula engine
• Custom Data Cleaning tools
• Modern UI

Developer: Pramod Kumar Yadav
GitHub: {REPO_URL}
Email: {REPORT_EMAIL}
"""
