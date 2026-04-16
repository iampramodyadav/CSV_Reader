"""
User Plugin File — edit freely, changes reflect instantly in the app.
Each function receives: df (current sheet as DataFrame), app (TableEditor instance).
Return a modified df, or None if you just do side effects.

user_plugins.py <-- edit this anytime
        │
        │  loaded fresh on every run click
        v
plugin_manager.py :: run_plugin()
        │
        │  passes df + app reference
        v
your function(df, app)
        │
        ├─ return None       --> no sheet change
        └─ return df_out     --> sheet updates live
"""
import pandas as pd
import os
import sys
from pathlib import Path

TOOL_PATHS = {
    'BEA_VALIDATION': Path(r"K:\Common\Tools\CSVEditor\plugins"),
}
sys.path.extend(str(path) for path in TOOL_PATHS.values())
import conta_catalogue_from_coord as coord
# ─── PLUGIN REGISTRY ──────────────────────────────────────────────────────────
# Add your function here to make it appear in the app menu.
# Format: "Menu Label": function_name
PLUGINS = {}

def register(label):
    """Decorator to register a plugin function."""
    def decorator(fn):
        PLUGINS[label] = fn
        return fn
    return decorator

# ─── YOUR CUSTOM FUNCTIONS ────────────────────────────────────────────────────

@register("Export: Save as TXT summary")
def export_txt_summary(df, app):
    """Example: writes a text file with basic info about the current sheet."""
    if app.current_file:
        out_path = os.path.splitext(app.current_file)[0] + "_summary.txt"
    else:
        out_path = os.path.join(os.path.expanduser("~"), "sheet_summary.txt")
    
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(f"Sheet: {app.current_sheet_name}\n")
        f.write(f"Rows: {len(df)}, Columns: {len(df.columns)}\n\n")
        f.write("Column names:\n")
        for col in df.columns:
            f.write(f"  - {col}\n")
        f.write("\nFirst 5 rows:\n")
        f.write(df.head(5).to_string(index=False))
    
    app.set_status(f"Exported summary → {out_path}")
    return None  # no df modification needed


@register("Transform: Fill blanks with 0")
def fill_blanks_zero(df, app):
    """Replaces all empty cells with 0."""
    df_out = df.replace("", "0")
    app.set_status("Filled all blank cells with 0")
    return df_out  # returning modified df updates the sheet


@register("Transform: Uppercase all value")
def uppercase_all(df, app):
    """Converts all string values to uppercase."""
    df_out = df.map(lambda x: str(x).upper() if pd.notna(x) else x)
    app.set_status("Converted all values to uppercase")
    return df_out
    
@register("Export: Sensor names to TXT")
def export_sensor_names(df, app):
    """Write first column values to a text file."""
    if df.empty:
        app.set_status("Sheet is empty.")
        return None
    
    out = os.path.join(os.path.expanduser("~"), "sensor_names.txt")
    with open(out, "w") as f:
        for val in df.iloc[:, 0]:
            f.write(str(val) + "\n")
    app.set_status(f"Sensor names → {out}")
    return None
    
@register("Write: Tesp CSV Conta from Coord")
def create_conta_csv_temp(df, app):

    #abs_path = os.path.abspath(app.current_file)
    dir_name = os.path.dirname(app.current_file)
    out_file = os.path.join(dir_name, "temp_conta_file.csv")
    
    if 'coordinateTable' in app.current_file:
        coord.create_csv_conta(app.current_file, out_file)
        app.set_status(f"temp_conta_file.csv' written in {out_file}")
    else:
        app.set_status(f"current file is not coordinate file")
    
    return None
