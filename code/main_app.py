"""
Professional Multi-Sheet Spreadsheet Editor
Main Entry Point
Version 3.0
"""
import os
import sys
import ttkbootstrap as tb
from table_editor import TableEditor


def resource_path(relative_path):
    """Get resource path for bundled resources."""
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def main():
    # Check for file argument
    initial_file = None
    if len(sys.argv) > 1:
        initial_file = sys.argv[1].strip('"\'')
        if not os.path.exists(initial_file):
            initial_file = None
    
    # Create main window
    root = tb.Window(themename="flatly")
    root.geometry("1400x800")
    
    # Initialize application
    app = TableEditor(root, initial_file)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    
    # Global keyboard shortcuts
    root.bind("<Control-n>", lambda e: app.new_file())
    root.bind("<Control-o>", lambda e: app.open_file())
    root.bind("<Control-s>", lambda e: app.save_file())
    root.bind("<Control-z>", lambda e: app.get_current_sheet().undo() if app.get_current_sheet() else None)
    root.bind("<Control-y>", lambda e: app.get_current_sheet().redo() if app.get_current_sheet() else None)
    
    # Set icon
    try:
        root.iconbitmap(resource_path("csv.ico"))
    except:
        pass
    
    root.mainloop()


if __name__ == "__main__":
    main()
