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
    initial_file = None
    if len(sys.argv) > 1:
        initial_file = sys.argv[1].strip('"\'')
        if not os.path.exists(initial_file):
            initial_file = None
    
    root = tb.Window(themename="flatly")
    root.geometry("1400x800")
    
    app = TableEditor(root, initial_file)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    
    # Global keyboard shortcuts
    root.bind("<Control-n>", lambda e: app.new_file())
    root.bind("<Control-o>", lambda e: app.open_file())
    root.bind("<Control-s>", lambda e: app.save_file())
    # root.bind("<Control-z>", lambda e: app.get_current_sheet().undo() if app.get_current_sheet() else None)
    # root.bind("<Control-y>", lambda e: app.get_current_sheet().redo() if app.get_current_sheet() else None)
    root.bind("<Control-z>", lambda e: app.undo())
    root.bind("<Control-y>", lambda e: app.redo())
    
    root.bind("<Control-Shift-O>", lambda e: app.add_file())
    root.bind("<Control-d>", lambda e: app.open_diff_dialog())
    root.bind("<Control-Shift-V>",   lambda e: app.toggle_split_view())
    
    root.bind("<Control-p>",         lambda e: app.open_command_palette())
    root.bind("<Control-w>",         lambda e: app.close_current_tab())
    
    root.bind("<Control-Next>",       lambda e: app.next_tab())
    root.bind("<Control-Prior>", lambda e: app.prev_tab())
    
    root.bind("<Control-s>",         lambda e: app._save_file_silent())  # silent save
    root.bind("<Control-Shift-S>",   lambda e: app.save_all())           # save all

    
    try:
        root.iconbitmap(resource_path("csv.ico"))
    except:
        pass
    app.open_recent_on_startup(n=5)
    root.mainloop()


if __name__ == "__main__":
    main()
