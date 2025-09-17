#!C:\Users\pramod.kumar\pythonvEnv\myenv12\Scripts\python.exe
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import tksheet
import ttkbootstrap as tb
import webbrowser 
import getpass
from datetime import datetime
import numpy as np

APP_Name    = "MDT CSV Editor"
APP_Version = "v1.0"
APP_Name = f"{APP_Name} {APP_Version}"

class TableEditor:
    def __init__(self, root, initial_file=None):
        self.root = root
        self.root.title(f"{APP_Name} - Untitled")
        self.df = pd.DataFrame()  # Start with empty dataframe
        self.current_file = None
        self.modified = False
        self.dark_mode = False  # Track theme

        # Style
        self.style = tb.Style("flatly")

        # Toolbar frame
        frame = tb.Frame(root)
        frame.pack(fill=tk.X, pady=5)

        # Toolbar buttons
        tb.Button(frame, text="New File", command=self.new_file, bootstyle="secondary").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="Open File", command=self.open_file, bootstyle="primary").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="Save", command=self.save_file, bootstyle="success").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="Save As", command=self.save_file_as, bootstyle="info").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="Undo", command=lambda: self.sheet.undo(), bootstyle="warning").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="Redo", command=lambda: self.sheet.redo(), bootstyle="warning").pack(side=tk.LEFT, padx=5)
        tb.Button(frame, text="MDT Help", command=self.open_help, bootstyle="link").pack(side=tk.RIGHT, padx=5)

        # Rounded dark mode toggle (right side)
        self.dark_mode_var = tk.BooleanVar(value=False)
        tb.Checkbutton(
            frame,
            text="Dark Mode",
            variable=self.dark_mode_var,
            bootstyle="switch",
            command=self.toggle_dark_mode
        ).pack(side=tk.RIGHT, padx=5)
        
        # Suggestion Mode toggle
        self.suggestion_mode_var = tk.BooleanVar(value=False)
        tb.Checkbutton(
            frame,
            text="MDT Suggestion Mode",
            variable=self.suggestion_mode_var,
            bootstyle="switch",
            command=self.toggle_suggestion_mode
        ).pack(side=tk.RIGHT, padx=5)
        
        # Spreadsheet widget
        self.sheet = tksheet.Sheet(root,
                                   show_x_scrollbar=True,
                                   show_y_scrollbar=True,
                                   show_top_left=True,
                                   headers=None)
        
        self.sheet.enable_bindings(("single_select",
                                    "row_select",
                                    "column_select",
                                    "drag_select",
                                    "edit_cell",
                                    "edit_header",
                                    "copy",
                                    "cut",
                                    "paste",
                                    "delete",
                                    "undo",
                                    "redo",
                                    "find",
                                    "replace",
                                    "arrowkeys",
                                    "sort_columns",
                                    "sort_rows",
                                    "rc_insert_row",
                                    "rc_delete_row",
                                    "row_height_resize",
                                    "column_height_resize",
                                    "column_width_resize",
                                    "row_width_resize",
                                    "rc_insert_column",
                                    "rc_delete_column"))
        self.sheet.pack(fill=tk.BOTH, expand=True)

        # Bindings
        self.sheet.extra_bindings([
            ("end_edit_cell", self.mark_modified),
            ("end_edit_header", self.mark_modified),
            ("rc_insert_col", self.add_column),
            ("rc_delete_col", self.delete_column),
            ("rc_insert_row", self.insert_row),
        ])

        # Keyboard shortcuts
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_file_as())
        self.root.bind("<Control-z>", lambda e: self.sheet.undo())
        self.root.bind("<Control-y>", lambda e: self.sheet.redo())

        # Footer / status bar
        self.status_frame = tk.Frame(root, relief=tk.SUNKEN, bd=1)
        self.status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.left_status = tk.Label(self.status_frame, text="© 2025 SGRE", anchor="w")
        self.left_status.pack(side=tk.LEFT, padx=5)

        self.right_status = tk.Label(self.status_frame, text="Ready", anchor="e")
        self.right_status.pack(side=tk.RIGHT, padx=5)

        self.report_label = tk.Label(self.status_frame,text="Report Issue",fg="blue",cursor="hand2",anchor="w")
        self.report_label.pack(side=tk.LEFT, padx=10)
        
        self.report_label.bind("<Button-1>",lambda e: webbrowser.open("mailto:pkyadav01234@gmail.com?subject=MDT CSV Editor Issue&body=Describe the issue here..."))
        # Load initial file if provided
        if initial_file and os.path.exists(initial_file):
            self.load_file(initial_file)
        else:
            # Show initial empty table
            self.show_table()
        
    # ----------------- Web help -----------------   
    def open_help(self):
        webbrowser.open("https://github.com/iampramodyadav/CSV_Reader")
        
    # ----------------- Dark Mode -----------------
    def toggle_dark_mode(self):
        self.dark_mode = self.dark_mode_var.get()
        theme = "darkly" if self.dark_mode else "flatly"
        self.style.theme_use(theme)

        # Colors
        if self.dark_mode:
            bg, fg = "#2e2e2e", "white"
            sel_bg, sel_fg = "#555555", "white"
            header_bg, header_fg = "#444444", "white"
            status_bg, status_fg = "#2e2e2e", "white"
        else:
            bg, fg = "white", "black"
            sel_bg, sel_fg = "#cce5ff", "black"
            header_bg, header_fg = "#f0f0f0", "black"
            status_bg, status_fg = "#f0f0f0", "black"

        # Sheet options
        self.sheet.set_options(
            bg=bg,
            fg=fg,
            selected_bg=sel_bg,
            selected_fg=sel_fg,
            header_bg=header_bg,
            header_fg=header_fg,
            row_index_bg=bg,
            row_index_fg=fg,
            even_bg=bg,
            odd_bg=bg,
            align="center"
        )

        # Update all existing cells
        for r in range(self.sheet.total_rows()):
            for c in range(self.sheet.total_columns()):
                self.sheet.highlight_cells(row=r, column=c, fg=fg, bg=bg)

        # Update footer colors
        self.status_frame.config(bg=status_bg)
        self.left_status.config(bg=status_bg, fg=status_fg)
        self.right_status.config(bg=status_bg, fg=status_fg)
        if self.suggestion_mode_var.get():
            self.customize_turbine_columns()
        self.sheet.refresh()
        
    def clean_all_dropdown(self):
        # Clear all dropdowns (make all cells editable text again)
        for r in range(self.sheet.total_rows()):
            for c in range(self.sheet.total_columns()):
                self.sheet.del_dropdown(r, c)
                
                if self.dark_mode_var.get():
                    bg, fg = "#2e2e2e", "white"
                else:
                    bg, fg = "white", "black"
                self.sheet.highlight_cells(row=r, column=c, bg=bg, fg=fg)
                
    def toggle_suggestion_mode(self):
        if self.suggestion_mode_var.get():
            self.set_status("Suggestion Mode ON")
            self.customize_turbine_columns()
        else:
            # Clear all dropdowns (make all cells editable text again)
            self.clean_all_dropdown()
            # for r in range(self.sheet.total_rows()):
                # for c in range(self.sheet.total_columns()):
                    # self.sheet.del_dropdown(r, c)
                    
                    # if self.dark_mode_var.get():
                        # bg, fg = "#2e2e2e", "white"
                    # else:
                        # bg, fg = "white", "black"
                    # self.sheet.highlight_cells(row=r, column=c, bg=bg, fg=fg)
            
            self.sheet.refresh()
            self.set_status("Suggestion Mode OFF")

    # ----------------- Status -----------------
    def set_status(self, text):
        self.right_status.config(text=text)

    def mark_modified(self, event=None):
        self.modified = True
        self.update_title()
        if self.suggestion_mode_var.get():
            self.customize_turbine_columns()
        self.sheet.refresh()
        self.set_status("Modified")
        
    def update_title(self):
        name = os.path.basename(self.current_file) if self.current_file else "Untitled"
        if self.modified:
            self.root.title(f"{APP_Name} - {name}*")
        else:
            self.root.title(f"{APP_Name} - {name}")

    # ----------------- File Operations -----------------

    def log_usage(self, file_path):
        """Log tool usage with timestamp, user, and file path."""
        try:
            # Customize log location (network/shared path or local)
            log_dir = r"K:\Users\pramod.kumar\shared\tool\log"
            os.makedirs(log_dir, exist_ok=True)  # ensure dir exists
            log_file = os.path.join(log_dir, "mdt_csv_editor_log.log")

            with open(log_file, "a", encoding="utf-8") as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                user = getpass.getuser()
                f.write(f"{timestamp} | {user} | {file_path} | version={APP_Version}\n")
        except Exception:
            # Fail silently if log can't be written
            pass

    def new_file(self):
        if self.modified:  # check unsaved changes
            if not self.ask_save_changes("create a new file"):
                return
        self.df = pd.DataFrame(columns=["Column1", "Column2"])
        self.current_file = None
        self.modified = False
        self.show_table()
        self.update_title()
        self.set_status("New file created")

    def load_file(self, file_path):
        """Load a file without user interaction (for command line usage)"""
        try:
            self.txt_sepa = ' '
            if file_path.endswith(".tsv"):
                self.df = pd.read_csv(file_path, sep=r"\t", engine="python")
            elif file_path.endswith(".txt"):
                # Detect if comma exists in first line
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    first_line = f.readline()
                
                if "," in first_line:
                    self.df = pd.read_csv(file_path, keep_default_na=False, na_values=[])
                    self.txt_sepa = ','
                else:
                    self.df = pd.read_csv(file_path, sep=r"\s+", engine="python")
                    
            else:
                # Default CSV/other
                self.df = pd.read_csv(file_path, keep_default_na=False, na_values=[])

            self.current_file = file_path
            self.modified = False
            self.show_table()
            self.update_title()
            self.set_status(f"Opened {file_path}")
            if self.suggestion_mode_var.get():
                self.customize_turbine_columns()  

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file:\n{e}")
            self.set_status("Error opening file")
            self.show_table()  # fallback to empty table

    def open_file(self):
        if self.modified:  # check unsaved changes
            if not self.ask_save_changes("open a different file"):
                return
                
        file_path = filedialog.askopenfilename(filetypes=[("Data files", "*.csv *.txt *.tsv"), ("All files", "*.*")])
        if not file_path:
            return
        self.load_file(file_path)
        self.log_usage(file_path)

    def save_file(self):
        if self.current_file is None:
            self.save_file_as()
            return
        self.update_dataframe_from_sheet()
        
        df_to_save = self.df.copy()
        df_to_save.replace(r'^\s*$', 'NaN', regex=True, inplace=True)
        
        if self.current_file.endswith(".tsv"):
            df_to_save.to_csv(self.current_file, sep="\t", index=False)
        elif self.current_file.endswith(".txt"):
            df_to_save.to_csv(self.current_file, sep=self.txt_sepa)
        else:
            self.df.to_csv(self.current_file, index=False,na_rep="")
        self.modified = False
        self.update_title()
        self.set_status(f"Saved to {self.current_file}")
        messagebox.showinfo("Saved", f"File saved to {self.current_file}")

    def save_file_as(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV", "*.csv"), ("TSV", "*.tsv"), ("Text", "*.txt")])
        if not file_path:
            return
        self.current_file = file_path
        self.save_file()
        
    def ask_save_changes(self, action="continue"):
        """Ask user if they want to discard unsaved changes."""
        result = messagebox.askyesnocancel("Unsaved Changes",
            f"You have unsaved changes. Do you want to save before {action}?")
        if result:  # Yes → Save
            self.save_file()
            return True
        elif result is None:  # Cancel
            return False
        else:  # No → Discard
            return True
            
    def on_close(self):
        if self.modified:
            if not self.ask_save_changes("exit"):
                return
        self.root.destroy()    
    # ----------------- Table Operations -----------------
    def show_table(self):
        if self.df.empty:
            self.sheet.headers(["Column1", "Column2"])
            self.sheet.set_sheet_data([["", ""]])
        else:
            self.sheet.headers([str(c) for c in self.df.columns])
            self.sheet.set_sheet_data(self.df.astype(str).values.tolist())
        self.toggle_dark_mode()

        self.sheet.refresh()

    def update_dataframe_from_sheet(self):
        self.df = pd.DataFrame(self.sheet.get_sheet_data(), columns=self.sheet.headers())

    def add_column(self, event=None):
        col_index = self.sheet.get_currently_selected()[1] if self.sheet.get_currently_selected() else len(self.df.columns)
        new_col = f"Column{len(self.df.columns) + 1}"
        self.df.insert(col_index, new_col, "")
        self.modified = True
        self.show_table()
        self.update_title()
        self.toggle_dark_mode()
        self.set_status(f"Added column {new_col}")

    def delete_column(self, event=None):
        if not self.df.empty:
            col_index = self.sheet.get_currently_selected()[1]
            if col_index is not None and col_index < len(self.df.columns):
                col_name = self.df.columns[col_index]
                self.df.drop(self.df.columns[col_index], axis=1, inplace=True)
                self.modified = True
                self.show_table()
                self.update_title()
                self.set_status(f"Deleted column {col_name}")


    def set_column_dropdown(self, column_name, options):
        """
        Attach dropdowns to all cells of a column.
        Keeps existing values unchanged.
        """
        if column_name not in self.df.columns:
            return
        col_idx = self.df.columns.get_loc(column_name)

        for r in range(len(self.df)):
            # Get the existing cell value
            existing_value = self.sheet.get_cell_data(r, col_idx)
            
            # Attach the dropdown to the cell
            self.sheet.dropdown(r, col_idx, values=options, state="normal", validate_input = False)
            
            # Set the original value back to the cell
            self.sheet.set_cell_data(r, col_idx, value=existing_value)
            
            if self.dark_mode_var.get():
                bg, fg = "#2e2e2e", "white"
            else:
                bg, fg = "white", "black"
            if existing_value not in options:
                self.sheet.highlight_cells(row=r, column=col_idx, bg="lightcoral", fg=fg)
            else:
                self.sheet.highlight_cells(row=r, column=col_idx, bg=bg, fg=fg)


    def insert_row(self, event=None):
        """
        Insert a new row both in the sheet and in the DataFrame.
        Avoid infinite recursion by disabling event trigger.
        """
        # If called from right-click menu, event is not None
        if event:
            # Just handle DataFrame update (the sheet already added the row)
            idx = self.sheet.get_currently_selected()[0] if self.sheet.get_currently_selected() else len(self.df)
        else:
            # Programmatic insert
            idx = len(self.df)
            self.sheet.insert_row(idx=idx, deselect_all=True, redraw=True)

        # Insert empty row into DataFrame
        empty_row = pd.Series([""] * len(self.df.columns), index=self.df.columns)
        self.df = pd.concat(
            [self.df.iloc[:idx], empty_row.to_frame().T, self.df.iloc[idx:]]
        ).reset_index(drop=True)
        
        self.modified = True
        self.update_title()
        self.set_status("Row inserted")
        self.toggle_dark_mode()
        # Re-apply dropdowns
        if self.suggestion_mode_var.get():
            self.customize_turbine_columns()

    def customize_turbine_columns(self):
        if not self.current_file:
            return
        fname = os.path.basename(self.current_file)
        self.clean_all_dropdown()
        if fname.startswith("VVV"):
                pass
        elif fname.startswith("XXX"):
                pass
        elif fname.startswith("ABC"):
                pass

           
        else:     
            self.suggestion_mode_var.set(False)
            self.toggle_suggestion_mode()
            self.set_status("Suggestion Mode OFF (this file not supported)")
            return

def main():

    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS  # folder PyInstaller uses
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
        
    # Check for command line arguments
    initial_file = None
    if len(sys.argv) > 1:
        # File path passed as command line argument
        initial_file = sys.argv[1]
        # Handle quoted paths and normalize
        initial_file = initial_file.strip('"\'')
        if not os.path.exists(initial_file):
            print(f"Warning: File not found: {initial_file}")
            initial_file = None
    
    root = tb.Window(themename="flatly")
    app = TableEditor(root, initial_file)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    
    icon_path = resource_path("csv.ico")
    # # Try to set icon, but don't fail if it doesn't exist
    try:
        root.iconbitmap(icon_path)
    except:
        pass  # Icon file not found, continue without it
    
    root.mainloop()

if __name__ == "__main__":
    main()
