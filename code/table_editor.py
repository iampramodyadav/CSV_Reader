"""
Main Table Editor Class - Part 1
Initialization, UI Setup, Sheet Management
"""
import os
import re
from datetime import datetime
from typing import Dict, Optional
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import tksheet
import ttkbootstrap as tb
from formula_engine import FormulaEngine
from ui_components import UIComponents
from custom_helpers import plot_interactive_component_graph_coordinate, get_mdt_dropdown_config, get_file_list_for_column, plot_interactive_component_graph, trim_whitespace, read_label_format, write_label_format,fill_nulls
from config import APP_TITLE, DEFAULT_THEME,DEFAULT_DARK_THEME, LOG_DIR, LOG_FILE, APP_VERSION, FORMULA_HELP_TEXT,SHEET_LIGHT_THEME, SHEET_DARK_THEME



class TableEditor:
    """Main application class for spreadsheet editing."""

    def __init__(self, root, initial_file=None):
        self.root = root
        self.root.title(f"{APP_TITLE} - Untitled")
        
        # Core data
        self.df = pd.DataFrame()
        self.current_file = None
        self.modified = False
        self.dark_mode = False
        self.current_sep = None
        self.txt_sepa = ','
        
        # Multi-sheet support
        self.workbook_sheets: Dict[str, pd.DataFrame] = {}
        self.current_sheet_name = None
        self.formula_engine = None
        self.formula_cells = {}  # {sheet_name: {row,col: formula}}
        
        # Theme
        self.style = tb.Style(DEFAULT_THEME)
        
        # Build UI
        self._setup_ui()
        
        # Load initial file or show empty sheet
        if initial_file and os.path.exists(initial_file):
            self.load_file(initial_file)
        else:
            self.update_sheet_from_dataframe()

    def _setup_ui(self):
        """Setup all UI components."""
        UIComponents.create_menu_bar(self.root, self)
        UIComponents.create_toolbar(self.root, self)
        
        # Content frame with notebook
        content_frame = tb.Frame(self.root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.sheet_notebook = ttk.Notebook(content_frame)
        self.sheet_notebook.pack(fill=tk.BOTH, expand=True)
        self.sheet_notebook.bind('<<NotebookTabChanged>>', self._on_sheet_change)
        
        # Create first sheet
        self._create_sheet_tab("Sheet1")
        
        UIComponents.create_formula_bar(self.root, self)
        UIComponents.create_status_bar(self.root, self)

    def _create_sheet_tab(self, sheet_name: str, df: Optional[pd.DataFrame] = None):
        """Create a new sheet tab."""
        tab_frame = tb.Frame(self.sheet_notebook)
        self.sheet_notebook.add(tab_frame, text=sheet_name)
        
        sheet = tksheet.Sheet(
            tab_frame, show_x_scrollbar=True, show_y_scrollbar=True,
            show_top_left=True, headers=None
        )
        sheet.enable_bindings(("single_select", "row_select", "column_select", "drag_select",
                              "edit_cell", "edit_header", "copy", "cut", "paste", "delete",
                              "undo", "redo", "find", "replace", "arrowkeys",
                              "sort_columns", "sort_rows", "rc_insert_row", "rc_delete_row",
                              "row_height_resize", "column_height_resize", "column_width_resize",
                              "row_width_resize", "rc_insert_column", "rc_delete_column"))
        
        sheet.pack(fill=tk.BOTH, expand=True)
        
        # Bind events
        sheet.extra_bindings([
            ("rc_insert_column", self.add_column),
            ("rc_insert_row", self.insert_row),
            ("cell_select", self.update_selection_status),
            ("drag_select_cells", self.update_selection_status),
            ("end_edit_cell", self.on_cell_edit),
            ("end_edit_header", self.mark_modified),
            ("rc_delete_column", self.mark_modified),
            ("rc_delete_row", self.mark_modified),
            ("paste", self.mark_modified),
            ("delete", self.mark_modified),
            ("cut", self.mark_modified),
            ("undo", self.mark_modified),
            ("redo", self.mark_modified),
        ])
        
        # Custom popup menu
        sheet.popup_menu_add_command("Autofill (Ctrl+R)", self.autofill_selection)
        sheet.popup_menu_add_command("Rename Sheet", self.rename_sheet)
        sheet.popup_menu_add_command("Trim Spaces(❌ Undo)", self.clean_whitespace)
        sheet.bind("<Control-r>", self.autofill_selection)
        
        # Store DataFrame
        if df is not None:
            self.workbook_sheets[sheet_name] = df
        else:
            self.workbook_sheets[sheet_name] = pd.DataFrame()
        
        self.current_sheet_name = sheet_name
        
        # Initialize formula engine
        if not hasattr(self, 'formula_engine') or self.formula_engine is None:
            self.formula_engine = FormulaEngine(sheet)
        
        return sheet

    def get_current_sheet(self) -> tksheet.Sheet:
        """Get currently active sheet."""
        current_tab = self.sheet_notebook.select()
        if not current_tab:
            return None
        tab_frame = self.sheet_notebook.nametowidget(current_tab)
        for widget in tab_frame.winfo_children():
            if isinstance(widget, tksheet.Sheet):
                return widget
        return None

    def rename_sheet(self):
            """Opens a dialog to rename the currently selected sheet (notebook tab)."""
            current_sheet_name = self.current_sheet_name
            if not current_sheet_name:
                messagebox.showinfo("Rename Sheet", "No sheet is currently selected.")
                return
            
            # 1. Ask for new name
            new_name = simpledialog.askstring(
                "Rename Sheet", 
                f"Enter new name for '{current_sheet_name}':", 
                initialvalue=current_sheet_name
            )

            if not new_name or new_name.strip() == current_sheet_name:
                self.set_status(f"Entered the same sheet name {self.current_sheet_name}")
                return # User canceled or entered the same name
            
            new_name = new_name.strip()

            # 2. Check for uniqueness and validity
            if new_name in self.workbook_sheets:
                messagebox.showerror("Error", f"Sheet name '{new_name}' already exists.")
                return
            
            if not new_name:
                messagebox.showerror("Error", "Sheet name cannot be empty.")
                return

            # 3. Update the Notebook tab
            try:
                # current_tab_id = self.notebook.select()
                current_tab_id = self.sheet_notebook.select()
                self.sheet_notebook.tab(current_tab_id, text=new_name)
                
                # 4. Update the internal dictionary and references
                sheet_data = self.workbook_sheets.pop(current_sheet_name) # Remove with old key
                sheet_data.name = new_name                       # Update name in SheetData object
                self.workbook_sheets[new_name] = sheet_data               # Insert with new key
                self.current_sheet_name = new_name               # Update the current reference
                
                self.set_status(f"Sheet renamed from '{current_sheet_name}' to '{new_name}'.")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to rename sheet: {e}")

    def _on_sheet_change(self, event=None):
        """Handle sheet tab change."""
        sheet = self.get_current_sheet()
        if sheet:
            self.current_sheet_name = self.sheet_notebook.tab(self.sheet_notebook.select(), "text")
            self.df = self.workbook_sheets.get(self.current_sheet_name, pd.DataFrame())
            self.formula_engine.sheet = sheet
            self.update_selection_status()
            self.set_status(f"Sheet changed to {self.current_sheet_name}")

    def add_new_sheet(self):
        """Add new sheet with save options for non-Excel."""
        sheet_num = len(self.workbook_sheets) + 1
        sheet_name = f"Sheet{sheet_num}"
        while sheet_name in self.workbook_sheets:
            sheet_num += 1
            sheet_name = f"Sheet{sheet_num}"
        
        is_excel = self.current_file and self.current_file.endswith(('.xlsx', '.xls'))
        
        if not is_excel and len(self.workbook_sheets) > 0:
            response = messagebox.askyesnocancel(
                "New Sheet",
                f"Current file is not Excel format.\n\n"
                f"YES: Save new sheet as separate file\n"
                f"NO: Add sheet (will convert to Excel on save)\n"
                f"CANCEL: Don't add sheet"
            )
            
            if response is None:
                return
            elif response is True:
                self._save_sheet_as_separate_file(sheet_name)
                return
        
        df = pd.DataFrame(columns=["Column1", "Column2"])
        self._create_sheet_tab(sheet_name, df)
        self.sheet_notebook.select(len(self.workbook_sheets) - 1)
        self.modified = True
        self.update_title()
        self.set_status(f"Added new sheet: {sheet_name}")
        self.update_dataframe_from_sheet()
        self.update_sheet_from_dataframe()
    
    def _save_sheet_as_separate_file(self, sheet_name: str):
        """Save new sheet as separate file."""
        current_dir = os.path.dirname(self.current_file) if self.current_file else os.getcwd()
        
        file_path = filedialog.asksaveasfilename(
            initialdir=current_dir,
            initialfile=f"{sheet_name}.csv",
            defaultextension=".csv",
            filetypes=[
                ("CSV", "*.csv"),
                ("TSV", "*.tsv"),
                ("Text", "*.txt"),
                ("Excel", "*.xlsx")
            ],
            title=f"Save '{sheet_name}' as separate file"
        )
        
        if not file_path:
            return
        
        try:
            df = pd.DataFrame(columns=["Column1", "Column2"])
            
            if file_path.endswith(('.xlsx', '.xls')):
                df.to_excel(file_path, index=False, sheet_name=sheet_name)
            elif file_path.endswith('.tsv'):
                df.to_csv(file_path, sep='\t', index=False)
            else:
                df.to_csv(file_path, sep=self.txt_sepa or ',', index=False)
            
            messagebox.showinfo("Success", f"New sheet saved as:\n{file_path}")
            
            if messagebox.askyesno("Open File?", "Do you want to open the new file?"):
                self.load_file(file_path)
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save separate file: {e}")

    def _populate_sheet(self, sheet_widget: tksheet.Sheet, df: pd.DataFrame):
        """Populate sheet with DataFrame."""
        if df.empty:
            sheet_widget.headers(["Column1", "Column2"])
            sheet_widget.set_sheet_data([["", ""]])
        else:
            sheet_widget.headers([str(c) for c in df.columns])
            sheet_widget.set_sheet_data(df.astype(str).values.tolist())
        
        self.toggle_dark_mode()
        sheet_widget.refresh()

    def update_sheet_from_dataframe(self):
        """Refresh table display."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        if self.df.empty:
            sheet.headers(["Column1", "Column2"])
            sheet.set_sheet_data([["", ""]])
            self.update_dataframe_from_sheet()

        else:
            sheet.headers([str(c) for c in self.df.columns])
            sheet.set_sheet_data(self.df.astype(str).values.tolist())

        self.workbook_sheets[self.current_sheet_name] = self.df
        
        self.toggle_dark_mode()
        sheet.refresh()

    def update_dataframe_from_sheet(self):
        """Sync DataFrame from sheet."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        self.df = pd.DataFrame(sheet.get_sheet_data(), columns=sheet.headers())
        self.workbook_sheets[self.current_sheet_name] = self.df

    def toggle_dark_mode(self):
        """Apply dark/light theme."""
        self.dark_mode = self.dark_mode_var.get()
        self.style.theme_use(DEFAULT_DARK_THEME if self.dark_mode else DEFAULT_THEME)
        
        # sheet_theme = "dark" if self.dark_mode else "light blue"
        sheet_theme = SHEET_DARK_THEME if self.dark_mode else SHEET_LIGHT_THEME

        for tab_id in self.sheet_notebook.tabs():
            tab_frame = self.sheet_notebook.nametowidget(tab_id)
            for widget in tab_frame.winfo_children():
                if isinstance(widget, tksheet.Sheet):
                    widget.set_options(theme=sheet_theme, align="center")
                    widget.refresh()
        
        status_bg, status_fg = ("#2e2e2e", "white") if self.dark_mode else ("#f0f0f0", "black")
        self.status_frame.config(bg=status_bg)
        self.left_status.config(bg=status_bg, fg=status_fg)
        self.right_status.config(bg=status_bg, fg=status_fg)
        self.select_status.config(bg=status_bg, fg=status_fg)
        self.cell_info.config(bg=status_bg, fg=status_fg)

    def mark_modified(self, event=None):
        """Mark sheet modified."""
        self.modified = True
        self.update_title()
        self.update_dataframe_from_sheet()
        if self.suggestion_mode_var.get():
            self.clean_dropdown_value()
            self.customize_turbine_columns()

        sheet = self.get_current_sheet()
        if sheet:
            sheet.refresh()
        self.set_status("Modified")

    def update_title(self):
        """Update window title."""
        name = os.path.basename(self.current_file) if self.current_file else "Untitled"
        self.root.title(f"{APP_TITLE} - {name}{'*' if self.modified else ''}")

    def set_status(self, text):
        """Update status bar."""
        self.right_status.config(text=text)

    def set_sel_status(self, text):
        """Update selection status."""
        self.select_status.config(text=text)

    def log_usage(self, file_path):
        """Log usage."""
        try:
            log_file = os.path.join(LOG_DIR, LOG_FILE)
            with open(log_file, "a", encoding="utf-8") as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                user = os.getlogin()
                f.write(f"{timestamp} | {user} | {file_path} | {APP_VERSION}\n")
        except:
            pass

    def ask_save_changes(self, action="continue"):
        """Ask to save changes."""
        result = messagebox.askyesnocancel("Unsaved Changes", 
                                          f"Save changes before {action}?")
        if result is True:
            self.save_file()
            return True
        if result is None:
            return False
        return True

    def on_close(self):
        """Handle window close."""
        if self.modified and not self.ask_save_changes("exit"):
            return
        self.root.destroy()
    def load_file(self, file_path: str):
        """Load file with Excel multi-sheet support."""
        try:
            for tab_id in self.sheet_notebook.tabs():
                self.sheet_notebook.forget(tab_id)
            self.workbook_sheets.clear()
            
            if file_path.endswith(('.xlsx', '.xls')):
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, keep_default_na=False)
                    sheet_widget = self._create_sheet_tab(sheet_name, df)
                    self._populate_sheet(sheet_widget, df)
                
                self.txt_sepa = None
                
            else:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    first_line = f.readline()

                if self.current_sep=="#label" and "#label" in first_line.lower():
                    # print(file_path)
                    _, _, _, self.df = read_label_format(file_path)
                    self.txt_sepa = "#label"

                elif self.current_sep:
                    sep_used = self.current_sep
                    self.df = pd.read_csv(file_path, sep=sep_used, engine="python", keep_default_na=False, na_values=[])
                    self.txt_sepa = sep_used
                else:
                    # with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    #     first_line = f.readline()
                    if file_path.endswith(".tsv") or "\t" in first_line:
                        self.df = pd.read_csv(file_path, sep=r"\t", engine="python", keep_default_na=False, na_values=[])
                        self.txt_sepa = "\t"
                    elif ";" in first_line:
                        self.df = pd.read_csv(file_path, sep=";", engine="python", keep_default_na=False, na_values=[])
                        self.txt_sepa = ";"

                    elif "," in first_line:
                        self.df = pd.read_csv(file_path, sep=",", engine="python", keep_default_na=False, na_values=[])
                        self.txt_sepa = ","

                    elif file_path.endswith(".txt") and "#label" in first_line.lower():
                        _, _, _, self.df = read_label_format(file_path)
                        self.txt_sepa = "#label"

                    elif file_path.endswith(".txt") and " " in first_line:
                        self.df = pd.read_csv(file_path, sep=r"\s+", engine="python", keep_default_na=False, na_values=[])
                        self.txt_sepa = r"\s+"
                    else:
                        self.df = pd.read_csv(file_path, keep_default_na=False, na_values=[])
                        self.txt_sepa = ","
                
                sheet_name = os.path.splitext(os.path.basename(file_path))[0]
                sheet_widget = self._create_sheet_tab(sheet_name, self.df)
                self._populate_sheet(sheet_widget, self.df)
            
            self.current_file = file_path
            self.modified = False
            self.update_title()  
            if self.suggestion_mode_var.get():
                self.customize_turbine_columns()
            self.set_status(f"Opened: {os.path.basename(file_path)}")          
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {e}")
            self.set_status("Error opening file")


    def save_file(self):
        """Save file with Excel multi-sheet support."""
        if self.current_file is None:
            self.save_file_as()
            return
        
        self.update_dataframe_from_sheet()
        
        try:
            if self.current_file.endswith(('.xlsx', '.xls')):
                with pd.ExcelWriter(self.current_file, engine='openpyxl') as writer:
                    for sheet_name, df in self.workbook_sheets.items():
                        df_to_save = df.copy()
                        df_to_save.replace(r'^\s*', '', regex=True, inplace=True)
                        df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df_to_save = self.df.copy()
                # df_to_save.replace(r'^\s*', '', regex=True, inplace=True)
                df_to_save.replace(r'^\s*$', 'NaN', regex=True, inplace=True)
                if self.txt_sepa == "#label":
                    write_label_format(self.current_file, self.df, sep="::")
                elif self.txt_sepa == r'\s+' or re.match(r'\s+', self.txt_sepa):
                    df_to_save.to_string(buf=self.current_file, index=False, col_space=15, justify='left')
                elif self.current_file.endswith(".tsv"):
                    df_to_save.to_csv(self.current_file, sep="\t", index=False)
                else:
                    self.df.to_csv(self.current_file, sep=self.txt_sepa, index=False, na_rep="")
            
            self.modified = False
            self.update_title()
            self.set_status(f"Saved: {os.path.basename(self.current_file)}")
            messagebox.showinfo("Saved", "File saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")
            self.set_status("Error saving file")


    def save_file_as(self):
        """Save as with format selection."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("CSV", "*.csv"),
                ("TSV", "*.tsv"),
                ("Text", "*.txt")
            ]
        )
        if not file_path:
            return
        self.current_file = file_path
        self.save_file()


    def export_to_excel(self):
        """Export to Excel with formatting."""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
        except ImportError:
            messagebox.showerror("Error", "openpyxl not installed. Using basic export.")
            self.save_file_as()
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not file_path:
            return
        
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet_name, df in self.workbook_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    worksheet = writer.sheets[sheet_name]
                    
                    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    header_font = Font(color="FFFFFF", bold=True)
                    
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(cell.value)
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "Excel file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_csv(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not file_path:
            return
        
        try:
            self.df.to_csv(file_path, sep=',', index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "CSV file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_tsv(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".tsv", filetypes=[("TSV", "*.tsv")])
        if not file_path:
            return
        
        try:
            self.df.to_csv(file_path, sep='\t', index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "TSV file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_txt(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt *.dat")])
        if not file_path:
            return
            
        sep = simpledialog.askstring("Export Options", "Enter separator (use \\t for tab):", 
                                       initialvalue=",")
        if not sep:
            return
            
        if sep == "\\t": sep = "\t"
        
        try:
            self.df.to_csv(file_path, sep=sep, index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "Text file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_lib(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Library Label", "*.txt")])
        if not file_path:
            return
        try:
            write_label_format(file_path, self.df, sep="::")
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "Library label file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}") 
            
    def export_to_md(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".md", filetypes=[("Markdown", "*.md")])
        if not file_path:
            return
        
        try:
            self.df.to_markdown(file_path, index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "Markdown file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_html(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML", "*.html")])
        if not file_path:
            return
        
        try:
            self.df.to_html(file_path, index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "HTML file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_json(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not file_path:
            return
        
        try:
            self.df.to_json(file_path, orient='records', indent=4)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "JSON file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_xml(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML", "*.xml")])
        if not file_path:
            return
        
        try:
            self.df.to_xml(file_path, index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "XML file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def export_to_tex(self):
        """Export to Excel with formatting."""
        
        file_path = filedialog.asksaveasfilename(defaultextension=".tex", filetypes=[("LaTeX", "*.tex")])
        if not file_path:
            return
        
        try:
            self.df.to_latex(file_path, index=False)
            self.set_status(f"Exported to: {os.path.basename(file_path)}")
            messagebox.showinfo("Success", "LaTeX file exported with formatting!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
            
    def open_file(self):
        """Open file with multi-format support."""
        if self.modified and not self.ask_save_changes("open a different file"):
            return
        
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("All Supported", "*.csv *.txt *.tsv *.xlsx *.xls"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Text files", "*.txt *.tsv"),
                ("All files", "*.*")
            ]
        )
        if not file_path:
            return
        
        self.load_file(file_path)
        self.log_usage(file_path)


    def new_file(self):
        """Create new file."""
        if self.modified and not self.ask_save_changes("create a new file"):
            return
        
        for tab_id in self.sheet_notebook.tabs():
            self.sheet_notebook.forget(tab_id)
        self.workbook_sheets.clear()
        
        self.df = pd.DataFrame(columns=["Column1", "Column2"])
        sheet_widget = self._create_sheet_tab("Sheet1", self.df)
        self._populate_sheet(sheet_widget, self.df)
        
        self.current_file = None
        self.modified = False
        self.update_title()
        self.update_sheet_from_dataframe()
        self.update_dataframe_from_sheet()
        self.set_status("New file created")


    # Formula Methods
    def _on_formula_enter(self, event=None):
        """Handle formula entry from formula bar."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_currently_selected()
        if not selected:
            return
        
        row, col = selected[0], selected[1]
        formula = self.formula_var.get()
        
        if formula.startswith('='):
            if self.current_sheet_name not in self.formula_cells:
                self.formula_cells[self.current_sheet_name] = {}
            
            cell_key = f"{row},{col}"
            self.formula_cells[self.current_sheet_name][cell_key] = formula
            
            if self.auto_calc_var.get():
                result = self.formula_engine.evaluate_formula(formula, row, col)
                sheet.set_cell_data(row, col, str(result), redraw=True)
                sheet.highlight_cells(row=row, column=col, bg="lightyellow", fg = "black")
                self.set_status(f"Formula: {formula} = {result}")
            else:
                sheet.set_cell_data(row, col, formula, redraw=True)
                sheet.highlight_cells(row=row, column=col, bg="lightblue", fg = "black")
                self.set_status(f"Formula stored (not calculated). Click 'Calc' button.")
        else:
            sheet.set_cell_data(row, col, formula, redraw=True)
            if self.current_sheet_name in self.formula_cells:
                cell_key = f"{row},{col}"
                self.formula_cells[self.current_sheet_name].pop(cell_key, None)
                sheet.dehighlight_cells(row=row, column=col)
        
        self.mark_modified()


    def _on_formula_focus_out(self, event=None):
        """Handle when formula bar loses focus."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_currently_selected()
        if not selected:
            return
        
        row, col = selected[0], selected[1]
        formula = self.formula_var.get()
        
        cell_key = f"{row},{col}"
        sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})
        current_formula = sheet_formulas.get(cell_key, "")
        current_value = sheet.get_cell_data(row, col)
        
        if formula != current_formula and formula != current_value:
            self._on_formula_enter()


    def calculate_current_cell(self):
        """Calculate the currently selected cell (manual calc mode)."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_currently_selected()
        if not selected:
            messagebox.showinfo("Info", "Please select a cell first")
            return
        
        row, col = selected[0], selected[1]
        cell_key = f"{row},{col}"
        sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})
        
        if cell_key not in sheet_formulas:
            messagebox.showinfo("Info", "Selected cell does not contain a formula")
            return
        
        formula = sheet_formulas[cell_key]
        result = self.formula_engine.evaluate_formula(formula, row, col)
        sheet.set_cell_data(row, col, str(result), redraw=True)
        sheet.highlight_cells(row=row, column=col, bg="lightyellow", fg = "black")
        self.set_status(f"Calculated: {formula} = {result}")
        self.mark_modified()


    def on_cell_edit(self, event=None):
        """Handle cell edit with formula evaluation."""
        self.mark_modified()
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_currently_selected()
        if not selected:
            return
        
        row, col = selected[0], selected[1]
        value = sheet.get_cell_data(row, col)
        
        if isinstance(value, str) and value.startswith('='):
            if self.current_sheet_name not in self.formula_cells:
                self.formula_cells[self.current_sheet_name] = {}
            
            cell_key = f"{row},{col}"
            self.formula_cells[self.current_sheet_name][cell_key] = value
            
            if self.auto_calc_var.get():
                result = self.formula_engine.evaluate_formula(value, row, col)
                sheet.set_cell_data(row, col, str(result), redraw=True)
                sheet.highlight_cells(row=row, column=col, bg="lightyellow", fg = "black")
                self.set_status(f"Formula: {value} = {result}")
            else:
                sheet.highlight_cells(row=row, column=col, bg="lightblue", fg = "black")
                self.set_status(f"Formula entered (not calculated). Use 'Calc' button.")
        else:
            if self.current_sheet_name in self.formula_cells:
                cell_key = f"{row},{col}"
                self.formula_cells[self.current_sheet_name].pop(cell_key, None)
                sheet.dehighlight_cells(row=row, column=col)


    def recalculate_all(self):
        """Recalculate all formulas in current sheet."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        count = 0
        sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})
        
        if not sheet_formulas:
            messagebox.showinfo("Info", "No formulas found in current sheet")
            return
        
        for cell_key, formula in sheet_formulas.items():
            try:
                r, c = map(int, cell_key.split(','))
                result = self.formula_engine.evaluate_formula(formula, r, c)
                sheet.set_cell_data(r, c, str(result), redraw=False)
                sheet.highlight_cells(row=r, column=c, bg="lightyellow", fg = "black")
                count += 1
            except Exception as e:
                print(f"Error recalculating {cell_key}: {e}")
        
        sheet.refresh()
        self.set_status(f"Recalculated {count} formulas")
        messagebox.showinfo("Recalculate All", f"Successfully recalculated {count} formulas!")


    def insert_formula_template(self, template: str):
        """Insert formula template."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_currently_selected()
        if not selected:
            messagebox.showinfo("Info", "Please select a cell first")
            return
        
        row, col = selected[0], selected[1]
        
        if self.current_sheet_name not in self.formula_cells:
            self.formula_cells[self.current_sheet_name] = {}
        
        cell_key = f"{row},{col}"
        self.formula_cells[self.current_sheet_name][cell_key] = template
        
        if self.auto_calc_var.get():
            result = self.formula_engine.evaluate_formula(template, row, col)
            sheet.set_cell_data(row, col, str(result), redraw=True)
            sheet.highlight_cells(row=row, column=col, bg="lightyellow", fg = "black")
            self.set_status(f"Formula inserted: {template} = {result}")
        else:
            sheet.set_cell_data(row, col, template, redraw=True)
            sheet.highlight_cells(row=row, column=col, bg="lightblue", fg = "black")
            self.set_status(f"Formula inserted (not calculated): {template}")
        
        self.formula_var.set(template)
        self.mark_modified()

    def update_selection_status(self, event=None):
        """Enhanced selection status with cell info."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = sheet.get_selected_cells()
        if not selected:
            self.set_sel_status("Ready")
            self.cell_info.config(text="")
            self.formula_var.set("")
            return
        
        selected = list(selected)
        
        if len(selected) == 1:
            r, c = selected[0]
            headers = sheet.headers()
            col_name = headers[c] if c < len(headers) else f"Col{c + 1}"
            
            col_letter = self._num_to_col_letter(c)
            cell_ref = f"{col_letter}{r + 1}"
            
            cell_key = f"{r},{c}"
            sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})
            if cell_key in sheet_formulas:
                self.formula_var.set(sheet_formulas[cell_key])
            else:
                value = sheet.get_cell_data(r, c)
                self.formula_var.set(value or '')
            
            self.set_sel_status(f"Cell: {cell_ref} ({col_name})")
            self.cell_info.config(text=f"Row {r + 1}, Col {c + 1}")
            return
        
        rows = [r for r, _ in selected]
        cols = [c for _, c in selected]
        row_span = max(rows) - min(rows) + 1
        col_span = max(cols) - min(cols) + 1
        
        values = []
        for r, c in selected:
            try:
                val = sheet.get_cell_data(r, c)
                if val not in ("", None):
                    values.append(float(val))
            except ValueError:
                pass
        
        stats = ""
        if values:
            count = len(values)
            s = sum(values)
            avg = s / count if count else 0
            stats = f" | Count={count}, Sum={s:.2f}, Avg={avg:.2f}"
        
        self.set_sel_status(f"Selection: {row_span}R × {col_span}C{stats}")
        self.cell_info.config(text=f"{len(selected)} cells")


    def _num_to_col_letter(self, n: int) -> str:
        """Convert column number to Excel letter."""
        result = ""
        n += 1
        while n > 0:
            n -= 1
            result = chr(65 + (n % 26)) + result
            n //= 26
        return result


    def autofill_selection(self, event=None):
        """Enhanced autofill with pattern detection."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        selected = list(sheet.get_selected_cells())
        if not selected:
            return
        
        rows = [r for r, _ in selected]
        cols = [c for _, c in selected]
        min_r, max_r = min(rows), max(rows)
        min_c, max_c = min(cols), max(cols)
        
        # Vertical fill
        if max_r > min_r and min_c == max_c:
            col = min_c
            values = [sheet.get_cell_data(r, col) for r in range(min_r, max_r + 1)]
            nums = []
            for v in values:
                if v and re.search(r"(\d+)$", str(v)):
                    nums.append(int(re.search(r"(\d+)$", str(v)).group(1)))
                elif v:
                    nums.append(v)
            
            step = nums[-1] - nums[-2] if len([x for x in nums if isinstance(x, int)]) >= 2 else 0
            
            last_val = None
            for i, v in enumerate(values):
                if v:
                    last_val = v
                    continue
                if isinstance(last_val, str) and re.search(r"(\d+)$", last_val):
                    prefix = re.sub(r"\d+$", "", last_val)                
                    if prefix=='-':
                        prefix = ""
                        num = int(re.search(r"(\d+)$", last_val).group(1))*-1 + step
                        new_val = f"{prefix}{num}"
                    else:
                        num = int(re.search(r"(\d+)$", last_val).group(1)) + step
                    new_val = f"{prefix}{num}"
                elif isinstance(last_val, int):
                    new_val = last_val + step
                else:
                    new_val = last_val
                values[i] = str(new_val)
                last_val = values[i]
                sheet.set_data(min_r + i, col, data=values[i], undo=True)
        
        # Horizontal fill
        elif max_c > min_c and min_r == max_r:
            row = min_r
            values = [sheet.get_cell_data(row, c) for c in range(min_c, max_c + 1)]
            nums = []
            for v in values:
                if v and re.search(r"(\d+)$", str(v)):
                    nums.append(int(re.search(r"(\d+)$", str(v)).group(1)))
                elif v:
                    nums.append(v)
            
            step = nums[-1] - nums[-2] if len([x for x in nums if isinstance(x, int)]) >= 2 else 0
            
            last_val = None
            for i, v in enumerate(values):
                if v:
                    last_val = v
                    continue
                if isinstance(last_val, str) and re.search(r"(\d+)$", last_val):
                    prefix = re.sub(r"\d+$", "", last_val)
                    if prefix=='-':
                        prefix = ""
                        num = int(re.search(r"(\d+)$", last_val).group(1))*-1 + step
                    else:
                        num = int(re.search(r"(\d+)$", last_val).group(1)) + step

                    new_val = f"{prefix}{num}"
                elif isinstance(last_val, int):
                    new_val = last_val + step
                else:
                    new_val = last_val
                values[i] = str(new_val)
                last_val = values[i]
                sheet.set_data(row, min_c + i, data=values[i], undo=True)

        self.update_dataframe_from_sheet()     
        if self.suggestion_mode_var.get():
            self.customize_turbine_columns()
        self.modified = True
        self.update_title()
        self.set_status("Autofill applied")

    def add_column(self, event=None):
        """Insert column."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        if event:
            col_index = sheet.get_currently_selected()[1] if sheet.get_currently_selected() else len(self.df.T)
            new_col = f"Column{len(self.df.columns) + 1}"
            # self.update_dataframe_from_sheet()
            print('HereC1')
        else:
            col_index = len(self.df.T)
            new_col = f"Column{len(self.df.columns) + 1}"
            # self.df.insert(col_index, new_col, "")
            # self.update_sheet_from_dataframe()
            sheet.insert_column(idx=col_index, header=new_col)

        sheet.set_header_data(value=new_col, c=col_index)
        self.update_dataframe_from_sheet()
        
        self.modified = True
        self.update_title()
        self.set_status(f"Added column {new_col}")


    def insert_row(self, event=None):
        """Insert row."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        print('Here')
        if event:
            idx = sheet.get_currently_selected()[0] if sheet.get_currently_selected() else len(self.df)
            self.update_dataframe_from_sheet()
            print('Here1')
        else:
            idx = len(self.df)
            # empty_row = pd.Series([""] * len(self.df.columns), index=self.df.columns)
            # self.df = pd.concat([self.df.iloc[:idx], empty_row.to_frame().T, self.df.iloc[idx:]]).reset_index(drop=True)
            sheet.insert_row(idx=idx, redraw=True)
            self.update_dataframe_from_sheet()
        
        self.modified = True
        self.update_title()
        
        if self.suggestion_mode_var.get():
            self.customize_turbine_columns()
        self.set_status(f"Added row {idx+1}")
        
    def clean_whitespace(self, event=None):
        """Insert column."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        if event:
            col_index = sheet.get_currently_selected()[1] if sheet.get_currently_selected() else len(self.df.T)
            col = list(self.df.columns)
            clean_col = col[col_index]
            col_name = clean_col
        else:
            col_name = None
            clean_col = 'All'
            
        self.df = trim_whitespace(df = self.df, columns = col_name)
        self.update_sheet_from_dataframe()
        
        self.modified = True
        self.update_title()
        self.set_status(f"Extra space cleaned on columns (No Undo): {clean_col}")


    def clean_nan(self, event=None):
        """Clean NaN/None values, prompting the user for the fill value."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        if event:
            col_index = sheet.get_currently_selected()[1] if sheet.get_currently_selected() else len(self.df.T)
            col = list(self.df.columns)
            clean_col = col[col_index]
            col_name = [clean_col] 
        else:
            col_name = None
            clean_col = 'All'

        # The simpledialog.askstring returns None if the user cancels the dialog.
        fill_value_input = simpledialog.askstring(
            "Fill Nulls", 
            f"Enter the fill value for column: {clean_col}",
            initialvalue=''
        )

        # Check if the user cancelled the dialog
        if fill_value_input is None:
            self.set_status("NaN cleaning cancelled.")
            return

        # For now, we treat the input as a string for simplicity, as per your original " " fill_value.
        fill_val = fill_value_input if fill_value_input != "" else " " 

        # 3. Apply the cleaning function
        self.df = fill_nulls(
            df = self.df, 
            columns = col_name, 
            fill_value = fill_val, 
            only_object_columns=False
        )
        
        self.update_sheet_from_dataframe()
        
        self.modified = True
        self.update_title()
        self.set_status(f"Cleaned NaN (No Undo): {clean_col} col with value '{fill_val}'")

    def change_separator(self):
        """Change CSV separator."""
        popup = tb.Toplevel(self.root)
        popup.title("Change Separator")
        popup.geometry("350x250")
        
        tb.Label(popup, text="Select separator:", font=("Segoe UI", 10, "bold")).pack(pady=10)
        
        sep_var = tk.StringVar(value=self.current_sep or self.txt_sepa)
        sep_choices = [
            ("None", "Auto"),
            (",", "Comma (,)"),
            (";", "Semicolon (;)"),
            ("\t", "Tab (\\t)"),
            (r"\s+", "Spaces(   )"),
            (" ", "Space ( )"),
            ("|", "Pipe (|)"),
            ("#label", "Library format"),
        ]
        combo = tb.Combobox(popup, values=[f"{s[0]} - {s[1]}" for s in sep_choices], textvariable=sep_var)
        combo.pack(pady=5, padx=15, fill="x")
        
        def apply_separator():
            sep_text = sep_var.get().split(" - ")[0] if " - " in sep_var.get() else sep_var.get()
            if sep_text.lower() in ["\\t", "tab"]:
                sep_text = "\t"
            elif sep_text.lower() in ["none", "auto"]:
                sep_text = None
            
            self.current_sep = sep_text
            self.txt_sepa = sep_text
            popup.destroy()
            
            if self.current_file:
                self.load_file(self.current_file)
            
            self.set_status(f"Separator: {repr(sep_text)}")
        
        tb.Button(popup, text="Apply", command=apply_separator, bootstyle="success").pack(pady=10)


    # MDT Features
    def toggle_suggestion_mode(self):
        """Toggle MDT suggestion mode."""
        if self.suggestion_mode_var.get():
            self.set_status("MDT Suggestion Mode: ON")
            self.customize_turbine_columns()
        else:
            self.clean_all_dropdown()
            self.set_status("MDT Suggestion Mode: OFF")


    def customize_turbine_columns(self):
        """Attach dropdown suggestions based on file type prefix."""
        if not self.current_file:
            return
        
        fname = os.path.basename(self.current_file)
        self.clean_all_dropdown()
        
        # Get dropdown config for this file type
        dropdown_config = get_mdt_dropdown_config(fname)
        
        if not dropdown_config:
            self.suggestion_mode_var.set(False)
            self.toggle_suggestion_mode()
            self.set_status("Suggestion Mode OFF (this file not supported)")
            return
        
        # Special handling for Parasolid column
        if fname.startswith("Turbine_Comp") and "Parasolid" in self.df.columns:
            folder_path = os.path.dirname(self.current_file)
            xt_files = get_file_list_for_column(folder_path, "Parasolid")
            self.set_column_dropdown("Parasolid", xt_files)
        
        # Apply all other dropdowns
        for column_name, options in dropdown_config.items():
            if column_name in self.df.columns and column_name != "Parasolid":
                self.set_column_dropdown(column_name, options)


    def set_column_dropdown(self, column_name, options):
        """Attach dropdown to all cells in a column."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        if column_name not in self.df.columns:
            return
        col_idx = self.df.columns.get_loc(column_name)
        
        if isinstance(options, dict):
            dropdown_display = [f"{val} - {desc}" for val, desc in options.items()]
            valid_values = set(options.keys())
        else:
            dropdown_display = list(options)
            valid_values = set(options)
        
        for r in range(len(self.df)):
            existing_value = sheet.get_cell_data(r, col_idx)
            sheet.dropdown(r, col_idx, values=dropdown_display, state="normal", validate_input=False)
            sheet.set_cell_data(r, col_idx, value=existing_value)
            if existing_value not in valid_values:
                sheet.highlight_cells(row=r, column=col_idx, bg="lightcoral")


    def clean_dropdown_value(self):
        """If a selected cell shows 'value - description', store only 'value'."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        selected = list(sheet.get_selected_cells())
        if len(selected) != 1:
            return
        r, c = selected[0]
        if r is None or c is None:
            return
        val = sheet.get_cell_data(r, c)
        if val and " - " in str(val):
            clean_val = str(val).split(" - ")[0].strip()
            sheet.set_cell_data(r, c, clean_val)


    def clean_all_dropdown(self):
        """Remove all dropdowns."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        
        for r in range(sheet.total_rows()):
            for c in range(sheet.total_columns()):
                sheet.del_dropdown(r, c)

    def run_component_graphs(self):
        """Plot interactive component graphs for matching Turbine_Conta_Catalogue_*.csv files."""
        if not self.current_file:
            self.set_status("No File: Please open a file first.")
            return

        root_dir = os.path.dirname(self.current_file)

        conta_files = [
            f for f in os.listdir(root_dir)
            if f.startswith("Turbine_Conta_Catalogue_") and f.endswith(".csv")
        ]
        comp_file = next(
            (f for f in os.listdir(root_dir)
             if f.startswith("Turbine_Comp_Catalogue_") and f.endswith(".csv")),
            None
        )
        comp_path = os.path.join(root_dir, comp_file) if comp_file else None

        if not conta_files:
            self.set_status("No Matching Files: No Turbine_Conta_Catalogue_*.csv files found.")
            return

        for fname in conta_files:
            fpath = os.path.join(root_dir, fname)
            out_html = os.path.join(root_dir, f"graph_{fname.split('.')[0]}.html")
            try:
                plot_interactive_component_graph(
                    fpath, output_file=out_html, open_browser=True, comp_csv=comp_path
                )
                self.set_status(f"Graph plot done: open {out_html}")
            except Exception as e:
                self.set_status(f"Graph Error: Error processing {fname}: {e}")

    def run_component_graphs_coord(self):
        """Plot interactive component graphs for matching Turbine_Conta_Catalogue_*.csv files."""
        if not self.current_file:
            self.set_status("No File: Please open a file first.")
            return

        root_dir = os.path.dirname(self.current_file)

        conta_files = [
            f for f in os.listdir(root_dir)
            if f.startswith("coordinateTable") and f.endswith(".txt")
        ]
        comp_file = next(
            (f for f in os.listdir(root_dir)
             if f.startswith("Turbine_Comp_Catalogue_") and f.endswith(".csv")),
            None
        )
        comp_path = os.path.join(root_dir, comp_file) if comp_file else None

        if not conta_files:
            self.set_status("No Matching Files: No coordinateTable*.txt files found.")
            return

        for fname in conta_files:
            fpath = os.path.join(root_dir, fname)
            out_html = os.path.join(root_dir, f"graph_{fname.split('.')[0]}.html")
            try:
                _, _, _, df = read_label_format(fpath)
                plot_interactive_component_graph_coordinate(
                    df, output_file=out_html, open_browser=True, comp_csv=comp_path
                )
                self.set_status(f"Graph plot done: open {out_html}")
            except Exception as e:
                self.set_status(f"Graph Error: Error processing {fname}: {e}")

    def show_formula_help(self):
        """Show formula help dialog."""
        help_window = tb.Toplevel(self.root)
        help_window.title("Formula Reference")
        help_window.geometry("700x600")
        
        frame = tb.Frame(help_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = tb.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text = tk.Text(frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, font=("Consolas", 10))
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)
        
        help_text = FORMULA_HELP_TEXT
        
        text.insert(1.0, help_text)
        text.config(state=tk.DISABLED)
        
        tb.Button(help_window, text="Close", command=help_window.destroy, bootstyle="secondary").pack(pady=10)
