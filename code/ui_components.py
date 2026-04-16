"""
UI Components - Menus, Toolbars, Status Bar
"""
import os
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
from ttkbootstrap.tooltip import ToolTip
import webbrowser
from config import REPORT_EMAIL, MDT_HELP_URL, USER_GUIDE_PATH, ABOUT_TEXT, LEFT_STATUS

class UIComponents:
    """Handles creation of UI elements like menus and toolbars."""
    
    @staticmethod
    def create_menu_bar(root, app):
        """Create the menu bar with all menus."""
        menubar = tk.Menu(root)
        root.config(menu=menubar)
        
        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Session", command=app.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Open File", command=app.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Add File to Session", command=app.add_file, accelerator="Ctrl+Shift+O")       
        
        # Build recent files submenu
        recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label='Recent Files', menu=recent_menu)
        def _populate_recent_menu():
            """Clear and rebuild the recent files submenu."""
            recent_menu.delete(0, 'end')   # clear all entries first
            paths = app._load_recent()
            if not paths:
                recent_menu.add_command(label='(no recent files)', state='disabled')
            else:
                for fp in paths:
                    # label = os.path.basename(fp)   # show filename only
                    label = fp   # show full path
                    # Use default arg to capture fp by value (classic Python loop closure fix)
                    recent_menu.add_command(
                        label=label,
                        command=lambda p=fp: app.load_file(p)
                    )
            recent_menu.add_separator()
            recent_menu.add_command(label='Clear Recent Files',
                                    command=lambda: _clear_recent())
        
        def _clear_recent():
            import json
            try:
                rp = app._recent_path()
                with open(rp, 'w') as f: json.dump([], f)
            except Exception: pass
            _populate_recent_menu()
            app.set_status('Recent files list cleared')
        _populate_recent_menu()

        file_menu.add_separator()
        file_menu.add_command(label="Save", command=app.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As", command=app.save_file_as)
        file_menu.add_command(label="Save All   Ctrl+Shift+S",  command=app.save_all)
        file_menu.add_separator()
        file_menu.add_command(label="Copy Full Path",           command=app.copy_full_path)
        file_menu.add_command(label="Copy File Name",           command=app.copy_file_name)
        file_menu.add_command(label="Open Containing Folder",   command=app.open_containing_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Export to Excel", command=app.export_to_excel)
        file_menu.add_command(label="Export to CSV", command=app.export_to_csv)
        file_menu.add_command(label="Export to TSV", command=app.export_to_tsv)
        file_menu.add_command(label="Export to Text", command=app.export_to_txt)
        file_menu.add_command(label="Export to Library", command=app.export_to_lib)
        file_menu.add_command(label="Export to Markdown", command=app.export_to_md)
        file_menu.add_command(label="Export to HTML", command=app.export_to_html)
        file_menu.add_command(label="Export to JSON", command=app.export_to_json)
        file_menu.add_command(label="Export to XML", command=app.export_to_xml)
        file_menu.add_command(label="Export to LaTeX", command=app.export_to_tex)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=app.on_close)
        
        # Edit Menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Undo", command=lambda: app.undo(), accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=lambda: app.redo(), accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="Compare", command=lambda: app.open_diff_dialog(), accelerator="Ctrl+D")
        edit_menu.add_separator() 
        edit_menu.add_command(label="Trim Whitespace", command=lambda: app.clean_whitespace())
        edit_menu.add_command(label="Fill Nulls", command=lambda: app.clean_nan())
        edit_menu.add_command(label="Autofill Selection", command=lambda: app.autofill_selection(), accelerator="Ctrl+R")
        edit_menu.add_separator()
        # edit_menu.add_command(label="Find", command=lambda: app.get_current_sheet().find(), accelerator="Ctrl+F")
        
        # Insert Menu
        insert_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Insert", menu=insert_menu)
        insert_menu.add_command(label="Insert Row", command=app.insert_row)
        insert_menu.add_command(label="Insert Column", command=app.add_column)
        insert_menu.add_separator()
        insert_menu.add_command(label="Insert Sheet", command=app.add_new_sheet)
        insert_menu.add_command(label="Close sheet", command=app.close_current_tab)
        insert_menu.add_command(label="Rename sheet", command=app.rename_sheet)
        
        # View Menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Open Other View",command=lambda: app.toggle_split_view(), accelerator="Ctrl+Shift+V")
        view_menu.add_command(label="Open Current Tab in Other View",command=lambda: app.open_in_other_view(app.current_sheet_name))
        view_menu.add_separator()     
        view_menu.add_command(label="Toggle Filter Bar",   command=app.toggle_filter_bar)
        view_menu.add_separator()
        view_menu.add_command(label="Column Width: Wide",        command=lambda: app.set_col_width_preset("wide"))
        view_menu.add_command(label="Column Width: Normal",      command=lambda: app.set_col_width_preset("normal"))
        view_menu.add_command(label="Column Width: Compact",     command=lambda: app.set_col_width_preset("compact"))
        view_menu.add_command(label="Column Width: Fit Content", command=lambda: app.set_col_width_preset("fit"))
        
        # Formulas Menu
        UIComponents._create_formula_menu(menubar, app)
        
        # Settings menu 
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Preferences...",        command=app.open_preferences)
        settings_menu.add_command(label="Theme Picker",          command=app.open_theme_picker)
        view_menu.add_separator()
        view_menu.add_command(label="Command Palette",   command=lambda: app.open_command_palette(), accelerator="Ctrl+P")

        # Tools Menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        app.suggestion_mode_var = tk.BooleanVar(value=False)
        tools_menu.add_checkbutton(label="MDT Suggestion Mode", variable=app.suggestion_mode_var, command=app.toggle_suggestion_mode)
        tools_menu.add_command(label="Plot COMP Graph Conta", command=app.run_component_graphs)
        tools_menu.add_command(label="Plot COMP Graph Coord", command=app.run_component_graphs_coord)
        tools_menu.add_separator()
        tools_menu.add_command(label="Change Separator", command=app.change_separator)
        tools_menu.add_command(label="Column stat", command=app.show_column_stats)
        
        # Format
        format_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Format", menu=format_menu)
        format_menu.add_command(label="Conditional Formatting",command=app.open_cf_panel)
        format_menu.add_command(label="Clear All CF Highlights",command=app.clear_cf_for_tab)
        
        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=UIComponents.open_app_help)
        help_menu.add_command(label="MDT Documentation", command=UIComponents.open_help)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=UIComponents.show_about)
    
    @staticmethod
    def _create_formula_menu(menubar, app):
        """Create the Formulas menu with submenus."""
        formula_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Formulas", menu=formula_menu)
        formula_menu.add_command(label="Recalculate All", command=app.recalculate_all)
        formula_menu.add_separator()
        
        # Math Functions
        math_menu = tk.Menu(formula_menu, tearoff=0)
        formula_menu.add_cascade(label="Math Functions", menu=math_menu)
        math_menu.add_command(label="SUM", command=lambda: app.insert_formula_template("=SUM(A1:A10)"))
        math_menu.add_command(label="AVERAGE", command=lambda: app.insert_formula_template("=AVERAGE(A1:A10)"))
        math_menu.add_command(label="MAX", command=lambda: app.insert_formula_template("=MAX(A1:A10)"))
        math_menu.add_command(label="MIN", command=lambda: app.insert_formula_template("=MIN(A1:A10)"))
        math_menu.add_command(label="SQRT", command=lambda: app.insert_formula_template("=SQRT(A1)"))
        math_menu.add_command(label="POWER", command=lambda: app.insert_formula_template("=POWER(A1,2)"))
        
        # Logical Functions
        logical_menu = tk.Menu(formula_menu, tearoff=0)
        formula_menu.add_cascade(label="Logical Functions", menu=logical_menu)
        logical_menu.add_command(label="IF", command=lambda: app.insert_formula_template('=IF(A1>0,"Yes","No")'))
        logical_menu.add_command(label="COUNTIF", command=lambda: app.insert_formula_template('=COUNTIF(A1:A10,">5")'))
        logical_menu.add_command(label="SUMIF", command=lambda: app.insert_formula_template('=SUMIF(A1:A10,">5")'))
        
        # Text Functions
        text_menu = tk.Menu(formula_menu, tearoff=0)
        formula_menu.add_cascade(label="Text Functions", menu=text_menu)
        text_menu.add_command(label="CONCAT", command=lambda: app.insert_formula_template('=CONCAT(A1,B1)'))
        text_menu.add_command(label="UPPER", command=lambda: app.insert_formula_template('=UPPER(A1)'))
        text_menu.add_command(label="LOWER", command=lambda: app.insert_formula_template('=LOWER(A1)'))
        
        # Python Expressions
        python_menu = tk.Menu(formula_menu, tearoff=0)
        formula_menu.add_cascade(label="Python Expressions", menu=python_menu)
        python_menu.add_command(label="Basic", command=lambda: app.insert_formula_template('=PYTHON(A1*2+B1)'))
        python_menu.add_command(label="Square Root", command=lambda: app.insert_formula_template('=PYTHON(sqrt(A1))'))
        python_menu.add_command(label="Trigonometry", command=lambda: app.insert_formula_template('=PYTHON(sin(A1))'))
        
        formula_menu.add_separator()
        formula_menu.add_command(label="Formula Help", command=app.show_formula_help)
    
    @staticmethod
    def create_toolbar(root, app):
        """Create the toolbar with buttons."""
        toolbar = tb.Frame(root)
        toolbar.pack(fill=tk.X, pady=5, padx=5)
        
        # File Frame
        file_frame = tb.Labelframe(toolbar, text="File", padding=5)
        file_frame.pack(side=tk.LEFT, padx=5)
        
        b1 = tb.Button(file_frame, text="📄", command=app.add_new_sheet, bootstyle="success-outline", width=4)
        b2 = tb.Button(file_frame, text="📂", command=app.add_file, bootstyle="primary-outline", width=4)
        b3 = tb.Button(file_frame, text="💾", command=app.save_file, bootstyle="primary-outline", width=4)
        b4 = tb.Button(file_frame, text="💽", command=app.save_all, bootstyle="primary-outline", width=4)
        
        b1.pack(side=tk.LEFT, padx=2)
        b2.pack(side=tk.LEFT, padx=2)
        b3.pack(side=tk.LEFT, padx=2)
        b4.pack(side=tk.LEFT, padx=2)
        
        ToolTip(b1, text="New Sheet")
        ToolTip(b2, text="Open File")
        ToolTip(b3, text="Save File")
        ToolTip(b4, text="Save All")
        
        # Sheet Frame
        sheet_frame = tb.Labelframe(toolbar, text="Sheet", padding=5)
        sheet_frame.pack(side=tk.LEFT, padx=5)

        s1 = tb.Button(sheet_frame, text="❌", command=app.close_current_tab, bootstyle="warning-outline", width=4)
        s2 = tb.Button(sheet_frame, text="🆚", command=app.open_diff_dialog, bootstyle="success-outline", width=4)
        s3 = tb.Button(sheet_frame, text="📊", command=app.show_column_stats, bootstyle="secondary-outline", width=4)
        # s4 = tb.Button(sheet_frame, text="📄🗄️", command=lambda: app.open_in_other_view(app.current_sheet_name), bootstyle="secondary-outline", width=4)
        s4 = tb.Button(sheet_frame, text="📖", command=lambda: app.toggle_split_view(), bootstyle="primary-outline", width=4)
        s5 = tb.Button(sheet_frame, text="☰", command=lambda: app.toggle_filter_bar(), bootstyle="primary-outline", width=4)
        
        s1.pack(side=tk.LEFT, padx=2)
        s2.pack(side=tk.LEFT, padx=2)
        s3.pack(side=tk.LEFT, padx=2)
        s4.pack(side=tk.LEFT, padx=2)
        s5.pack(side=tk.LEFT, padx=2)
        
        ToolTip(s1, text="Close Current Tab")
        ToolTip(s2, text="Compare Sheets (Diff)")
        ToolTip(s3, text="Column Statistics")
        ToolTip(s4, text="Move to Other View")
        ToolTip(s5, text="Toggle Filter Bar")

        # Edit Frame
        edit_frame = tb.Labelframe(toolbar, text="Edit", padding=5)
        edit_frame.pack(side=tk.LEFT, padx=5)

        e1 = tb.Button(edit_frame, text="↶", command=lambda: app.undo(), bootstyle="warning-outline", width=4)
        e2 = tb.Button(edit_frame, text="↷", command=lambda: app.redo(), bootstyle="warning-outline", width=4)
        e3 = tb.Button(edit_frame, text="⚡", command=app.autofill_selection, bootstyle="info-outline", width=4)
        e4 = tb.Button(edit_frame, text="┆", command=app.change_separator, bootstyle="warning-outline", width=4)

        e1.pack(side=tk.LEFT, padx=2)
        e2.pack(side=tk.LEFT, padx=2)
        e3.pack(side=tk.LEFT, padx=2)
        e4.pack(side=tk.LEFT, padx=2)

        ToolTip(e1, text="Undo Last Action")
        ToolTip(e2, text="Redo Last Action")
        ToolTip(e3, text="Autofill Selection")
        ToolTip(e4, text="Change CSV Separator")

        # Formula Frame
        formula_frame = tb.Labelframe(toolbar, text="Formulas", padding=5)
        formula_frame.pack(side=tk.LEFT, padx=5)

        f1 = tb.Button(formula_frame, text="Σ", command=lambda: app.insert_formula_template("=SUM(A1:A10)"), bootstyle="info-outline", width=4)
        f2 = tb.Button(formula_frame, text="</>", command=lambda: app.insert_formula_template('=PYTHON(A1*2+B1)'), bootstyle="info-outline", width=4)
        f3 = tb.Button(formula_frame, text="🧮", command=app.recalculate_all, bootstyle="success-outline", width=4)

        f1.pack(side=tk.LEFT, padx=2)
        f2.pack(side=tk.LEFT, padx=2)
        f3.pack(side=tk.LEFT, padx=2)

        ToolTip(f1, text="Insert SUM Template")
        ToolTip(f2, text="Insert Python Formula")
        ToolTip(f3, text="Recalculate All Formulas")

        # Clean Frame
        data_clean_frame = tb.Labelframe(toolbar, text="Clean", padding=5)
        data_clean_frame.pack(side=tk.LEFT, padx=5)

        c1 = tb.Button(data_clean_frame, text="🛸", command=app.clean_whitespace, bootstyle="warning-outline", width=4)
        c2 = tb.Button(data_clean_frame, text="🧹", command=app.clean_nan, bootstyle="warning-outline", width=4)

        c1.pack(side=tk.LEFT, padx=2)
        c2.pack(side=tk.LEFT, padx=2)

        ToolTip(c1, text="Trim Whitespace")
        ToolTip(c2, text="Replace NaN/Null Values")

        # MDT Frame
        MDT_frame = tb.Labelframe(toolbar, text="MDT", padding=5)
        MDT_frame.pack(side=tk.LEFT, padx=5)

        # For Checkbuttons, we can also add Tooltips!
        m_toggle = tb.Checkbutton(MDT_frame, text="💡", variable=app.suggestion_mode_var, bootstyle="square-toggle", command=app.toggle_suggestion_mode)
        m1 = tb.Button(MDT_frame, text="🝖", command=app.run_component_graphs, bootstyle="success-outline", width=4)
        m2 = tb.Button(MDT_frame, text="🜉", command=app.run_component_graphs_coord, bootstyle="success-outline", width=4)
        m3 = tb.Button(MDT_frame, text="ℹ️", command=UIComponents.open_help, bootstyle="info-outline", width=4)

        m_toggle.pack(side=tk.LEFT, padx=5)
        m1.pack(side=tk.LEFT, padx=2)
        m2.pack(side=tk.LEFT, padx=2)
        m3.pack(side=tk.LEFT, padx=2)

        ToolTip(m_toggle, text="Enable/Disable MDT Suggestions")
        ToolTip(m1, text="Run Component Graphs (Conta)")
        ToolTip(m2, text="Run Component Graphs (Coord)")
        ToolTip(m3, text="Open Help Documentation")
        
        settings_frame = tb.Labelframe(toolbar, text="Dark", padding=5)
        settings_frame.pack(side=tk.RIGHT, padx=5)
        app.dark_mode_var = tk.BooleanVar(value=False)
        tb.Checkbutton(settings_frame, text="🌙", variable=app.dark_mode_var, bootstyle="round-toggle", command=app.toggle_dark_mode).pack(side=tk.LEFT, padx=5)
    
    @staticmethod
    def create_formula_bar(root, app):
        """Create the formula bar."""
        formula_frame = tb.Frame(root)
        formula_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        tb.Label(formula_frame, text="fx", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        
        app.formula_var = tk.StringVar()
        app.formula_entry = tb.Entry(formula_frame, textvariable=app.formula_var, font=("Consolas", 10))
        app.formula_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        app.formula_entry.bind('<Return>', app._on_formula_enter)
        app.formula_entry.bind('<FocusOut>', app._on_formula_focus_out)
        
        # Auto calculation toggle
        app.auto_calc_var = tk.BooleanVar(value=False)
        calc_check = tb.Checkbutton(
            formula_frame, 
            text="Auto Calc", 
            variable=app.auto_calc_var,
            bootstyle="success-round-toggle"
        )
        calc_check.pack(side=tk.LEFT, padx=5)
        
        # Calculate button
        app.calc_button = tb.Button(
            formula_frame, 
            text="🧮 Calc", 
            command=app.calculate_current_cell,
            bootstyle="success-outline",
            width=8
        )
        app.calc_button.pack(side=tk.LEFT, padx=2)
    
    @staticmethod
    def create_status_bar(root, app):
        """Create the status bar."""
        app.status_frame = tk.Frame(root, relief=tk.SUNKEN, bd=1)
        app.status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        app.left_status = tk.Label(app.status_frame, text= LEFT_STATUS, anchor="w")
        app.left_status.pack(side=tk.LEFT, padx=5)
        
        app.report_label = tk.Label(app.status_frame, text="✉︎ Report Issue", fg="blue", cursor="hand2", anchor="w")
        app.report_label.pack(side=tk.LEFT, padx=10)
        ToolTip(app.report_label , text=f"Mail to: {REPORT_EMAIL}", position = "top")
        app.report_label.bind("<Button-1>", lambda e: webbrowser.open(f"mailto:{REPORT_EMAIL}?subject=Spreadsheet Editor Issue"))
        
        app.select_status = tk.Label(app.status_frame, text="Ready", anchor="w")
        app.select_status.pack(side=tk.LEFT, padx=5)
        
        app.cell_info = tk.Label(app.status_frame, text="", anchor="e")
        app.cell_info.pack(side=tk.RIGHT, padx=5)
        
        app.right_status = tk.Label(app.status_frame, text="Ready", anchor="e")
        app.right_status.pack(side=tk.RIGHT, padx=5)
    
    @staticmethod
    def open_help():
        """Open MDT help."""
        webbrowser.open(MDT_HELP_URL)
    
    @staticmethod
    def open_app_help():
        """Open app help."""
        try:
            os.startfile(USER_GUIDE_PATH)
        except:
            messagebox.showinfo("Help", "Help documentation not found")
    
    @staticmethod
    def show_about():
        """Show about dialog."""
        about_text = ABOUT_TEXT
        
        messagebox.showinfo("About", about_text)
        