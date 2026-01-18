"""
UI Components - Menus, Toolbars, Status Bar
"""
import os
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
import webbrowser
from config import APP_NAME, APP_VERSION, REPORT_EMAIL, MDT_HELP_URL, USER_GUIDE_PATH, ABOUT_TEXT, LEFT_STATUS


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
        file_menu.add_command(label="New", command=app.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Open", command=app.open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=app.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As", command=app.save_file_as)
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
        edit_menu.add_command(label="Undo", command=lambda: app.get_current_sheet().undo(), accelerator="Ctrl+Z")
        edit_menu.add_command(label="Redo", command=lambda: app.get_current_sheet().redo(), accelerator="Ctrl+Y")
        edit_menu.add_separator()
        # edit_menu.add_command(label="Find", command=lambda: app.get_current_sheet().find(), accelerator="Ctrl+F")
        
        # Insert Menu
        insert_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Insert", menu=insert_menu)
        insert_menu.add_command(label="Insert Row", command=app.insert_row)
        insert_menu.add_command(label="Insert Column", command=app.add_column)
        insert_menu.add_separator()
        insert_menu.add_command(label="Insert Sheet", command=app.add_new_sheet)
        
        # Formulas Menu
        UIComponents._create_formula_menu(menubar, app)
        
        # Tools Menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        app.suggestion_mode_var = tk.BooleanVar(value=False)
        tools_menu.add_checkbutton(label="MDT Suggestion Mode", variable=app.suggestion_mode_var, command=app.toggle_suggestion_mode)
        tools_menu.add_command(label="Plot COMP Graph Conta", command=app.run_component_graphs)
        tools_menu.add_command(label="Plot COMP Graph Coord", command=app.run_component_graphs_coord)
        tools_menu.add_separator()
        tools_menu.add_command(label="Change Separator", command=app.change_separator)
        
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
        tb.Button(file_frame, text="📄 New", command=app.new_file, bootstyle="secondary-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(file_frame, text="📂 Open", command=app.open_file, bootstyle="primary-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(file_frame, text="💾 Save", command=app.save_file, bootstyle="success-outline", width=8).pack(side=tk.LEFT, padx=2)
        
        # Edit Frame
        edit_frame = tb.Labelframe(toolbar, text="Edit", padding=5)
        edit_frame.pack(side=tk.LEFT, padx=5)
        tb.Button(edit_frame, text="↶ Undo", command=lambda: app.get_current_sheet().undo(), bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(edit_frame, text="↷ Redo", command=lambda: app.get_current_sheet().redo(), bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(edit_frame, text="⚡ Fill", command=app.autofill_selection, bootstyle="info-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(edit_frame, text="⨾ Sep", command=app.change_separator, bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)
        # tools_menu.add_command(label="Change Separator", command=app.change_separator)
        # Formula Frame
        formula_frame = tb.Labelframe(toolbar, text="Formulas", padding=5)
        formula_frame.pack(side=tk.LEFT, padx=5)
        tb.Button(formula_frame, text="Σ SUM", command=lambda: app.insert_formula_template("=SUM(A1:A10)"), bootstyle="info-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(formula_frame, text="🐍 Py", command=lambda: app.insert_formula_template('=PYTHON(A1*2+B1)'), bootstyle="info-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(formula_frame, text="fx Calc", command=app.recalculate_all, bootstyle="success-outline", width=8).pack(side=tk.LEFT, padx=2)
        
        # Clean Frame
        data_clean_frame = tb.Labelframe(toolbar, text="Clean cells (No Undo)", padding=5)
        data_clean_frame.pack(side=tk.LEFT, padx=5)
        tb.Button(data_clean_frame, text="🧹 Space", command=app.clean_whitespace, bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(data_clean_frame, text="🧹 NaN", command=app.clean_nan,          bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)
        # tb.Button(data_clean_frame, text="Whitespace", command=app.clean_whitespace, bootstyle="warning-outline", width=8).pack(side=tk.LEFT, padx=2)

        # MDT Frame
        MDT_frame = tb.Labelframe(toolbar, text="MDT", padding=5)
        MDT_frame.pack(side=tk.LEFT, padx=5)
        tb.Checkbutton(MDT_frame, text="Suggestion", variable=app.suggestion_mode_var, bootstyle="square-toggle", command=app.toggle_suggestion_mode).pack(side=tk.LEFT, padx=5)
        tb.Button(MDT_frame, text="🝖 Conta", command=app.run_component_graphs, bootstyle="success-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(MDT_frame, text="🜉 Coord", command=app.run_component_graphs_coord, bootstyle="success-outline", width=8).pack(side=tk.LEFT, padx=2)
        tb.Button(MDT_frame, text="ℹ️ Help", command=UIComponents.open_help, bootstyle="info-outline", width=8).pack(side=tk.LEFT, padx=2)

        settings_frame = tb.Labelframe(toolbar, text="Settings", padding=5)
        settings_frame.pack(side=tk.RIGHT, padx=5)
        app.dark_mode_var = tk.BooleanVar(value=False)
        tb.Checkbutton(settings_frame, text="🌙 Dark", variable=app.dark_mode_var, bootstyle="round-toggle", command=app.toggle_dark_mode).pack(side=tk.LEFT, padx=5)
    
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
            text="📊 Calc", 
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
        
        app.report_label = tk.Label(app.status_frame, text="🔧 Report Issue", fg="blue", cursor="hand2", anchor="w")
        app.report_label.pack(side=tk.LEFT, padx=10)
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
