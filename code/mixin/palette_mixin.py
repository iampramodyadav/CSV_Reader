"""
mixins/palette_mixin.py  —  Command Palette  (Ctrl+P)
=====================================================

A floating search-and-run popup listing every app action.
Type to fuzzy-filter, arrow keys to navigate, Enter to run.

Features
--------
* All registered commands searchable by name
* Fuzzy match: type "fre" → finds "Find & Replace All"
* Keyboard-only: Up/Down/Enter/Escape
* Shows keyboard shortcut hint next to each command
* Auto-closes after running
* Commands are registered via a simple decorator pattern
* Section headers group commands visually

INTEGRATION
-----------
1.  In table_editor.py imports:
        from mixins.palette_mixin import PaletteMixin

2.  Add to class declaration:
        class TableEditor(..., PaletteMixin):

3.  In __init__, after _setup_ui and _setup_plugin_menu, call:
        self._register_palette_commands()

4.  In main_app.py:
        root.bind("<Control-p>", lambda e: app.open_command_palette())

5.  In ui_components.py (View or Help menu):
        view_menu.add_command(
            label="Command Palette   Ctrl+P",
            command=app.open_command_palette
        )
"""

import tkinter as tk
# from tkinter import ttk
# import ttkbootstrap as tb


class PaletteMixin:
    """Mixin: command palette (Ctrl+P) for instant access to any app action."""

    #  Command registry 

    def _register_palette_commands(self):
        """
        Build the full command list.
        Each entry: (section, display_name, shortcut_hint, callable)

        Call once from __init__ after all setup is done.
        """
        self._palette_commands = []

        def add(section, name, shortcut, fn):
            self._palette_commands.append(
                (section, name, shortcut, fn)
            )

        #  File 
        add("File", "New File",                     "Ctrl+N",        self.new_file)
        add("File", "Open File",                    "Ctrl+O",        self.open_file)
        add("File", "Add File to Session",          "Ctrl+Shift+O",  self.add_file)
        add("File", "Save",                         "Ctrl+S",        self.save_file)
        add("File", "Save As",                      "",              self.save_file_as)
        add("File", "Copy Full Path",               "",              self.copy_full_path)
        add("File", "Copy File Name",               "",              self.copy_file_name)
        add("File", "Open Containing Folder",       "",              self.open_containing_folder)

        #  Edit 
        add("Edit", "Undo",                         "Ctrl+Z",        self.undo)
        add("Edit", "Redo",                         "Ctrl+Y",        self.redo)
        add("Edit", "Trim Whitespace",              "",              self.clean_whitespace)
        add("Edit", "Fill Nulls",                   "",              self.clean_nan)
        add("Edit", "Autofill Selection",           "Ctrl+R",        self.autofill_selection)

        #  View 
        add("View", "Toggle Filter Bar",            "",              self.toggle_filter_bar)
        add("View", "Toggle Split View",            "Ctrl+Shift+V",  self.toggle_split_view)
        add("View", "Column Width: Wide",           "",              lambda: self.set_col_width_preset("wide"))
        add("View", "Column Width: Normal",         "",              lambda: self.set_col_width_preset("normal"))
        add("View", "Column Width: Compact",        "",              lambda: self.set_col_width_preset("compact"))
        add("View", "Column Width: Fit Content",    "",              lambda: self.set_col_width_preset("fit"))

        #  Sheet 
        add("Sheet", "Add New Sheet",               "",              self.add_new_sheet)
        add("Sheet", "Rename Sheet",                "",              self.rename_sheet)
        add("Sheet", "Close Current Tab",           "Ctrl+W",        self.close_current_tab)
        add("Sheet", "Column Stats",                "",              self.show_column_stats)
        add("Sheet", "Next Tab",                    "Ctrl+PageDown",      self.next_tab)
        add("Sheet", "Previous Tab",                "Ctrl+PageUp",        self.prev_tab)

        #  Formula 
        add("Formula", "Recalculate All",           "",              self.recalculate_all)
        add("Formula", "Calculate Current Cell",    "",              self.calculate_current_cell)
        add("Formula", "Formula Browser",           "",              self.show_formula_browser)

        #  Format 
        add("Format", "Conditional Formatting",     "",              self.open_cf_panel)
        add("Format", "Clear All CF Highlights",    "",              self.clear_cf_for_tab)

        #  Diff 
        add("Tools", "Compare Two Sheets",          "",              self.open_diff_dialog)

        #  Settings 
        add("Settings", "Preferences",              "",              self.open_preferences)
        add("Settings", "Theme Picker",             "",              self.open_theme_picker)
        add("Settings", "Change Separator",         "",              self.change_separator)

        #  Export 
        add("Export", "Export as Excel",            "",              self.export_to_excel)
        add("Export", "Export as CSV",              "",              self.export_to_csv)
        add("Export", "Export as Markdown",         "",              self.export_to_md)
        add("Export", "Export as HTML",             "",              self.export_to_html)
        add("Export", "Export as JSON",             "",              self.export_to_json)

        #  Plugins 
        add("Plugins", "Reload Plugins",            "",              self._reload_plugin_menu)
        add("Plugins", "Open Plugin File",          "",              lambda: self._open_plugin_file(
            getattr(self, '_plugin_file_path',
                    __import__('config').PLUGIN_FILE_PATH if hasattr(__import__('config'), 'PLUGIN_FILE_PATH') else '')
        ))

    def _add_palette_command(self, section: str, name: str,
                              shortcut: str, fn):
        """
        Register an extra command at runtime.
        Use this to add plugin-generated commands to the palette.
        """
        if not hasattr(self, '_palette_commands'):
            self._register_palette_commands()
        self._palette_commands.append((section, name, shortcut, fn))

    #  Open palette 

    def open_command_palette(self):
        """Open the command palette popup."""
        if not hasattr(self, '_palette_commands'):
            self._register_palette_commands()

        # Only one palette at a time
        if hasattr(self, '_palette_win') and self._palette_win:
            try:
                self._palette_win.focus_set()
                return
            except Exception:
                pass

        win = tk.Toplevel(self.root)
        win.overrideredirect(True)     # no title bar
        win.attributes("-topmost", True)
        self._palette_win = win

        # Position: centred horizontally, near top of window
        self.root.update_idletasks()
        rw = self.root.winfo_width()
        rx = self.root.winfo_rootx()
        ry = self.root.winfo_rooty()
        pw, ph = 520, 380
        wx = rx + (rw - pw) // 2
        wy = ry + 60
        win.geometry(f"{pw}x{ph}+{wx}+{wy}")

        #  Border frame (gives the floating panel its outline) 
        outer = tk.Frame(win, bg="#2E75B6", padx=1, pady=1)
        outer.pack(fill="both", expand=True)

        inner = tk.Frame(outer, bg="white")
        inner.pack(fill="both", expand=True)

        #  Search entry 
        search_frame = tk.Frame(inner, bg="white")
        search_frame.pack(fill="x", padx=8, pady=(8, 4))

        tk.Label(search_frame, text="⌨",
                 bg="white", fg="#2E75B6",
                 font=("Segoe UI", 14)).pack(side="left", padx=(0, 6))

        query_var = tk.StringVar()
        entry = tk.Entry(
            search_frame, textvariable=query_var,
            font=("Segoe UI", 11),
            relief="flat", bg="white",
            insertbackground="#2E75B6"
        )
        entry.pack(side="left", fill="x", expand=True)

        tk.Frame(inner, bg="#E0E0E0", height=1).pack(fill="x", padx=8)

        #  Results listbox 
        list_frame = tk.Frame(inner, bg="white")
        list_frame.pack(fill="both", expand=True, padx=4, pady=4)

        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        lb = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 9),
            selectmode="single",
            activestyle="none",
            relief="flat",
            bg="white",
            selectbackground="#2E75B6",
            selectforeground="white",
            borderwidth=0,
            highlightthickness=0,
        )
        scrollbar.config(command=lb.yview)
        lb.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        #  Footer 
        footer = tk.Frame(inner, bg="#F5F5F5", height=22)
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer,
            text="  ↑↓ navigate    Enter run    Esc close",
            bg="#F5F5F5", fg="#888888",
            font=("Segoe UI", 8)
        ).pack(side="left", pady=3)

        #  Populate + filter logic 
        all_cmds = self._palette_commands   # list of (section, name, hint, fn)
        displayed = []                       # parallel to listbox entries

        def _fuzzy_match(query: str, text: str) -> bool:
            """True if all chars of query appear in order in text (case-insensitive)."""
            q = query.lower()
            t = text.lower()
            it = iter(t)
            return all(c in it for c in q)

        def refresh(*_):
            q = query_var.get().strip()
            lb.delete(0, "end")
            displayed.clear()

            prev_section = None
            for section, name, hint, fn in all_cmds:
                if q and not _fuzzy_match(q, name) and not _fuzzy_match(q, section):
                    continue

                # Section header (only when no query — headers clutter filtered view)
                if not q and section != prev_section:
                    lb.insert("end", f"   {section} ")
                    lb.itemconfig(lb.size() - 1,
                                  fg="#888888", selectbackground="white",
                                  selectforeground="#888888")
                    displayed.append(None)   # sentinel — header, not runnable
                    prev_section = section

                # Command entry
                shortcut_part = f"  {hint}" if hint else ""
                label = f"  {name}{shortcut_part}"
                lb.insert("end", label)
                displayed.append(fn)

            # Auto-select first runnable item
            for i, fn in enumerate(displayed):
                if fn is not None:
                    lb.selection_set(i)
                    lb.see(i)
                    break

        def _run_selected():
            sel = lb.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx >= len(displayed) or displayed[idx] is None:
                return
            fn = displayed[idx]
            _close()
            try:
                fn()
            except Exception as exc:
                self.set_status(f"Command error: {exc}")

        def _close():
            self._palette_win = None
            win.destroy()

        def _on_key(event):
            if event.keysym == "Escape":
                _close()
                return "break"
            if event.keysym in ("Return", "KP_Enter"):
                _run_selected()
                return "break"
            if event.keysym == "Down":
                _move_selection(1)
                return "break"
            if event.keysym == "Up":
                _move_selection(-1)
                return "break"

        def _move_selection(delta: int):
            sel = lb.curselection()
            if not sel:
                return
            idx  = sel[0]
            size = lb.size()
            new  = idx + delta

            # Skip section headers
            while 0 <= new < size and displayed[new] is None:
                new += delta

            if 0 <= new < size:
                lb.selection_clear(0, "end")
                lb.selection_set(new)
                lb.see(new)

        query_var.trace_add("write", refresh)
        entry.bind("<Key>", _on_key)
        lb.bind("<Double-Button-1>", lambda e: _run_selected())
        lb.bind("<Return>", lambda e: _run_selected())

        # Close when focus leaves the window
        win.bind("<FocusOut>", lambda e: (
            _close() if win.focus_get() is None else None
        ))

        # Init
        refresh()
        entry.focus_set()

    #  Tab navigation (needed by palette) 

    def next_tab(self):
        """Switch to the next tab (wraps around)."""
        tabs = self.sheet_notebook.tabs()
        if not tabs:
            return
        cur = self.sheet_notebook.index(self.sheet_notebook.select())
        self.sheet_notebook.select((cur + 1) % len(tabs))

    def prev_tab(self):
        """Switch to the previous tab (wraps around)."""
        tabs = self.sheet_notebook.tabs()
        if not tabs:
            return
        cur = self.sheet_notebook.index(self.sheet_notebook.select())
        self.sheet_notebook.select((cur - 1) % len(tabs))
