"""
mixins/prefs_mixin.py  —  User Preferences, Column Widths, Themes, File Utilities
==================================================================================

FEATURES IN THIS FILE
----------------------
1. UserPreferences  — persistent JSON settings (font size, default sep,
                       startup file, confirm-on-close, row numbers, etc.)
2. Column width presets  — Wide / Normal / Compact / Fit-to-content
3. Custom theme picker   — all ttkbootstrap themes with live preview
4. Copy full file path   — like Notepad++ "Copy Full Path"
5. Open containing folder — like Notepad++ "Open Containing Folder"

INTEGRATION — add to class:
    from mixins.prefs_mixin import PrefsMixin
    class TableEditor(UndoMixin, CFMixin, FormulaMixin, SplitViewMixin, PrefsMixin):

CALL IN __init__ (after _setup_ui):
    self._load_prefs()          # loads prefs and applies startup settings

ADD TO MENUS in ui_components.py:
    # File menu:
    file_menu.add_command(label="Copy Full Path",           command=app.copy_full_path)
    file_menu.add_command(label="Open Containing Folder",   command=app.open_containing_folder)

    # View menu:
    view_menu.add_separator()
    view_menu.add_command(label="Column Width: Wide",       command=lambda: app.set_col_width_preset("wide"))
    view_menu.add_command(label="Column Width: Normal",     command=lambda: app.set_col_width_preset("normal"))
    view_menu.add_command(label="Column Width: Compact",    command=lambda: app.set_col_width_preset("compact"))
    view_menu.add_command(label="Column Width: Fit Content",command=lambda: app.set_col_width_preset("fit"))

    # Settings / Tools menu:
    settings_menu.add_command(label="Preferences...",       command=app.open_preferences)
    settings_menu.add_command(label="Theme Picker...",      command=app.open_theme_picker)
"""

import os
import json
import subprocess
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
# import tksheet
from config import APP_NAME
#  Preference defaults 
PREF_DEFAULTS = {
    "theme":              "flatly",      # ttkbootstrap theme name
    "dark_mode":          False,
    "col_width":          120,           # default column width in pixels
    "font_size":          10,            # tksheet font size (future use)
    "default_sep":        ",",           # fallback separator for new files
    "confirm_on_close":   True,          # ask before closing dirty tabs
    "show_row_numbers":   True,          # show tksheet row index
    "max_recent":         10,
    "startup_file":       "",            # auto-open on launch
    "autosave_mins":      0,             # 0 = disabled
    "filter_bar_visible": False,         # row filter bar shown by default
    "sheet_theme_light":  "light blue",   # tksheet theme when app is light
    "sheet_theme_dark":   "dark",         # tksheet theme when app is dark
}

# All ttkbootstrap themes grouped by style
THEMES_LIGHT = [
    "flatly", "litera", "minty", "lumen", "sandstone",
    "yeti", "pulse", "united", "morph", "journal", "simplex",
]

THEMES_DARK = [
    "darkly", "cyborg", "superhero", "solar", "slate",
    "cosmo", "vapor",
]

TKSHEET_THEMES = [
    "light blue", "light green", "dark", "black",
    "green", "blue", "purple", "dark blue",
]

class PrefsMixin:
    """Mixin: user preferences, column widths, themes, file path utilities."""

    #  Prefs path 
    def get_appdata_path(self):
        # 'Roaming' is best for settings that follow the user
        app_data = os.getenv('APPDATA') 
        appdata_folder = os.path.join(app_data, APP_NAME)
        if not os.path.exists(appdata_folder):
            os.makedirs(appdata_folder)
        return appdata_folder #os.path.join(app_folder, "settings.json")
    
    def _prefs_path(self) -> str:
        """Return full path to preferences JSON file."""
        pathapp = self.get_appdata_path() 
        os.makedirs(pathapp, exist_ok=True) 
        return os.path.join(pathapp, "preferences.json")

    #  Load / save prefs 

    def _load_prefs(self):
        """
        Load preferences from disk and store in self._prefs.
        Apply startup settings immediately.
        Call once from __init__ after _setup_ui().
        """
        self._prefs = dict(PREF_DEFAULTS)
        try:
            with open(self._prefs_path(), "r", encoding="utf-8") as f:
                saved = json.load(f)
            self._prefs.update(saved)
        except Exception:
            pass   # file missing or corrupt — defaults are fine

        # Apply immediately
        self._apply_prefs_to_ui()

    def _save_prefs(self):
        """Persist self._prefs to disk."""
        try:
            with open(self._prefs_path(), "w", encoding="utf-8") as f:
                json.dump(self._prefs, f, indent=2)
        except Exception as e:
            messagebox.showwarning("Preferences", f"Could not save preferences:\n{e}")

    def get_pref(self, key, default=None):
        """Read one preference value. Safe to call before _load_prefs."""
        prefs = getattr(self, '_prefs', PREF_DEFAULTS)
        return prefs.get(key, PREF_DEFAULTS.get(key, default))

    def set_pref(self, key, value, save=True):
        """Set one preference and optionally persist."""
        if not hasattr(self, '_prefs'):
            self._prefs = dict(PREF_DEFAULTS)
        self._prefs[key] = value
        if save:
            self._save_prefs()

    def _apply_prefs_to_ui(self):
        """Apply loaded prefs to the live UI. Called after load and after Apply."""
        prefs = getattr(self, '_prefs', PREF_DEFAULTS)

        # Theme
        theme = prefs.get("theme", "flatly")
        try:
            self.style.theme_use(theme)
        except Exception:
            pass

        # Dark mode checkbox sync
        dark = prefs.get("dark_mode", False)
        var = getattr(self, 'dark_mode_var', None)
        if var is not None:
            var.set(dark)

        # Column width
        col_w = prefs.get("col_width", 120)
        self._apply_col_width_to_all(col_w)
        
        # pply sheet theme 
        self._apply_sheet_theme_from_prefs()
        

    def _apply_sheet_theme_from_prefs(self):
        """
        Apply the sheet grid theme from preferences.
        Called from _apply_prefs_to_ui() and from toggle_dark_mode().
        Reads sheet_theme_light / sheet_theme_dark from prefs depending on
        whether dark mode is currently active.
        """
        prefs = getattr(self, '_prefs', PREF_DEFAULTS)
        dark  = getattr(self, 'dark_mode', False)
    
        if dark:
            sheet_theme = prefs.get("sheet_theme_dark", "dark")
        else:
            sheet_theme = prefs.get("sheet_theme_light", "light blue")
    
        # Apply to all open sheet tabs
        try:
            for tab_id in self.sheet_notebook.tabs():
                tab_frame = self.sheet_notebook.nametowidget(tab_id)
                for widget in tab_frame.winfo_children():
                    import tksheet
                    if isinstance(widget, tksheet.Sheet):
                        widget.set_options(theme=sheet_theme, align="center")
                        widget.refresh()
        except Exception:
            pass
        
    #  Preferences dialog 
    def open_preferences(self):
        """Open the Preferences dialog."""
        if not hasattr(self, '_prefs'):
            self._load_prefs()

        dlg = tb.Toplevel(self.root)
        dlg.title("Preferences")
        dlg.geometry("520x720")
        dlg.resizable(False, False)
        dlg.grab_set()

        # Header
        hdr = tb.Frame(dlg, bootstyle="primary")
        hdr.pack(fill="x")
        tb.Label(hdr, text="  Preferences",
                 font=("Segoe UI", 11, "bold"),
                 bootstyle="inverse-primary").pack(side="left", pady=8)

        body = tb.Frame(dlg)
        body.pack(fill="both", expand=True, padx=20, pady=12)

        def section(text):
            tb.Label(body, text=text,
                     font=("Segoe UI", 9, "bold"),
                     foreground="#2E75B6").pack(anchor="w", pady=(12, 2))
            tb.Separator(body, orient="horizontal").pack(fill="x", pady=(0, 6))

        def row(label, widget_fn):
            f = tb.Frame(body)
            f.pack(fill="x", pady=3)
            tb.Label(f, text=label, width=22, anchor="w",
                     font=("Segoe UI", 9)).pack(side="left")
            return widget_fn(f)

        p = self._prefs

        #  Appearance 
        section("Appearance")

        theme_var = tk.StringVar(value=p.get("theme", "flatly"))
        all_themes = THEMES_LIGHT + THEMES_DARK
        row("Theme:", lambda f: tb.Combobox(
            f, textvariable=theme_var,
            values=all_themes, state="readonly", width=20
        ).pack(side="left"))

        dark_var = tk.BooleanVar(value=p.get("dark_mode", False))
        row("Dark mode:", lambda f: tb.Checkbutton(
            f, variable=dark_var, bootstyle="round-toggle"
        ).pack(side="left"))
        #---------------
        sheet_light_var = tk.StringVar(value=p.get("sheet_theme_light", "light blue"))
        row("Sheet theme (light):", lambda f: tb.Combobox(
            f, textvariable=sheet_light_var,
            values=TKSHEET_THEMES, state="readonly", width=18
        ).pack(side="left"))
        
        sheet_dark_var = tk.StringVar(value=p.get("sheet_theme_dark", "dark"))
        row("Sheet theme (dark):", lambda f: tb.Combobox(
            f, textvariable=sheet_dark_var,
            values=TKSHEET_THEMES, state="readonly", width=18
        ).pack(side="left"))
        #  Sheet display 
        section("Sheet Display")

        colw_var = tk.IntVar(value=p.get("col_width", 120))
        def _colw_row(f):
            tb.Scale(f, from_=60, to=300, variable=colw_var,
                     orient="horizontal", length=160,
                     bootstyle="primary").pack(side="left", padx=(0,8))
            tb.Label(f, textvariable=colw_var, width=4).pack(side="left")
        row("Default column width:", _colw_row)

        rownums_var = tk.BooleanVar(value=p.get("show_row_numbers", True))
        row("Show row numbers:", lambda f: tb.Checkbutton(
            f, variable=rownums_var, bootstyle="round-toggle"
        ).pack(side="left"))

        #  Behaviour 
        section("Behaviour")

        confirm_var = tk.BooleanVar(value=p.get("confirm_on_close", True))
        row("Confirm on close:", lambda f: tb.Checkbutton(
            f, variable=confirm_var, bootstyle="round-toggle"
        ).pack(side="left"))

        sep_var = tk.StringVar(value=p.get("default_sep", ","))
        sep_choices = [",", ";", "\\t", "|", "\\s+"]
        row("Default separator:", lambda f: tb.Combobox(
            f, textvariable=sep_var,
            values=sep_choices, width=8
        ).pack(side="left"))

        recent_var = tk.IntVar(value=p.get("max_recent", 10))
        row("Max recent files:", lambda f: tb.Spinbox(
            f, from_=5, to=30, textvariable=recent_var, width=5
        ).pack(side="left"))

        autosave_var = tk.IntVar(value=p.get("autosave_mins", 0))
        def _auto_row(f):
            tb.Spinbox(f, from_=0, to=60, textvariable=autosave_var,
                       width=5).pack(side="left", padx=(0,6))
            tb.Label(f, text="min  (0 = off)", font=("Segoe UI", 8),
                     foreground="#888888").pack(side="left")
        row("Auto-save every:", _auto_row)

        #  Startup 
        section("Startup")

        startup_var = tk.StringVar(value=p.get("startup_file", ""))
        def _startup_row(f):
            e = tb.Entry(f, textvariable=startup_var, width=28)
            e.pack(side="left", padx=(0, 4))
            tb.Button(f, text="Browse",
                      command=lambda: startup_var.set(
                          tk.filedialog.askopenfilename() or startup_var.get()
                      ),
                      bootstyle="outline-secondary", width=8).pack(side="left")
        row("Auto-open file:", _startup_row)

        #  Buttons 
        btns = tb.Frame(dlg)
        btns.pack(pady=12)

        def apply_prefs():
            self._prefs.update({
                "theme":            theme_var.get(),
                "dark_mode":        dark_var.get(),
                "col_width":        colw_var.get(),
                "show_row_numbers": rownums_var.get(),
                "confirm_on_close": confirm_var.get(),
                "default_sep":      sep_var.get(),
                "max_recent":       recent_var.get(),
                "autosave_mins":    autosave_var.get(),
                "startup_file":     startup_var.get(),
                "sheet_theme_light": sheet_light_var.get(),
                "sheet_theme_dark":  sheet_dark_var.get(),
            })
            self._save_prefs()
            self._apply_prefs_to_ui()
            self.set_status("Preferences saved and applied.")

        def ok():
            apply_prefs()
            dlg.destroy()

        tb.Button(btns, text="OK",
                  command=ok, bootstyle="success", width=12).pack(side="left", padx=6)
        tb.Button(btns, text="Apply",
                  command=apply_prefs, bootstyle="primary", width=12).pack(side="left", padx=6)
        tb.Button(btns, text="Cancel",
                  command=dlg.destroy, bootstyle="secondary", width=12).pack(side="left", padx=6)
        tb.Button(btns, text="Reset Defaults",
                  command=lambda: (self._prefs.update(PREF_DEFAULTS),
                                   self._save_prefs(),
                                   self._apply_prefs_to_ui(),
                                   dlg.destroy()),
                  bootstyle="outline-danger", width=14).pack(side="left", padx=6)

    #  Theme picker 

    def open_theme_picker(self):
        """Popup with live theme preview — click any theme to apply instantly."""
        if not hasattr(self, '_prefs'):
            self._load_prefs()

        win = tb.Toplevel(self.root)
        win.title("Theme Picker")
        win.geometry("360x460")
        win.resizable(False, True)
        win.grab_set()

        hdr = tb.Frame(win, bootstyle="primary")
        hdr.pack(fill="x")
        tb.Label(hdr, text="  Theme Picker — click to preview",
                 font=("Segoe UI", 10, "bold"),
                 bootstyle="inverse-primary").pack(side="left", pady=6)

        current_var = tk.StringVar(value=self._prefs.get("theme", "flatly"))

        #  Light themes 
        body = tb.Frame(win)
        body.pack(fill="both", expand=True, padx=12, pady=8)

        tb.Label(body, text="Light themes",
                 font=("Segoe UI", 9, "bold"),
                 foreground="#555555").pack(anchor="w", pady=(4, 2))
        light_frame = tb.Frame(body)
        light_frame.pack(fill="x")

        tb.Label(body, text="Dark themes",
                 font=("Segoe UI", 9, "bold"),
                 foreground="#555555").pack(anchor="w", pady=(10, 2))
        dark_frame = tb.Frame(body)
        dark_frame.pack(fill="x")

        def _preview(theme):
            current_var.set(theme)
            try:
                self.style.theme_use(theme)
                # Sync dark_mode state so sheet theme picks the right variant
                self.dark_mode = theme in THEMES_DARK
                # Apply sheet theme from prefs (dark or light variant)
                if hasattr(self, '_apply_sheet_theme_from_prefs'):
                    self._apply_sheet_theme_from_prefs()
                self.set_status(f"Theme preview: {theme}")
            except Exception as e:
                self.set_status(f"Theme error: {e}")

        def _make_btn(parent, theme):
            is_dark = theme in THEMES_DARK
            btn = tk.Button(
                parent, text=theme,
                font=("Segoe UI", 8),
                bg="#2C2C2C" if is_dark else "#F5F5F5",
                fg="white" if is_dark else "#333333",
                relief="flat", bd=1,
                padx=8, pady=4,
                cursor="hand2",
                command=lambda t=theme: _preview(t)
            )
            return btn

        # Arrange in wrapping grid
        for i, t in enumerate(THEMES_LIGHT):
            btn = _make_btn(light_frame, t)
            btn.grid(row=i // 3, column=i % 3, padx=3, pady=3, sticky="ew")
        for i in range(3):
            light_frame.columnconfigure(i, weight=1)

        for i, t in enumerate(THEMES_DARK):
            btn = _make_btn(dark_frame, t)
            btn.grid(row=i // 3, column=i % 3, padx=3, pady=3, sticky="ew")
        for i in range(3):
            dark_frame.columnconfigure(i, weight=1)

        # Buttons
        btns = tb.Frame(win)
        btns.pack(pady=8)

        def save_theme():
            chosen = current_var.get()
            is_dark = chosen in THEMES_DARK
            self.set_pref("theme",     chosen)
            self.set_pref("dark_mode", is_dark)
            # Sync the dark_mode_var checkbox in the toolbar
            var = getattr(self, 'dark_mode_var', None)
            if var is not None:
                var.set(is_dark)
            self.dark_mode = is_dark
            self.set_status(f"Theme saved: {chosen}")
            win.destroy()

        def cancel():
            # Restore previous theme
            prev = self._prefs.get("theme", "flatly")
            try:
                self.style.theme_use(prev)
            except Exception:
                pass
            win.destroy()

        tb.Button(btns, text="Save Theme",
                  command=save_theme,
                  bootstyle="success", width=12).pack(side="left", padx=6)
        tb.Button(btns, text="Cancel",
                  command=cancel,
                  bootstyle="secondary", width=12).pack(side="left", padx=6)

    #  Column width presets 

    def set_col_width_preset(self, preset: str):
        """
        Apply a column width preset to all columns in all visible sheets.

        preset: "wide" | "normal" | "compact" | "fit"
        """
        widths = {"wide": 200, "normal": 120, "compact": 80}

        if preset == "fit":
            self._fit_columns_to_content()
            return

        w = widths.get(preset, 120)
        self._apply_col_width_to_all(w)
        self.set_pref("col_width", w)
        self.set_status(f"Column width: {preset} ({w}px)")

    def _apply_col_width_to_all(self, width: int):
        """Set all columns in the current sheet to the given pixel width."""
        sheet = self.get_current_sheet()
        if not sheet:
            return
        n_cols = sheet.total_columns()
        for c in range(n_cols):
            sheet.column_width(column=c, width=width)
        sheet.refresh()

    def _fit_columns_to_content(self):
        """
        Auto-size each column to fit its widest content.
        Uses a simple character-width heuristic (no font metrics needed).
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return

        headers = sheet.headers() or []
        data    = sheet.get_sheet_data()
        n_cols  = sheet.total_columns()

        # Character-width heuristic: ~7px per char, min 60, max 400
        CHAR_W = 10
        PAD    = 5

        for c in range(n_cols):
            max_chars = len(str(headers[c])) if c < len(headers) else 4
            for row in data:
                if c < len(row):
                    max_chars = max(max_chars, len(str(row[c])))
            w = max(60, min(400, max_chars * CHAR_W + PAD))
            sheet.column_width(column=c, width=w)

        sheet.refresh()
        self.set_status("Columns fitted to content")

    #  File path utilities 

    def copy_full_path(self):
        """Copy the full file path of the current tab to clipboard."""
        meta = self.tab_meta.get(self.current_sheet_name, {})
        fp   = meta.get("file_path") or self.current_file

        if not fp:
            self.set_status("No file — nothing to copy")
            messagebox.showinfo(
                "Copy Full Path",
                "This tab has not been saved to a file yet."
            )
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(fp)
        self.root.update()   # required on some platforms
        self.set_status(f"Copied: {fp}")

    def copy_file_name(self):
        """Copy just the filename (no path) to clipboard."""
        meta = self.tab_meta.get(self.current_sheet_name, {})
        fp   = meta.get("file_path") or self.current_file
        if not fp:
            self.set_status("No file")
            return
        name = os.path.basename(fp)
        self.root.clipboard_clear()
        self.root.clipboard_append(name)
        self.root.update()
        self.set_status(f"Copied filename: {name}")

    def open_containing_folder(self):
        """Open the folder containing the current file in Windows Explorer."""
        meta = self.tab_meta.get(self.current_sheet_name, {})
        fp   = meta.get("file_path") or self.current_file

        if not fp:
            self.set_status("No file nothing to open")
            messagebox.showinfo(
                "Open Containing Folder",
                "This tab has not been saved to a file yet."
            )
            return

        folder = os.path.dirname(fp)

        if not os.path.isdir(folder):
            self.set_status(f"Folder not found: {folder}")
            return

        try:
            if os.name == "nt":                          # Windows
                # subprocess.Popen(f'explorer /select,"{fp}"')
                subprocess.Popen(['explorer', '/select,', os.path.normpath(fp)])
            elif os.sys.platform == "darwin":            # macOS
                subprocess.Popen(["open", "-R", os.path.normpath(fp)])
            else:                                        # Linux
                subprocess.Popen(["xdg-open", folder])
                
            self.set_status(f"Opened: {folder}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")
