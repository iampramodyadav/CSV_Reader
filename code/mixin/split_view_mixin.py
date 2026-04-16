"""
mixins/split_view_mixin.py
==========================
Side-by-side split view for TableEditor  —  like Notepad++ "Other View".

WHAT IT DOES
------------
Splits the main content area into two panes using a tk.PanedWindow.
Left pane = the existing self.sheet_notebook (all existing behaviour unchanged).
Right pane = a second read-only tksheet showing any sheet from workbook_sheets.

The right pane is READ-ONLY.  All editing still happens in the left pane.
This is a deliberate design choice — it avoids any conflict with the undo stack,
formula engine, mark_modified, and all other existing state.

The right pane auto-refreshes when:
  - The user picks a different sheet for it via the header dropdown
  - update_sheet_from_dataframe() is called (any df change in left pane)
  - mark_modified() fires (any cell edit in left pane)
  - The user switches tabs in the left pane

LAYOUT
------
┌─────────────────────────────────────────────────────────────────┐
│  Toolbar                                                        │
├──────────────────────────────┬──────────────────────────────────┤
│  LEFT PANE                   │  RIGHT PANE                     │
│  [Tab1] [Tab2] [Tab3]        │  ▾ sensors_v2  [↻] [✕]         │
│  ┌────────────────────────┐  │  ┌──────────────────────────┐   │
│  │  main sheet_notebook   │  │  │  read-only tksheet       │   │
│  │  (all existing tabs)   │  │  │  (any open sheet)        │   │
│  │                        │  │  │                          │   │
│  └────────────────────────┘  │  └──────────────────────────┘   │
├──────────────────────────────┴──────────────────────────────────┤
│  Formula bar                                                    │
│  Status bar                                                     │
└─────────────────────────────────────────────────────────────────┘

HOW TO INTEGRATE
----------------
1.  In table_editor.py imports:
        from mixins.split_view_mixin import SplitViewMixin

2.  Add to class declaration:
        class TableEditor(UndoMixin, CFMixin, FormulaMixin, SplitViewMixin):

3.  In _setup_ui(), replace the two lines:
        content_frame = tb.Frame(self.root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.sheet_notebook = ttk.Notebook(content_frame)
        self.sheet_notebook.pack(fill=tk.BOTH, expand=True)

    With:
        content_frame = tb.Frame(self.root)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._split_content_frame = content_frame          # ← store ref
        self.sheet_notebook = self._build_left_pane(content_frame)

4.  In update_sheet_from_dataframe(), add at the very end:
        self._refresh_right_pane()

5.  In mark_modified(), add at the very end (after apply_cf_rules):
        self._refresh_right_pane()

6.  In _on_sheet_change(), add at the very end:
        self._refresh_right_pane()

7.  In ui_components.py, add a menu item (View menu):
        view_menu.add_command(
            label="Open in Other View  (Ctrl+Shift+V)",
            command=app.toggle_split_view
        )

8.  In main_app.py, add shortcut:
        root.bind("<Control-Shift-V>", lambda e: app.toggle_split_view())

That is all.  No other existing methods need changing.
"""

import tkinter as tk
from tkinter import ttk
import ttkbootstrap as tb
import tksheet
# import pandas as pd


class SplitViewMixin:
    """
    Mixin: side-by-side split view.
    Provides toggle_split_view(), open_in_other_view(), and auto-refresh.
    """

    # ── State initialiser ─────────────────────────────────────────────────────

    def _init_split_state(self):
        """Create split view state dict. Safe to call multiple times."""
        if not hasattr(self, '_sv'):
            self._sv = {
                'active':        False,
                'paned':         None,   # tk.PanedWindow widget
                'right_frame':   None,   # tb.Frame containing right pane
                'right_sheet':   None,   # tksheet.Sheet widget (read-only)
                'right_header':  None,   # tb.Frame (top bar of right pane)
                'right_name':    None,   # sheet name currently shown in right
                'right_label':   None,   # tk.Label showing sheet name
                'sheet_combo':   None,   # tb.Combobox for sheet selection
            }

    # ── Called from _setup_ui ─────────────────────────────────────────────────

    def _build_left_pane(self, content_frame) -> ttk.Notebook:
        """
        Replace the simple notebook pack with a PanedWindow container.
        The PanedWindow starts with only the left pane; the right pane is
        added dynamically when split view is activated.

        Returns the ttk.Notebook so _setup_ui can assign self.sheet_notebook.
        """
        self._init_split_state()

        # PanedWindow fills content_frame — initially only left child visible
        paned = tk.PanedWindow(
            content_frame,
            orient=tk.HORIZONTAL,
            sashwidth=6,
            sashrelief="raised",
            sashcursor="",
            bg="#CCCCCC",
        )
        paned.pack(fill=tk.BOTH, expand=True)
        self._sv['paned'] = paned

        # Left frame holds the existing notebook
        left_frame = tb.Frame(paned)
        paned.add(left_frame, minsize=200, stretch="always")
        self._sv['left_frame'] = left_frame

        # Create the notebook inside the left frame
        notebook = ttk.Notebook(left_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        return notebook

    # ── Toggle ────────────────────────────────────────────────────────────────

    def toggle_split_view(self, sheet_name: str = None):
        """
        Open or close the right pane.

        If called with no argument, opens with a sheet-picker dialog.
        If called with a sheet_name, opens directly showing that sheet.
        If split view is already open, closes it.
        """
        self._init_split_state()

        if self._sv['active']:
            self._close_split_view()
        else:
            self._open_split_view(sheet_name)

    def open_in_other_view(self, sheet_name: str = None):
        """
        Open the given sheet (or current sheet) in the right pane.
        If split is already open, just switches the right pane content.
        """
        self._init_split_state()

        if sheet_name is None:
            sheet_name = self.current_sheet_name

        if self._sv['active']:
            self._set_right_pane_sheet(sheet_name)
        else:
            self._open_split_view(sheet_name)

    # ── Open ──────────────────────────────────────────────────────────────────

    def _open_split_view(self, sheet_name: str = None):
        """Build the right pane and add it to the PanedWindow."""
        self._init_split_state()

        if not self.workbook_sheets:
            from tkinter import messagebox
            messagebox.showinfo(
                "Split View", "Open at least one file first."
            )
            return

        # Pick which sheet to show — prefer the 'other' sheet, not the active one
        if sheet_name is None:
            names = list(self.workbook_sheets.keys())
            # Default to the second tab if it exists, else the first
            if len(names) > 1:
                sheet_name = names[1] if names[0] == self.current_sheet_name else names[0]
            else:
                sheet_name = names[0]

        # ── Build right frame ─────────────────────────────────────────────
        right_frame = tb.Frame(self._sv['paned'])
        self._sv['paned'].add(right_frame, minsize=200, stretch="always")
        self._sv['right_frame'] = right_frame

        # ── Right pane header bar ─────────────────────────────────────────
        hdr = tk.Frame(right_frame, bg="#E8F4FD", height=32)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        self._sv['right_header'] = hdr

        # "Other View:" label
        tk.Label(
            hdr, text=" Other View: ",
            bg="#E8F4FD", fg="#0C4A6E",
            font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(4, 0))

        # Sheet selector combobox
        combo_var = tk.StringVar(value=sheet_name)
        combo = ttk.Combobox(
            hdr,
            textvariable=combo_var,
            values=list(self.workbook_sheets.keys()),
            state="readonly",
            width=28,
            font=("Segoe UI", 9),
        )
        combo.pack(side="left", padx=4, pady=4)
        self._sv['sheet_combo'] = combo
        self._sv['combo_var']   = combo_var

        def _on_combo_change(event=None):
            chosen = combo_var.get()
            if chosen in self.workbook_sheets:
                self._set_right_pane_sheet(chosen)

        combo.bind("<<ComboboxSelected>>", _on_combo_change)

        # Refresh button
        tk.Button(
            hdr, text="↻",
            bg="#E8F4FD", fg="#0C4A6E",
            font=("Segoe UI", 10),
            relief="flat", cursor="hand2", padx=4,
            command=lambda: self._set_right_pane_sheet(combo_var.get())
        ).pack(side="left", padx=2)

        # Read-only badge
        tk.Label(
            hdr, text="[read-only]",
            bg="#E8F4FD", fg="#888888",
            font=("Segoe UI", 8, "italic")
        ).pack(side="left", padx=6)

        # Close button
        tk.Button(
            hdr, text="✕",
            bg="#E8F4FD", fg="#CC0000",
            font=("Segoe UI", 10, "bold"),
            relief="flat", cursor="hand2", padx=6,
            command=self._close_split_view
        ).pack(side="right", padx=4)

        # ── Right pane tksheet ────────────────────────────────────────────
        sheet_frame = tb.Frame(right_frame)
        sheet_frame.pack(fill=tk.BOTH, expand=True)

        right_sheet = tksheet.Sheet(
            sheet_frame,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            show_top_left=True,
            headers=None,
        )
        # Read-only: enable select and copy only
        right_sheet.enable_bindings((
            "single_select", "row_select", "column_select", "drag_select",
            "copy", "arrowkeys", "column_width_resize", "row_height_resize",
        ))
        right_sheet.pack(fill=tk.BOTH, expand=True)
        self._sv['right_sheet'] = right_sheet

        # Right-click menu for the right pane
        right_sheet.popup_menu_add_command(
            "Switch to this sheet (left pane)",
            lambda: self._switch_left_to_right_sheet()
        )
        right_sheet.popup_menu_add_command(
            "Close Other View",
            self._close_split_view
        )

        # ── Apply current dark/light theme ────────────────────────────────
        self._apply_theme_to_right_pane()

        # ── Populate with initial data ────────────────────────────────────
        self._sv['active']     = True
        self._sv['right_name'] = None   # force first populate
        self._set_right_pane_sheet(sheet_name)

        # Set equal split — 50/50
        self.root.update_idletasks()
        total_w = self._sv['paned'].winfo_width()
        if total_w > 10:
            self._sv['paned'].sash_place(0, total_w // 2, 0)

        self.set_status(f"Split view opened  —  right pane: {sheet_name}")

    # ── Close ─────────────────────────────────────────────────────────────────

    def _close_split_view(self):
        """Remove the right pane and restore full-width layout."""
        self._init_split_state()

        if not self._sv['active']:
            return

        rf = self._sv.get('right_frame')
        if rf:
            try:
                self._sv['paned'].forget(rf)
                rf.destroy()
            except Exception:
                pass

        self._sv['active']       = False
        self._sv['right_frame']  = None
        self._sv['right_sheet']  = None
        self._sv['right_header'] = None
        self._sv['right_name']   = None
        self._sv['right_label']  = None
        self._sv['sheet_combo']  = None
        self._sv['combo_var']    = None

        self.set_status("Split view closed")

    # ── Populate right pane ───────────────────────────────────────────────────

    def _set_right_pane_sheet(self, sheet_name: str):
        """
        Display the given sheet in the right pane.
        Safe to call at any time — checks everything before touching widgets.
        """
        self._init_split_state()

        if not self._sv['active']:
            return

        right_sheet = self._sv.get('right_sheet')
        if right_sheet is None:
            return

        if sheet_name not in self.workbook_sheets:
            return

        df = self.workbook_sheets[sheet_name]
        self._sv['right_name'] = sheet_name

        # Update combobox to match
        combo_var = self._sv.get('combo_var')
        if combo_var is not None:
            combo_var.set(sheet_name)

        # Update combobox values in case sheets were added/removed
        combo = self._sv.get('sheet_combo')
        if combo is not None:
            combo['values'] = list(self.workbook_sheets.keys())

        # Populate the sheet widget
        if df is None or df.empty:
            right_sheet.headers(["(empty)"])
            right_sheet.set_sheet_data([[""]])
        else:
            right_sheet.headers([str(c) for c in df.columns])
            right_sheet.set_sheet_data(df.astype(str).values.tolist())

        right_sheet.set_options(align="center")
        right_sheet.refresh()

    def _refresh_right_pane(self):
        """
        Re-populate the right pane with the latest data from workbook_sheets.

        Called automatically from:
          - update_sheet_from_dataframe()  (any df push)
          - mark_modified()               (any cell edit)
          - _on_sheet_change()            (tab switch)

        Only does work when split view is actually active.
        """
        self._init_split_state()

        if not self._sv['active']:
            return

        name = self._sv.get('right_name')
        if name and name in self.workbook_sheets:
            self._set_right_pane_sheet(name)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _switch_left_to_right_sheet(self):
        """
        Switch the left pane (main notebook) to the sheet currently shown
        in the right pane.  Convenience action from right-click menu.
        """
        self._init_split_state()

        name = self._sv.get('right_name')
        if not name:
            return

        # Find the tab in sheet_notebook whose text matches name
        for tab_id in self.sheet_notebook.tabs():
            if self.sheet_notebook.tab(tab_id, "text") == name:
                self.sheet_notebook.select(tab_id)
                self.set_status(f"Switched to: {name}")
                return

        self.set_status(f"Tab '{name}' not found in main notebook")

    def _apply_theme_to_right_pane(self):
        """Apply the current dark/light sheet theme to the right pane widget."""
        right_sheet = self._sv.get('right_sheet')
        if right_sheet is None:
            return

        # Read dark mode from existing app state
        dark = getattr(self, 'dark_mode', False)
        try:
            from config import SHEET_LIGHT_THEME, SHEET_DARK_THEME
            theme = SHEET_DARK_THEME if dark else SHEET_LIGHT_THEME
            right_sheet.set_options(theme=theme, align="center")
        except Exception:
            pass

    def update_right_pane_theme(self):
        """
        Call this from toggle_dark_mode() to keep right pane in sync with theme.
        Add one line at the end of toggle_dark_mode():
            self.update_right_pane_theme()
        """
        self._apply_theme_to_right_pane()
        right_sheet = self._sv.get('right_sheet')
        if right_sheet:
            right_sheet.refresh()

    # ── Sheet combo refresh on workbook change ────────────────────────────────

    def _sv_update_combo(self):
        """
        Refresh the right pane sheet selector list.
        Call from load_file() and close_current_tab() so newly added or
        removed sheets appear in the dropdown.

        Add these calls:
          load_file()          → after self.set_status(...)  add self._sv_update_combo()
          close_current_tab()  → after removing tab          add self._sv_update_combo()
        """
        self._init_split_state()

        if not self._sv['active']:
            return

        combo = self._sv.get('sheet_combo')
        if combo is not None:
            combo['values'] = list(self.workbook_sheets.keys())

        # If the sheet shown in right pane was closed, switch to first available
        name = self._sv.get('right_name')
        if name and name not in self.workbook_sheets:
            names = list(self.workbook_sheets.keys())
            if names:
                self._set_right_pane_sheet(names[0])
