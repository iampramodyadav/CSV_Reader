"""
mixins/extras_mixin.py
======================
Six self-contained productivity features:

  1. Sub-sheet from selection  — right-click → "Extract to New Tab"
  2. Auto-open recent files    — opens last N files on launch
  3. Duplicate file guard      — blocks opening the same file path twice
  4. End / Ctrl+End navigation — keyboard shortcuts tksheet doesn't provide
  5. Save All                  — saves every dirty tab in one keystroke
  6. Silent save               — removes the "File saved successfully!" popup

INTEGRATION
-----------
In table_editor.py:
    from mixins.extras_mixin import ExtrasMixin
    class TableEditor(..., ExtrasMixin):

In __init__, after _setup_plugin_menu():
    self._setup_extras()

In _create_sheet_tab(), add to popup menu and bindings — see Step 3 below.

In main_app.py:
    root.bind("<Control-s>",       lambda e: app.save_file())
    root.bind("<Control-Shift-S>", lambda e: app.save_all())

On application startup (end of main() before root.mainloop()):
    app.open_recent_on_startup(n=5)
"""

import os
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import ttkbootstrap as tb


class ExtrasMixin:
    """Mixin: sub-sheet, auto-open recent, duplicate guard, navigation, save-all, silent save."""

    # ── Init ─────────────────────────────────────────────────────────────────

    def _setup_extras(self):
        """Call once from __init__ after _setup_ui."""
        # Nothing to initialise at construction time;
        # all state is created lazily or per-call.
        pass

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 1 — Sub-sheet: extract selected data to a new tab
    # ─────────────────────────────────────────────────────────────────────────

    def extract_selection_to_new_tab(self):
        """
        Take the currently selected cells (any shape) and open them as a
        new tab preserving the original column headers.

        Behaviour:
          • Selects the columns whose cells are included in the selection.
          • Extracts only those rows where at least one selected cell exists.
          • The new tab is named  "<source>_extract"  and is unsaved (no file).
          • Full column headers are kept — not just the selected column subset.
            (If you want only selected columns, set FULL_COLUMNS = False below.)

        Add to popup menu in _create_sheet_tab():
            sheet.popup_menu_add_command(
                "Extract Selection to New Tab",
                self.extract_selection_to_new_tab
            )
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return

        selected = list(sheet.get_selected_cells())
        if not selected:
            self.set_status("Nothing selected — select cells first.")
            return

        # Sync df
        self.update_dataframe_from_sheet()
        df = self.df

        if df.empty:
            self.set_status("Sheet is empty.")
            return

        # Determine which rows and columns are in the selection
        sel_rows = sorted(set(r for r, _ in selected))
        sel_cols = sorted(set(c for _, c in selected))

        # Guard against out-of-bounds
        n_rows, n_cols = len(df), len(df.columns)
        sel_rows = [r for r in sel_rows if r < n_rows]
        sel_cols = [c for c in sel_cols if c < n_cols]

        if not sel_rows or not sel_cols:
            self.set_status("Selection is out of bounds.")
            return

        # Extract — keep ALL columns but only selected rows
        # Change to df.iloc[sel_rows, sel_cols] if you want only selected cols
        extracted = df.iloc[sel_rows].reset_index(drop=True)

        # Name the new tab
        base_name  = self.current_sheet_name
        new_name   = f"{base_name}_extract"
        counter    = 1
        while new_name in self.workbook_sheets:
            new_name = f"{base_name}_extract_{counter}"
            counter += 1

        # Create tab (unsaved — no file_path)
        widget = self._create_sheet_tab(
            new_name, extracted,
            file_path=None, sep=","
        )
        self._populate_sheet(widget, extracted)
        self.sheet_notebook.select(len(self.sheet_notebook.tabs()) - 1)
        self._sync_globals_from_current_tab()

        self.set_status(
            f"Extracted {len(sel_rows)} row(s) × {len(df.columns)} col(s) "
            f"→ '{new_name}'  (unsaved)"
        )

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 2 — Auto-open recent files on startup
    # ─────────────────────────────────────────────────────────────────────────

    def open_recent_on_startup(self, n: int = 5):
        """
        Open the last N recently-used files automatically at launch.
        Silently skips files that no longer exist.

        Call at the END of main() in main_app.py, just before root.mainloop():

            app.open_recent_on_startup(n=5)

        The value of n can also be read from preferences:
            n = app.get_pref('startup_recent', 0)
        so the user can disable it by setting n=0 in Preferences.
        """
        if n <= 0:
            return

        paths = self._load_recent()          # already filters non-existent files
        paths = paths[:n]

        if not paths:
            return

        for fp in paths:
            # Use the duplicate guard — skip already-open files
            if self._file_already_open(fp):
                continue
            try:
                self.load_file(fp)
            except Exception as exc:
                # Non-fatal — log and continue
                self.set_status(f"Could not auto-open: {os.path.basename(fp)}")
                print(f"open_recent_on_startup: {fp}: {exc}")

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 3 — Duplicate file guard
    # ─────────────────────────────────────────────────────────────────────────

    def _file_already_open(self, file_path: str) -> bool:
        """
        Return True if file_path is already open in any tab.
        Compares resolved absolute paths so relative vs absolute doesn't matter.
        Different files with the same name at different locations are allowed.
        """
        try:
            incoming = os.path.normcase(os.path.realpath(file_path))
        except Exception:
            return False

        for meta in self.tab_meta.values():
            fp = meta.get("file_path")
            if fp:
                try:
                    existing = os.path.normcase(os.path.realpath(fp))
                    if existing == incoming:
                        return True
                except Exception:
                    pass
        return False

    def load_file_guarded(self, file_path: str, force_sep: str = None):
        """
        load_file() with duplicate check.
        Use this in open_file() and add_file() instead of load_file().

        Returns True if the file was loaded, False if it was already open.
        """
        if self._file_already_open(file_path):
            # Find the tab name to highlight it
            tab_name = None
            try:
                incoming = os.path.normcase(os.path.realpath(file_path))
                for name, meta in self.tab_meta.items():
                    fp = meta.get("file_path")
                    if fp and os.path.normcase(os.path.realpath(fp)) == incoming:
                        tab_name = name
                        break
            except Exception:
                pass

            msg = (
                f"'{os.path.basename(file_path)}' is already open"
                + (f" in tab '{tab_name}'" if tab_name else "")
                + "."
            )
            messagebox.showwarning("Already Open", msg)

            # Switch to the existing tab
            if tab_name:
                for tab_id in self.sheet_notebook.tabs():
                    if self.sheet_notebook.tab(tab_id, "text") == tab_name:
                        self.sheet_notebook.select(tab_id)
                        break

            return False

        self.load_file(file_path, force_sep=force_sep)
        return True

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 4 — End / Ctrl+End navigation
    # ─────────────────────────────────────────────────────────────────────────

    def _nav_end_of_row(self, event=None):
        """
        End key — jump to the last non-empty cell in the current row.
        If all cells are empty, jumps to the last column.
        Mirrors Notepad++ / Excel End key behaviour.
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return "break"

        sel = sheet.get_currently_selected()
        if not sel:
            return "break"

        row = sel[0]
        n_cols = sheet.total_columns()
        if n_cols == 0:
            return "break"

        # Find last non-empty column in this row
        last_col = n_cols - 1
        for c in range(n_cols - 1, -1, -1):
            val = sheet.get_cell_data(row, c)
            if val and str(val).strip():
                last_col = c
                break

        sheet.see(row, last_col)
        sheet.select_cell(row, last_col, redraw=True)
        return "break"

    def _nav_end_of_col(self, event=None):
        """
        Ctrl+Down — jump to the last non-empty cell in the current column.
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return "break"

        sel = sheet.get_currently_selected()
        if not sel:
            return "break"

        col = sel[1]
        n_rows = sheet.total_rows()
        if n_rows == 0:
            return "break"

        last_row = n_rows - 1
        for r in range(n_rows - 1, -1, -1):
            val = sheet.get_cell_data(r, col)
            if val and str(val).strip():
                last_row = r
                break

        sheet.see(last_row, col)
        sheet.select_cell(last_row, col, redraw=True)
        return "break"

    def _nav_ctrl_end(self, event=None):
        """
        Ctrl+End — jump to the last cell of the last row with data.
        Excel / Notepad++ behaviour.
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return "break"

        n_rows = sheet.total_rows()
        n_cols = sheet.total_columns()
        if n_rows == 0 or n_cols == 0:
            return "break"

        # Find last row that has any non-empty cell
        last_row = 0
        last_col = 0
        for r in range(n_rows - 1, -1, -1):
            for c in range(n_cols - 1, -1, -1):
                val = sheet.get_cell_data(r, c)
                if val and str(val).strip():
                    last_row, last_col = r, c
                    break
            else:
                continue
            break

        sheet.see(last_row, last_col)
        sheet.select_cell(last_row, last_col, redraw=True)
        return "break"

    def _nav_ctrl_home(self, event=None):
        """Ctrl+Home — jump to cell A1 (already handled by tksheet but provided for completeness)."""
        sheet = self.get_current_sheet()
        if not sheet:
            return "break"
        if sheet.total_rows() > 0 and sheet.total_columns() > 0:
            sheet.see(0, 0)
            sheet.select_cell(0, 0, redraw=True)
        return "break"

    def bind_navigation_shortcuts(self, sheet):
        """
        Bind End / Ctrl+End / Ctrl+Down navigation keys to a sheet widget.
        Call this inside _create_sheet_tab() after sheet.pack():

            self.bind_navigation_shortcuts(sheet)

        Home and Ctrl+Home are already handled natively by tksheet when
        "arrowkeys" binding is enabled.
        """
        sheet.bind("<End>",          self._nav_end_of_row)
        sheet.bind("<Control-End>",  self._nav_ctrl_end)
        sheet.bind("<Control-Down>", self._nav_end_of_col)
        sheet.bind("<Control-Home>", self._nav_ctrl_home)   # override tksheet's if needed

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 5 — Save All
    # ─────────────────────────────────────────────────────────────────────────

    def save_all(self):
        """
        Save every tab that has unsaved changes.
        Tabs with no file_path (unsaved new files) are skipped with a notice.
        Does NOT show a popup per file — only a final summary in the status bar.

        Bind to Ctrl+Shift+S in main_app.py:
            root.bind("<Control-Shift-S>", lambda e: app.save_all())

        Add to File menu:
            file_menu.add_command(label="Save All   Ctrl+Shift+S", command=app.save_all)
        """
        dirty = [
            name for name, meta in self.tab_meta.items()
            if meta.get("modified")
        ]

        if not dirty:
            self.set_status("All tabs already saved.")
            return

        saved  = []
        failed = []
        skipped = []
        saved_paths = set()   # avoid writing the same Excel file twice

        original_tab = self.current_sheet_name

        for tab_name in dirty:
            meta = self.tab_meta.get(tab_name, {})
            fp   = meta.get("file_path")

            if not fp:
                skipped.append(tab_name)
                continue

            # For Excel files, all tabs sharing the same path are saved together
            if fp in saved_paths:
                continue

            # Temporarily switch context to this tab so save_file_silent works
            self.current_sheet_name = tab_name
            self._sync_globals_from_current_tab()

            try:
                self._save_file_silent()   # saves without popup
                saved_paths.add(fp)
                saved.append(tab_name)
            except Exception as exc:
                failed.append(f"{tab_name} ({exc})")

        # Restore original tab
        self.current_sheet_name = original_tab
        self._sync_globals_from_current_tab()

        # Status summary
        parts = []
        if saved:
            parts.append(f"Saved {len(saved)} tab(s)")
        if skipped:
            parts.append(f"{len(skipped)} unsaved (no file)")
        if failed:
            parts.append(f"FAILED: {', '.join(failed)}")
        self.set_status("  |  ".join(parts) if parts else "Nothing to save.")

        if skipped:
            self.set_status(
                self.right_status.cget("text") +
                f"  —  {len(skipped)} tab(s) have no file path yet (use Save As)"
            )

    # ─────────────────────────────────────────────────────────────────────────
    # FEATURE 6 — Silent save (no popup)
    # ─────────────────────────────────────────────────────────────────────────

    def _save_file_silent(self):
        """
        Exactly like save_file() but WITHOUT the messagebox.showinfo popup.
        Used internally by save_all() and replaces save_file() for Ctrl+S.

        To make ALL saves silent, replace save_file() calls with this,
        OR simply remove the messagebox.showinfo line from save_file().
        """
        meta = self.tab_meta.get(self.current_sheet_name)
        if not meta or not meta.get("file_path"):
            self.save_file_as()
            return

        self.update_dataframe_from_sheet()
        file_path = meta["file_path"]

        if meta.get("is_excel"):
            excel_sheets = {
                sname: self.workbook_sheets[sname]
                for sname, m in self.tab_meta.items()
                if m.get("file_path") == file_path and sname in self.workbook_sheets
            }
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for tab_label, df in excel_sheets.items():
                    real_sheet = (
                        tab_label.split(" » ", 1)[-1]
                        if " » " in tab_label else tab_label
                    )
                    df_save = df.copy()
                    df_save.replace(r'^\s*', '', regex=True, inplace=True)
                    df_save.to_excel(writer, sheet_name=real_sheet, index=False)
            for sname, m in self.tab_meta.items():
                if m.get("file_path") == file_path:
                    m["modified"] = False
        else:
            import re
            df_to_save = self.df.copy()
            sep = meta.get("sep", ",")
            df_to_save.replace(r'^\s*$', 'NaN', regex=True, inplace=True)
            from custom_helpers import write_label_format
            if sep == "#label":
                write_label_format(file_path, self.df, sep="::")
            elif sep == r'\s+' or (sep and re.match(r'^\s+$', sep)):
                df_to_save.to_string(
                    buf=file_path, index=False, col_space=15, justify='left'
                )
            elif file_path.endswith(".tsv"):
                df_to_save.to_csv(file_path, sep="\t", index=False)
            else:
                self.df.to_csv(file_path, sep=sep, index=False, na_rep="")
            meta["modified"] = False

        self.modified = any(m["modified"] for m in self.tab_meta.values())
        self.update_title()
        self.set_status(f"Saved: {os.path.basename(file_path)}")
        # ← No messagebox.showinfo here — that's the whole point
