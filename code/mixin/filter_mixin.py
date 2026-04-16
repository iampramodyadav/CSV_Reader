"""
mixins/filter_mixin.py  —  Row Filter Bar with Regex  (v2 — editable filter view)
==================================================================================

HOW EDIT PASSTHROUGH WORKS
---------------------------
When a filter is active the sheet widget shows only matching rows.
The key data structure is the "index map":

    _filter_index_maps[sheet_name] = [orig_row_0, orig_row_1, ...]

This list maps  filtered_row_position → original_row_index_in_backup_df.

Example:
  backup_df has 200 rows (indices 0-199).
  Filter matches rows 5, 12, 47.
  index_map = [5, 12, 47]
  Sheet widget shows 3 rows (positions 0, 1, 2).

When the user edits position 1 in the sheet:
  orig_idx = index_map[1]  →  12
  backup_df.iloc[12] = new_row_data

When the user adds a new row (it appears at the end of the filtered view):
  It's appended to backup_df as a brand new row.
  index_map is extended with the new index.

When the filter is cleared, backup_df (with all edits applied) is restored.
The edited rows and added rows are all there.

GUARD ON update_dataframe_from_sheet
--------------------------------------
The existing mark_modified() calls update_dataframe_from_sheet() which reads
back whatever the sheet widget shows and writes it to workbook_sheets.
While a filter is active this would overwrite the full backup with only the
filtered rows — corrupting data.

We block this by checking _filter_is_active() at the start of
update_dataframe_from_sheet() in table_editor.py:

    def update_dataframe_from_sheet(self):
        if hasattr(self, '_filter_is_active') and self._filter_is_active():
            return                    # ← ADD THIS GUARD
        sheet = self.get_current_sheet()
        ...

This is the ONE required change in table_editor.py.

INTEGRATION
-----------
1.  from mixins.filter_mixin import FilterMixin
    class TableEditor(..., FilterMixin):

2.  In _setup_ui(), AFTER create_toolbar, BEFORE content_frame:
        self._build_filter_bar(self.root)

3.  In _on_sheet_change(), at the very end:
        self._on_filter_tab_change()

4.  In update_sheet_from_dataframe(), at the very end:
        self._on_df_changed_outside_filter()

5.  In update_dataframe_from_sheet(), ADD GUARD at the very start:
        if hasattr(self, '_filter_is_active') and self._filter_is_active():
            return

6.  Bind end_edit_cell in _create_sheet_tab — this is already bound to
    self.on_cell_edit.  The filter mixin wraps that via _filter_sync_edit.
    In _create_sheet_tab after existing extra_bindings, add:
        sheet.extra_bindings([("end_edit_cell", self._filter_sync_edit)])

    OR: call self._filter_sync_edit(event) at the start of on_cell_edit()
    in formula_mixin.py / table_editor.py.

7.  main_app.py:
        root.bind("<Control-f>", lambda e: app.toggle_filter_bar())
"""

import re
import tkinter as tk
from tkinter import ttk
import pandas as pd
import ttkbootstrap as tb


class FilterMixin:
    """Mixin: editable per-tab row filter bar with regex support."""

    # ── State ─────────────────────────────────────────────────────────────────

    def _init_filter_state(self):
        if not hasattr(self, '_filter'):
            self._filter = {
                'bar_frame':  None,
                'visible':    False,
                'query_var':  None,
                'col_var':    None,
                'col_combo':  None,
                'case_var':   None,
                'regex_var':  None,
                'count_lbl':  None,
                'entry':      None,
            }
        if not hasattr(self, '_filter_backups'):
            # {sheet_name: pd.DataFrame}  — full unfiltered df
            self._filter_backups: dict = {}
        if not hasattr(self, '_filter_index_maps'):
            # {sheet_name: list[int]}  — filtered_row → original_row_index
            self._filter_index_maps: dict = {}
        if not hasattr(self, '_filter_states'):
            # {sheet_name: {query, col, case, regex}}
            self._filter_states: dict = {}

    # ── Public guard (used by update_dataframe_from_sheet) ────────────────────

    def _filter_is_active(self) -> bool:
        """
        Return True if a filter is currently applied on the current tab.
        When True, update_dataframe_from_sheet() must not read back the sheet
        widget (it would read only filtered rows and corrupt the backup).
        """
        self._init_filter_state()
        name = getattr(self, 'current_sheet_name', None)
        return bool(name and name in self._filter_backups)

    # ── Build bar ─────────────────────────────────────────────────────────────

    def _build_filter_bar(self, parent):
        """Build the filter bar widget. Call once from _setup_ui()."""
        self._init_filter_state()

        bar = tk.Frame(parent, bg="#EBF3FB", height=36)
        self._filter['bar_frame'] = bar

        left = tk.Frame(bar, bg="#EBF3FB")
        left.pack(side="left", fill="x", expand=True, padx=6, pady=4)

        tk.Label(left, text="Filter:", bg="#EBF3FB",
                 fg="#0C4A6E", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 4))

        col_var = tk.StringVar(value="All columns")
        col_combo = ttk.Combobox(left, textvariable=col_var,
                                 values=["All columns"], state="readonly",
                                 width=16, font=("Segoe UI", 9))
        col_combo.pack(side="left", padx=(0, 4))
        self._filter['col_var']   = col_var
        self._filter['col_combo'] = col_combo

        query_var = tk.StringVar()
        entry = tb.Entry(left, textvariable=query_var,
                         width=30, font=("Segoe UI", 9))
        entry.pack(side="left", padx=(0, 6))
        self._filter['query_var'] = query_var
        self._filter['entry']     = entry

        tk.Button(left, text="✕",
                  bg="#EBF3FB", fg="#CC0000",
                  relief="flat", font=("Segoe UI", 10, "bold"),
                  cursor="hand2", padx=4,
                  command=self.clear_filter).pack(side="left", padx=(0, 8))

        mid = tk.Frame(bar, bg="#EBF3FB")
        mid.pack(side="left", padx=4)

        case_var  = tk.BooleanVar(value=False)
        regex_var = tk.BooleanVar(value=False)
        self._filter['case_var']  = case_var
        self._filter['regex_var'] = regex_var

        tk.Checkbutton(mid, text="Case",  variable=case_var,
                       bg="#EBF3FB", font=("Segoe UI", 8),
                       command=self._apply_filter).pack(side="left", padx=2)
        tk.Checkbutton(mid, text="Regex", variable=regex_var,
                       bg="#EBF3FB", font=("Segoe UI", 8),
                       command=self._apply_filter).pack(side="left", padx=2)

        right = tk.Frame(bar, bg="#EBF3FB")
        right.pack(side="right", padx=6)

        count_lbl = tk.Label(right, text="",
                             bg="#EBF3FB", fg="#555555",
                             font=("Segoe UI", 9, "italic"))
        count_lbl.pack(side="left", padx=(0, 10))
        self._filter['count_lbl'] = count_lbl

        # Info label explaining editable mode
        tk.Label(right, text="[edits apply to original]",
                 bg="#EBF3FB", fg="#1A7A6E",
                 font=("Segoe UI", 7, "italic")).pack(side="left", padx=(0, 10))

        tk.Button(right, text="Close Filter",
                  bg="#EBF3FB", fg="#0C4A6E",
                  relief="flat", font=("Segoe UI", 8),
                  cursor="hand2", padx=6,
                  command=self.hide_filter_bar).pack(side="left")

        query_var.trace_add("write", lambda *_: self._apply_filter())
        col_var.trace_add("write",   lambda *_: self._apply_filter())
        entry.bind("<Escape>", lambda e: self.hide_filter_bar())
        entry.bind("<Return>", lambda e: self._apply_filter())

    # ── Show / hide ───────────────────────────────────────────────────────────

    def toggle_filter_bar(self):
        self._init_filter_state()
        if self._filter['visible']:
            self.hide_filter_bar()
        else:
            self.show_filter_bar()

    def show_filter_bar(self):
        self._init_filter_state()
        bar = self._filter.get('bar_frame')
        if bar is None:
            return

        bar.pack(fill="x", before=self._get_content_frame())
        self._filter['visible'] = True
        self._update_col_combo()

        entry = self._filter.get('entry')
        if entry:
            entry.focus_set()

        self.set_status("Filter: type to filter rows — edits write back to original data")

    def hide_filter_bar(self):
        self._init_filter_state()
        self.clear_filter()
        bar = self._filter.get('bar_frame')
        if bar:
            bar.pack_forget()
        self._filter['visible'] = False
        self.set_status("Filter closed")
        
    def _get_content_frame(self):
        """
        Return the widget that the filter bar should appear before.
        This MUST be a widget that was packed into the same parent 
        used in _build_filter_bar.
        """
        try:
            # If your UI structure has a main container frame:
            if hasattr(self, 'content_frame') and self.content_frame.winfo_manager() == 'pack':
                return self.content_frame
            
            # Fallback: Try to find the notebook's master
            if self.sheet_notebook.master.winfo_manager() == 'pack':
                return self.sheet_notebook.master
                
            return None
        except Exception:
            return None
    # def _get_content_frame(self):
    #     try:
    #         return self.sheet_notebook.master
    #     except Exception:
    #         return None

    # ── Apply filter ──────────────────────────────────────────────────────────

    def _apply_filter(self):
        """
        Filter the sheet display.
        Stores index_map so edits can be written back to the correct original rows.
        Never corrupts backup_df or workbook_sheets.
        """
        self._init_filter_state()

        query     = self._filter['query_var'].get().strip()
        col       = self._filter['col_var'].get()
        case      = self._filter['case_var'].get()
        use_regex = self._filter['regex_var'].get()

        name  = self.current_sheet_name
        sheet = self.get_current_sheet()
        if not sheet:
            return

        # ── Snapshot backup on first call ─────────────────────────────────
        # We call update_dataframe_from_sheet() directly here (bypassing the
        # guard) because we WANT the current sheet data before the filter.
        # This is the only place we intentionally do this.
        if name not in self._filter_backups:
            s = self.get_current_sheet()
            if s:
                self.df = pd.DataFrame(
                    s.get_sheet_data(), columns=s.headers()
                )
                self.workbook_sheets[name] = self.df
            self._filter_backups[name] = self.df.copy()

        backup_df = self._filter_backups[name]
        total     = len(backup_df)

        # ── No query — show all rows ──────────────────────────────────────
        if not query:
            # index_map = identity mapping
            self._filter_index_maps[name] = list(range(total))
            sheet.headers([str(c) for c in backup_df.columns])
            sheet.set_sheet_data(backup_df.astype(str).values.tolist())
            sheet.refresh()
            self._update_count_label(total, total)
            return

        # ── Build mask ────────────────────────────────────────────────────
        if col == "All columns":
            search_cols = list(backup_df.columns)
        else:
            search_cols = [col] if col in backup_df.columns else list(backup_df.columns)

        try:
            if use_regex:
                flags = 0 if case else re.IGNORECASE
                mask = backup_df[search_cols].apply(
                    lambda c: c.astype(str).str.contains(
                        query, flags=flags, regex=True, na=False)
                ).any(axis=1)
            else:
                escaped = re.escape(query).replace(r'\*', '.*')
                flags   = 0 if case else re.IGNORECASE
                mask = backup_df[search_cols].apply(
                    lambda c: c.astype(str).str.contains(
                        escaped, flags=flags, regex=True, na=False)
                ).any(axis=1)
        except re.error:
            entry = self._filter.get('entry')
            if entry:
                try:
                    entry.configure(bootstyle="danger")
                except Exception:
                    pass
            self._update_count_label(None, total, error=True)
            return

        entry = self._filter.get('entry')
        if entry:
            try:
                entry.configure(bootstyle="default")
            except Exception:
                pass

        # ── Build filtered view and store index map ───────────────────────
        # mask.index contains the original row indices in backup_df
        matched_indices = list(backup_df.index[mask])          # e.g. [5, 12, 47]
        filtered_df     = backup_df.loc[matched_indices]        # keep original index

        # Store the map: filtered_display_row → original_backup_index
        self._filter_index_maps[name] = matched_indices

        # Push to sheet as a clean 0-based display (reset_index only for display)
        sheet.headers([str(c) for c in filtered_df.columns])
        sheet.set_sheet_data(filtered_df.reset_index(drop=True).astype(str).values.tolist())
        sheet.refresh()

        self._update_count_label(len(matched_indices), total)
        self._filter_states[name] = {
            "query": query, "col": col, "case": case, "regex": use_regex
        }

    # ── Edit sync: write filtered-view edits back to backup_df ───────────────

    def _filter_sync_edit(self, event=None):
        """
        Called whenever the user finishes editing a cell while a filter is active.
        Reads the changed cell from the sheet widget and writes it back to
        the correct row in backup_df, then also to workbook_sheets.

        Hook this into end_edit_cell:
            In _create_sheet_tab(), add to extra_bindings:
                ("end_edit_cell", self._filter_sync_edit)

            OR call self._filter_sync_edit() at the top of on_cell_edit()
            before anything else.
        """
        self._init_filter_state()
        name = self.current_sheet_name

        if name not in self._filter_backups:
            return   # filter not active — nothing to do

        sheet = self.get_current_sheet()
        if not sheet:
            return

        selected = sheet.get_currently_selected()
        if not selected:
            return

        filtered_row, col_idx = selected[0], selected[1]
        new_value = sheet.get_cell_data(filtered_row, col_idx)

        index_map = self._filter_index_maps.get(name, [])
        if filtered_row >= len(index_map):
            # Row is beyond the current index map — it's a newly added row
            # (handled by _filter_sync_new_row)
            return

        orig_row_idx = index_map[filtered_row]
        backup_df    = self._filter_backups[name]

        if orig_row_idx not in backup_df.index:
            return

        col_name = backup_df.columns[col_idx] if col_idx < len(backup_df.columns) else None
        if col_name is None:
            return

        # Write the edit into backup_df at the original row
        backup_df.at[orig_row_idx, col_name] = new_value

        # Also update workbook_sheets so the full df stays in sync
        self.workbook_sheets[name] = backup_df

    def _filter_sync_new_row(self):
        """
        Called when the user inserts a new row while a filter is active.
        Appends the new row to backup_df and extends the index_map.

        Hook this into mark_modified or rc_insert_row if needed.
        For simplicity, it's called from _filter_sync_edit when the
        filtered_row is beyond the existing index_map length.
        """
        self._init_filter_state()
        name = self.current_sheet_name

        if name not in self._filter_backups:
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return

        backup_df = self._filter_backups[name]
        index_map = self._filter_index_maps.get(name, [])

        # Read how many rows the sheet widget has now
        n_sheet_rows = sheet.total_rows()
        n_mapped     = len(index_map)

        if n_sheet_rows <= n_mapped:
            return   # no new rows

        # For each new row in the filtered sheet (beyond existing map):
        for filtered_row in range(n_mapped, n_sheet_rows):
            # Read the row data from the sheet widget
            row_data = [
                sheet.get_cell_data(filtered_row, c)
                for c in range(sheet.total_columns())
            ]
            # Build a new Series with the backup_df columns
            new_series = pd.Series(row_data, index=backup_df.columns)
            # Append to backup_df — new index = len(backup_df)
            new_orig_idx = len(backup_df)
            backup_df.loc[new_orig_idx] = new_series
            index_map.append(new_orig_idx)

        self._filter_index_maps[name] = index_map
        self._filter_backups[name]    = backup_df
        self.workbook_sheets[name]    = backup_df

    # ── Clear filter ──────────────────────────────────────────────────────────

    def clear_filter(self):
        """Clear filter and restore the full df (with any edits applied)."""
        self._init_filter_state()

        name  = self.current_sheet_name
        sheet = self.get_current_sheet()

        if name in self._filter_backups:
            # Before restoring, sync any pending sheet edits that haven't
            # been captured yet (in case the user edited without triggering
            # end_edit_cell explicitly)
            self._flush_pending_edits_to_backup(name)

            restored_df = self._filter_backups.pop(name)
            self._filter_index_maps.pop(name, None)

            self.df = restored_df
            self.workbook_sheets[name] = self.df

            if sheet:
                sheet.headers([str(c) for c in self.df.columns])
                sheet.set_sheet_data(self.df.astype(str).values.tolist())
                sheet.refresh()

        qv = self._filter.get('query_var')
        if qv:
            qv.set("")

        self._filter_states.pop(name, None)
        self._update_count_label(None, None)

    def _flush_pending_edits_to_backup(self, name: str):
        """
        Read the current filtered view from the sheet widget and write all
        visible rows back to backup_df using the index map.
        This catches any edits that were made without end_edit_cell firing
        (e.g. pasting a block of cells).
        """
        sheet = self.get_current_sheet()
        if not sheet:
            return

        backup_df = self._filter_backups.get(name)
        index_map = self._filter_index_maps.get(name, [])
        if backup_df is None:
            return

        data    = sheet.get_sheet_data()
        headers = sheet.headers()

        for filtered_row, row_data in enumerate(data):
            if filtered_row < len(index_map):
                # Existing matched row — update in place
                orig_idx = index_map[filtered_row]
                if orig_idx in backup_df.index:
                    for col_idx, val in enumerate(row_data):
                        if col_idx < len(headers):
                            col_name = headers[col_idx]
                            if col_name in backup_df.columns:
                                backup_df.at[orig_idx, col_name] = val
            else:
                # New row beyond the index map — append to backup
                new_series  = pd.Series(
                    row_data[:len(backup_df.columns)],
                    index=backup_df.columns
                )
                new_orig_idx = len(backup_df)
                backup_df.loc[new_orig_idx] = new_series
                index_map.append(new_orig_idx)

        self._filter_backups[name]    = backup_df
        self._filter_index_maps[name] = index_map

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _update_count_label(self, matched, total, error=False):
        lbl = self._filter.get('count_lbl')
        if lbl is None:
            return
        if error:
            lbl.config(text="⚠ invalid regex", fg="#CC0000")
        elif matched is None or total is None:
            lbl.config(text="", fg="#555555")
        else:
            colour = "#1E6B3C" if matched == total else "#856404"
            lbl.config(text=f"{matched} / {total} rows", fg=colour)

    def _update_col_combo(self):
        self._init_filter_state()
        combo = self._filter.get('col_combo')
        col_v = self._filter.get('col_var')
        if combo is None or col_v is None:
            return
        cols = ["All columns"] + list(self.df.columns)
        combo['values'] = cols
        if col_v.get() not in cols:
            col_v.set("All columns")

    def _on_filter_tab_change(self):
        """Called from _on_sheet_change() — restore per-tab filter state."""
        self._init_filter_state()
        if not self._filter.get('visible'):
            return

        self._update_col_combo()
        name  = self.current_sheet_name
        state = self._filter_states.get(name)

        qv     = self._filter.get('query_var')
        cv     = self._filter.get('col_var')
        casev  = self._filter.get('case_var')
        regexv = self._filter.get('regex_var')

        if state and qv:
            qv.set(state.get("query", ""))
            cv.set(state.get("col", "All columns"))
            casev.set(state.get("case", False))
            regexv.set(state.get("regex", False))
            self._apply_filter()
        else:
            if qv:
                qv.set("")
            self._update_count_label(None, None)

    def _on_df_changed_outside_filter(self):
        """
        Called from update_sheet_from_dataframe().
        When a programmatic df change happens (undo, plugin, trim) while a
        filter is active, update the backup to the new full df and re-apply.
        """
        self._init_filter_state()
        name = self.current_sheet_name

        if name not in self._filter_backups:
            return

        qv    = self._filter.get('query_var')
        query = qv.get().strip() if qv else ""

        if not query:
            self._filter_backups.pop(name, None)
            self._filter_index_maps.pop(name, None)
            return

        # self.df here is the NEW full df from the external operation
        self._filter_backups[name] = self.df.copy()
        self._apply_filter()