"""
mixins/undo_mixin.py  —  Unified single-stack undo/redo for TableEditor
========================================================================

THE PROBLEM WITH THE OLD APPROACH
----------------------------------
The previous implementation had two separate undo systems running in parallel:
  1. tksheet's internal cell-level undo (Ctrl+Z via tksheet)
  2. Our df-snapshot undo (Ctrl+Z via our code)

This caused the bugs you saw:
  - Undo sometimes hit tksheet's stack (cell edit), sometimes ours (df op)
  - Redo from our stack after a tksheet undo left the two stacks out of sync
  - Typing in a cell and then undoing a df op could leave the sheet showing
    data that didn't match the df

THE FIX: ONE STACK, ONE SOURCE OF TRUTH
-----------------------------------------
Every undoable action — whether the user typed in a cell or we called
trim_whitespace() — goes through OUR df-snapshot stack. tksheet's own
undo/redo never runs.

How it works:
  1. In _create_sheet_tab(), we call disable_bindings("undo", "redo").
     This unbinds tksheet's Ctrl+Z/Ctrl+Y and sets undo_enabled=False.
     tksheet will no longer push anything to its own undo stack.

  2. We bind <<SheetModified>> on each sheet widget. tksheet fires this
     event AFTER every user edit (typing, paste, delete, insert row, etc.).
     Our handler (_on_sheet_modified) snapshots the df BEFORE the edit
     (we keep a "last known good df" per tab in _pre_edit_snapshots) and
     pushes that snapshot onto our stack.

  3. Ctrl+Z and Ctrl+Y in main_app.py call app.undo() and app.redo().
     Both operate entirely on our deque — no tksheet involvement at all.

  4. df-mutating methods (trim_whitespace, fill_nulls, plugins, etc.) call
     self._push_undo() before mutating. This works the same as before.

PER-TAB ISOLATION
-----------------
Every tab has its own undo deque and redo deque stored in:
  self.undo_stacks[sheet_name]   deque(maxlen=MAX_UNDO)
  self.redo_stacks[sheet_name]   deque(maxlen=MAX_UNDO)
  self._pre_edit_snapshots[sheet_name]   pd.DataFrame (the df as of last refresh)

Switching tabs gives you that tab's independent undo history.

WHAT TO CHANGE IN EXISTING CODE
---------------------------------
1.  _create_sheet_tab():
      After sheet.enable_bindings(...), add:
        sheet.disable_bindings("undo", "redo")
        sheet.bind("<<SheetModified>>", lambda e, s=sheet_name: self._on_sheet_modified(s))

2.  main_app.py:
      Replace:
        root.bind("<Control-z>", lambda e: app.get_current_sheet().undo() ...)
        root.bind("<Control-y>", lambda e: app.get_current_sheet().redo() ...)
      With:
        root.bind("<Control-z>", lambda e: app.undo())
        root.bind("<Control-y>", lambda e: app.redo())

3.  Add self._push_undo() as the first line of every df-mutating method:
        clean_whitespace, clean_nan, find_replace_all, any plugin runner.
    (Same requirement as before — just the name stays _push_undo().)

4.  close_current_tab(): add self._clear_undo_for_tab(name) before removing tab.

5.  load_file() / new_file(): call self._clear_undo_for_tab(name) when clearing tabs.

That is all. The mixin is self-contained; nothing else needs changing.
"""

from collections import deque
import pandas as pd


MAX_UNDO = 30   # history depth per tab


class UndoMixin:
    """
    Mixin: unified single-stack undo/redo for both cell edits and df operations.
    Attach to TableEditor alongside other mixins.
    """

    # ── Internal state initialiser ────────────────────────────────────────────
    # Called once — either from __init__ or lazily on first use via
    # _ensure_stacks(). Both are safe.

    def _init_undo_state(self):
        """Create the three dicts that hold per-tab undo state."""
        if not hasattr(self, 'undo_stacks'):
            self.undo_stacks: dict[str, deque] = {}
        if not hasattr(self, 'redo_stacks'):
            self.redo_stacks: dict[str, deque] = {}
        # Snapshot of the df as it was the last time we synced from the sheet.
        # Used by _on_sheet_modified to know what to push onto the undo stack.
        if not hasattr(self, '_pre_edit_snapshots'):
            self._pre_edit_snapshots: dict[str, pd.DataFrame] = {}

    def _ensure_stacks(self, name: str = None):
        """Guarantee stacks exist for the given tab name (default: current)."""
        self._init_undo_state()
        n = name or self.current_sheet_name
        if n and n not in self.undo_stacks:
            self.undo_stacks[n]  = deque(maxlen=MAX_UNDO)
            self.redo_stacks[n]  = deque(maxlen=MAX_UNDO)
            # Seed pre-edit snapshot with current df so first edit has a base
            self._pre_edit_snapshots[n] = (
                self.workbook_sheets.get(n, pd.DataFrame()).copy()
            )

    # ── Hook called from _create_sheet_tab ───────────────────────────────────

    def _bind_sheet_modified(self, sheet_widget, sheet_name: str):
        """
        Wire up the <<SheetModified>> listener for a newly created sheet tab.
    
        Call from _create_sheet_tab() after enable_bindings():
    
            sheet.disable_bindings("undo", "redo")
            self._bind_sheet_modified(sheet, sheet_name)
    
        NOTE: deliberately does NOT seed the pre-edit snapshot here.
        load_file() calls _refresh_pre_edit_snapshot(name) explicitly
        after _populate_sheet() so the snapshot has real data in it.
        """
        self._init_undo_state()
        # Create empty stacks for this tab (no snapshot yet — caller seeds it)
        if sheet_name not in self.undo_stacks:
            self.undo_stacks[sheet_name]  = deque(maxlen=MAX_UNDO)
            self.redo_stacks[sheet_name]  = deque(maxlen=MAX_UNDO)
        # Pre-edit snapshot is intentionally NOT seeded here
    
        def _handler(event, sname=sheet_name):
            self._on_sheet_modified(sname)
    
        sheet_widget.bind("<<SheetModified>>", _handler)

    # ── <<SheetModified>> handler ─────────────────────────────────────────────

    def _on_sheet_modified(self, sheet_name: str):
        """
        Called by tksheet immediately after any user edit that would normally
        go onto tksheet's own undo stack (typed edit, paste, delete,
        insert/delete row or column, sort, etc.).

        We push the PRE-EDIT state (the snapshot we kept from last time the
        sheet was synced) onto our undo stack, then update the snapshot to
        the current post-edit state.

        This is the key insight: because we capture the state BEFORE the edit
        (not after), pressing Ctrl+Z restores exactly the state the user saw
        before they made the change.
        """
        self._ensure_stacks(sheet_name)

        # The pre-edit snapshot is what we push — it's the state before
        # this modification happened
        pre = self._pre_edit_snapshots.get(sheet_name)
        if pre is None:
            # No baseline yet — snapshot the current state so the NEXT edit
            # has something to push. Don't push anything to undo yet.
            self._refresh_pre_edit_snapshot(sheet_name)
            return                        # ← early exit, nothing to undo yet
    
        self.undo_stacks[sheet_name].append(pre)
        self.redo_stacks[sheet_name].clear()
        self._refresh_pre_edit_snapshot(sheet_name)

    def _refresh_pre_edit_snapshot(self, sheet_name: str = None):
        """
        Update the pre-edit snapshot for a tab to its current state.
        Call this:
          - after update_sheet_from_dataframe()  (df → sheet direction)
          - after load_file() finishes populating a sheet
          - after undo/redo restores state (so next edit has correct base)
        """
        self._init_undo_state()
        n = sheet_name or self.current_sheet_name
        if not n:
            return

        # Try to read from workbook_sheets first (already synced df)
        df = self.workbook_sheets.get(n)
        if df is not None:
            self._pre_edit_snapshots[n] = df.copy()

    # ── Public: called from _push_undo() and df-mutating methods ─────────────

    def _push_undo(self, sheet_name: str = None):
        """
        Snapshot the current df BEFORE a programmatic mutation.

        Call this as the very first line of any method that changes self.df
        directly (trim_whitespace, fill_nulls, plugin runner, find_replace_all,
        sort, etc.).

        It syncs the sheet → df first to make sure the snapshot is current,
        then pushes and clears the redo stack.
        """
        n = sheet_name or self.current_sheet_name
        self._ensure_stacks(n)

        # Sync sheet → df so snapshot reflects any typed-but-not-synced edits
        self.update_dataframe_from_sheet()

        df = self.workbook_sheets.get(n)
        if df is not None:
            self.undo_stacks[n].append(df.copy())
            self.redo_stacks[n].clear()

        # Also update the pre-edit snapshot so the next SheetModified
        # doesn't double-push the same state
        self._refresh_pre_edit_snapshot(n)

    # ── Public: Ctrl+Z ───────────────────────────────────────────────────────

    def undo(self):
        """
        Undo one step on the current tab.

        Pops the top df snapshot, pushes current state to redo stack,
        restores the sheet. If the stack is empty, shows a status message.
        """
        n = self.current_sheet_name
        self._ensure_stacks(n)
        stack = self.undo_stacks.get(n)

        if not stack:
            self.set_status("Nothing to undo")
            return

        # Sync current state to df before pushing to redo
        self.update_dataframe_from_sheet()

        # Push current state onto redo stack
        current_df = self.workbook_sheets.get(n)
        if current_df is not None:
            self.redo_stacks[n].append(current_df.copy())

        # Restore previous state
        prev_df = stack.pop()
        self.df = prev_df
        self.workbook_sheets[n] = prev_df

        # Push to sheet display
        self.update_sheet_from_dataframe()

        # IMPORTANT: update pre-edit snapshot to this restored state
        # so the next user edit has the correct base to snapshot from
        self._refresh_pre_edit_snapshot(n)

        # Re-apply conditional formatting if available
        if hasattr(self, 'apply_cf_rules'):
            self.apply_cf_rules()

        remaining = len(stack)
        self.set_status(
            f"Undo  —  {remaining} step(s) remaining"
            if remaining else "Undo  —  no more history"
        )
        self.modified = True
        self.mark_modified()
        # self.update_title()

    # ── Public: Ctrl+Y ───────────────────────────────────────────────────────

    def redo(self):
        """
        Redo one step on the current tab.

        Pops the top redo snapshot, pushes current state back to undo stack,
        restores the sheet.
        """
        n = self.current_sheet_name
        self._ensure_stacks(n)
        stack = self.redo_stacks.get(n)

        if not stack:
            self.set_status("Nothing to redo")
            return

        # Sync current state
        self.update_dataframe_from_sheet()

        # Push current state back onto undo stack
        current_df = self.workbook_sheets.get(n)
        if current_df is not None:
            self.undo_stacks[n].append(current_df.copy())

        # Restore redo state
        next_df = stack.pop()
        self.df = next_df
        self.workbook_sheets[n] = next_df

        self.update_sheet_from_dataframe()
        self._refresh_pre_edit_snapshot(n)

        if hasattr(self, 'apply_cf_rules'):
            self.apply_cf_rules()

        self.set_status(f"Redo  —  {len(stack)} redo step(s) remaining" if stack else "Redo")
        self.modified = True
        self.mark_modified()
        # self.update_title()

    # ── Housekeeping ──────────────────────────────────────────────────────────

    def _clear_undo_for_tab(self, sheet_name: str):
        """
        Free memory when a tab is closed or replaced.
        Call from close_current_tab() and load_file() when clearing tabs.
        """
        self._init_undo_state()
        self.undo_stacks.pop(sheet_name, None)
        self.redo_stacks.pop(sheet_name, None)
        self._pre_edit_snapshots.pop(sheet_name, None)

    def get_undo_count(self) -> int:
        """How many undo steps are available for the current tab."""
        self._ensure_stacks()
        return len(self.undo_stacks.get(self.current_sheet_name, []))

    def get_redo_count(self) -> int:
        """How many redo steps are available for the current tab."""
        self._ensure_stacks()
        return len(self.redo_stacks.get(self.current_sheet_name, []))
    
    def _remove_empty_startup_sheet(self):
        """
        If the only open tab is the blank 'Sheet1' created at startup
        (no file loaded, df is empty, never modified), remove it silently.
    
        This prevents the empty startup state from appearing as an undo step
        when the user opens their first real file.
        """
        tabs = self.sheet_notebook.tabs()
        if len(tabs) != 1:
            return   # multiple tabs already open — don't touch anything
    
        only_name = self.sheet_notebook.tab(tabs[0], "text")
    
        # Check it really is the pristine startup sheet
        meta     = getattr(self, 'tab_meta', {}).get(only_name, {})
        df       = self.workbook_sheets.get(only_name, pd.DataFrame())
        modified = meta.get("modified", False) if meta else self.modified
        has_file = bool(meta.get("file_path")) if meta else bool(self.current_file)
    
        if only_name == "Sheet1" and df.empty and not modified and not has_file:
            # Remove from notebook and all registries
            self.sheet_notebook.forget(tabs[0])
            self.workbook_sheets.pop(only_name, None)
            self._clear_undo_for_tab(only_name)
            if hasattr(self, 'tab_meta'):
                self.tab_meta.pop(only_name, None)
            self.formula_cells.pop(only_name, None)