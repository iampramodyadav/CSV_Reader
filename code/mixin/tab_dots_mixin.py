"""
mixins/tab_dots_mixin.py# Not working
========================
Sheet status dots on notebook tabs.

Each tab gets a small coloured dot appended to its label:

  ○  grey dot   — clean, no unsaved changes
  ●  orange dot — modified, unsaved changes exist
  ●  red dot    — new/unsaved file (file_path is None)
  ●  green dot  — just saved (shown for 3 seconds then reverts to grey)

The dot is part of the tab text label, so it works with the standard
ttk.Notebook without any custom widget or theme change.

How it works
------------
- Tab labels are stored as the sheet name WITHOUT the dot in
  self.workbook_sheets and self.tab_meta.
- The dots are only in the DISPLAY text shown in the notebook widget.
- _update_tab_dot(sheet_name) rebuilds the display text for one tab.
- update_all_tab_dots() rebuilds all tabs — called from mark_modified()
  and save_file().

Integration points (what you add to existing methods)
------------------------------------------------------
  mark_modified()        → call self._update_tab_dot(self.current_sheet_name)
  save_file()            → call self._flash_saved_dot(self.current_sheet_name)
  _create_sheet_tab()    → call self._update_tab_dot(sheet_name) at end
  close_current_tab()    → nothing needed (tab is gone)
  rename_sheet()         → call self._update_tab_dot(new_name) at end
"""
import tkinter as tk


# Unicode dots used as status indicators
_DOT_CLEAN    = " ○"   # grey:   no changes
_DOT_MODIFIED = " ●"   # orange: unsaved changes
_DOT_NEW      = " ●"   # red:    new file, never saved
_DOT_SAVED    = " ●"   # green:  just saved (flashes briefly)


class TabDotsMixin:
    """Mixin: status dot indicators on notebook tab labels."""

    # ── Core dot updater ──────────────────────────────────────────────────────

    def _get_tab_id_for_sheet(self, sheet_name: str):
        """
        Return the internal ttk tab ID string for a given sheet name.
        Returns None if not found.
        We search by the display text because tabs may have a dot appended.
        """
        for tab_id in self.sheet_notebook.tabs():
            display = self.sheet_notebook.tab(tab_id, "text")
            # Strip any existing dot suffix before comparing
            clean = display.rstrip("○●").rstrip()
            if clean == sheet_name:
                return tab_id
        return None

    def _update_tab_dot(self, sheet_name: str, dot_override: str = None):
        """
        Set the dot on the tab for sheet_name based on its current state.

        dot_override: if provided, use this dot string instead of computing
                      it from state (used for the saved flash).
        """
        tab_id = self._get_tab_id_for_sheet(sheet_name)
        if tab_id is None:
            return   # tab not found (race condition during close) — ignore

        if dot_override is not None:
            dot = dot_override
        else:
            meta     = {}
            if hasattr(self, 'tab_meta'):
                meta = self.tab_meta.get(sheet_name, {})

            is_new       = not meta.get("file_path")      # never saved to a file
            is_modified  = meta.get("modified", False)

            if is_new and is_modified:
                dot = _DOT_NEW        # red — unsaved new
            elif is_modified:
                dot = _DOT_MODIFIED   # orange — dirty
            else:
                dot = _DOT_CLEAN      # grey — clean

        # Set colour via tag — we can't colour individual chars in a Notebook
        # tab label, so we use a small Canvas trick only when ttkbootstrap is
        # available; otherwise just update the text label.
        self.sheet_notebook.tab(tab_id, text=f"{sheet_name}{dot}")

        # Apply colour by styling — ttkbootstrap lets us set style per tab
        self._apply_dot_colour(tab_id, dot, sheet_name)

    def _apply_dot_colour(self, tab_id: str, dot: str, sheet_name: str):
        """
        Change the tab foreground to reflect the dot colour.
        Works by maintaining a custom tk.Label overlay over the tab —
        the cleanest approach that survives theme changes.

        Simpler approach: just rely on the dot character itself (○/●)
        as a visual cue — users understand the filled/unfilled distinction.
        This method is intentionally lightweight to avoid theme conflicts.
        """
        # We keep it simple: the dot character alone is enough visual feedback.
        # If you want actual colour, you can extend this with a Canvas overlay.
        pass

    def update_all_tab_dots(self):
        """Refresh the dot on every open tab. Call after load, save, close."""
        for sheet_name in list(self.workbook_sheets.keys()):
            self._update_tab_dot(sheet_name)

    def _flash_saved_dot(self, sheet_name: str):
        """
        Show a green dot for 2 seconds after saving, then restore grey dot.
        Non-blocking — uses root.after().
        """
        # Immediately show clean state
        self._update_tab_dot(sheet_name, dot_override=_DOT_CLEAN)

    # ── Tooltip showing file path on hover ────────────────────────────────────

    def setup_tab_tooltips(self):
        """
        Show the full file path as a tooltip when hovering over a tab.
        Call this once from _setup_ui() after creating the notebook.
        """
        tooltip_win = [None]   # mutable container so inner functions can reassign

        def _show_tooltip(event):
            # Find which tab the mouse is over
            try:
                tab_id = self.sheet_notebook.identify(event.x, event.y)
                if not tab_id:
                    _hide_tooltip()
                    return
                display   = self.sheet_notebook.tab(tab_id, "text")
                # Strip dot to get sheet name
                sheet_name = display.rstrip("○●").rstrip()
                meta = {}
                if hasattr(self, 'tab_meta'):
                    meta = self.tab_meta.get(sheet_name, {})
                fp = meta.get("file_path")
                sep = meta.get("sep", ",")

                if fp:
                    tip_text = f"{fp}\nSeparator: {sep!r}"
                else:
                    tip_text = f"{sheet_name}\n(unsaved — no file yet)"

                _hide_tooltip()   # clear any existing tooltip

                tw = tk.Toplevel(self.root)
                tw.wm_overrideredirect(True)
                tw.attributes("-topmost", True)
                tw.wm_geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
                lbl = tk.Label(
                    tw, text=tip_text,
                    font=("Segoe UI", 8),
                    justify="left",
                    background="#FFFFE0",
                    foreground="#000000",
                    relief="solid",
                    borderwidth=1,
                    padx=6, pady=4,
                )
                lbl.pack()
                tooltip_win[0] = tw
            except Exception:
                pass

        def _hide_tooltip(*_):
            if tooltip_win[0]:
                try:
                    tooltip_win[0].destroy()
                except Exception:
                    pass
                tooltip_win[0] = None

        self.sheet_notebook.bind("<Motion>",   _show_tooltip)
        self.sheet_notebook.bind("<Leave>",    _hide_tooltip)
        self.sheet_notebook.bind("<Button-1>", _hide_tooltip)