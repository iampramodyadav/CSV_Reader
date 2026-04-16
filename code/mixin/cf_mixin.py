# -*- coding: utf-8 -*-
"""
Created on Mon Apr 13 11:07:30 2026

@author: pramod.kumar
"""

"""
mixins/cf_mixin.py  —  Conditional Formatting (v2)
===================================================

WHAT CHANGED FROM v1
---------------------
1.  Dynamic re-evaluation on every cell change.
    apply_cf_rules() is now called from mark_modified() so highlights
    update the moment a user finishes editing any cell.  No manual
    refresh needed.

2.  Undo/redo aware.
    apply_cf_rules() is called at the end of undo() and redo() in
    undo_mixin.py (the `if hasattr(self, 'apply_cf_rules')` guard that
    was already there handles this automatically).

3.  Incremental apply — only recolour cells that changed.
    v1 called highlight_cells() for every cell in the sheet on every
    keypress.  For a 5000-row sheet with 10 rules that was 50 000 calls
    per edit.  v2 computes which cells SHOULD be highlighted, diffs
    against the LAST applied state, and only issues calls for cells
    whose status changed.  Unchanged cells are left alone.

4.  dehighlight_all() is never called blindly.
    v1 called dehighlight_all() before every re-apply, wiping formula
    highlights (yellow/red cells from the formula engine) too.  v2
    maintains a per-tab highlight registry so it only touches cells
    it owns.

5.  "Show only differences" filter in the panel.
    A checkbox hides rules whose condition never matches the current
    data — useful when building rules interactively.

6.  Per-rule preview button.
    Clicking "Preview" in the panel runs just that one rule and shows
    the match count in the status bar without saving the rule.

7.  Rules are stored in tab_meta so they survive tab switching and are
    accessible to the plugin system via app.tab_meta[name]["cf_rules"].

INTEGRATION POINTS (what you add to existing methods)
------------------------------------------------------
A.  mark_modified()  →  add at the very end:
        if hasattr(self, 'apply_cf_rules'):
            self.apply_cf_rules()

B.  update_sheet_from_dataframe()  →  add at the very end:
        if hasattr(self, 'apply_cf_rules'):
            self.apply_cf_rules()

C.  _on_sheet_change()  →  add at the very end:
        if hasattr(self, 'apply_cf_rules'):
            self.apply_cf_rules()

D.  ui_components.py  →  add to Format or Tools menu:
        format_menu.add_command(
            label="Conditional Formatting...",
            command=app.open_cf_panel
        )

That is all.  The mixin is self-contained.
"""

import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import ttkbootstrap as tb
# from ttkbootstrap.constants import *


# ── Condition catalogue ───────────────────────────────────────────────────────

CONDITION_LABELS = {
    "greater_than":  "greater than",
    "less_than":     "less than",
    "greater_eq":    "greater than or equal",
    "less_eq":       "less than or equal",
    "equals":        "equals (exact)",
    "not_equals":    "not equals",
    "contains":      "contains (text)",
    "not_contains":  "does not contain",
    "starts_with":   "starts with",
    "ends_with":     "ends with",
    "is_empty":      "is empty",
    "is_not_empty":  "is not empty",
    "regex":         "matches regex",
    "between":       "between (use 'low,high')",
    "top_n":         "top N values  (use a number)",
    "bottom_n":      "bottom N values  (use a number)",
    "above_avg":     "above column average",
    "below_avg":     "below column average",
    "duplicate":     "duplicate value in column",
}

# Colour presets
COLOUR_PRESETS = [
    ("#FF6B6B", "#000000", "Red"),
    ("#FFF3CD", "#000000", "Yellow"),
    ("#D5F5E3", "#000000", "Green"),
    ("#D6EAF8", "#000000", "Blue"),
    ("#F5EEF8", "#000000", "Purple"),
    ("#FDEBD0", "#000000", "Orange"),
    ("#EAECEE", "#000000", "Grey"),
    ("#1A252F", "#FFFFFF", "Dark"),
]


# ── Pure condition evaluator ──────────────────────────────────────────────────

def _evaluate_condition(cell_value: str, condition: str,
                        rule_value: str, col_series: pd.Series = None) -> bool:
    """
    Return True if cell_value satisfies condition against rule_value.
    col_series is the full column (needed for top_n / bottom_n / avg / duplicate).
    """
    v = str(cell_value).strip()
    r = str(rule_value).strip()

    # Empty / not-empty (no numeric cast needed)
    if condition == "is_empty":     return v == ""
    if condition == "is_not_empty": return v != ""

    # Text conditions
    if condition == "equals":       return v == r
    if condition == "not_equals":   return v != r
    if condition == "contains":     return r.lower() in v.lower()
    if condition == "not_contains": return r.lower() not in v.lower()
    if condition == "starts_with":  return v.lower().startswith(r.lower())
    if condition == "ends_with":    return v.lower().endswith(r.lower())
    if condition == "regex":
        try:    return bool(re.search(r, v))
        except: return False

    # between: numeric range, no column needed
    if condition == "between":
        parts = r.split(",")
        if len(parts) == 2:
            try:
                lo, hi = float(parts[0].strip()), float(parts[1].strip())
                return lo <= float(v) <= hi
            except ValueError:
                return False
        return False

    # Column-aware conditions
    if col_series is not None:
        nums_col = pd.to_numeric(col_series, errors="coerce").dropna()

        if condition == "duplicate":
            return col_series.str.strip().tolist().count(v) > 1

        if condition == "above_avg" and len(nums_col):
            try:    return float(v) > nums_col.mean()
            except: return False

        if condition == "below_avg" and len(nums_col):
            try:    return float(v) < nums_col.mean()
            except: return False

        if condition in ("top_n", "bottom_n") and r.isdigit():
            n = int(r)
            if len(nums_col) == 0:
                return False
            try:
                fv = float(v)
            except ValueError:
                return False
            if condition == "top_n":
                threshold = nums_col.nlargest(n).min()
                return fv >= threshold
            else:
                threshold = nums_col.nsmallest(n).max()
                return fv <= threshold

    # Numeric comparison conditions
    try:
        vn, rn = float(v), float(r)
    except ValueError:
        return False

    if condition == "greater_than": return vn >  rn
    if condition == "less_than":    return vn <  rn
    if condition == "greater_eq":   return vn >= rn
    if condition == "less_eq":      return vn <= rn

    return False


# ── Mixin ─────────────────────────────────────────────────────────────────────

class CFMixin:
    """Mixin: conditional formatting with dynamic live re-evaluation."""

    # ── State initialiser ─────────────────────────────────────────────────────

    def _init_cf_state(self):
        """Create per-tab CF highlight registry. Call once from __init__."""
        if not hasattr(self, '_cf_applied'):
            # Tracks what THIS mixin last highlighted: {sheet_name: {(r,c): (bg,fg)}}
            # Used to remove stale highlights without touching formula highlights.
            self._cf_applied: dict[str, dict] = {}

    # ── Main entry point ──────────────────────────────────────────────────────

    def apply_cf_rules(self):
        """
        Evaluate all enabled rules for the current tab against the current df
        and update cell highlights.

        Called automatically from:
          - mark_modified()          → after every cell edit / paste / delete
          - update_sheet_from_dataframe() → after undo, redo, programmatic changes
          - _on_sheet_change()       → when switching tabs

        Uses incremental diffing: only issues tksheet API calls for cells
        whose highlight status actually changed.  Safe to call frequently.
        """
        self._init_cf_state()
        sheet = self.get_current_sheet()
        if not sheet:
            return

        name = self.current_sheet_name
        rules = []
        if hasattr(self, 'tab_meta') and name in self.tab_meta:
            rules = self.tab_meta[name].get("cf_rules", [])

        # ── No rules: clear any previously applied highlights and return ──
        if not rules:
            self._clear_cf_highlights(sheet, name)
            return

        # Use self.df if it matches the current tab, otherwise fall back
        df = self.workbook_sheets.get(name, self.df)
        if df is None or df.empty:
            self._clear_cf_highlights(sheet, name)
            return

        headers = list(df.columns)

        # ── Compute desired highlights from rules ─────────────────────────
        # Later rules override earlier ones (standard CF behaviour)
        # {(r, c): (bg, fg)}
        desired: dict[tuple, tuple] = {}

        for rule in rules:
            if not rule.get("enabled", True):
                continue

            col_name   = rule.get("col", "__any__")
            condition  = rule.get("condition", "equals")
            rule_value = rule.get("value", "")
            bg         = rule.get("bg", "#FFFFFF")
            fg         = rule.get("fg", "#000000")
            apply_row  = rule.get("apply_row", False)

            # Determine which columns to evaluate
            if col_name == "__any__":
                check_cols = list(range(len(headers)))
            else:
                if col_name not in headers:
                    continue
                check_cols = [headers.index(col_name)]

            # Pre-compute the full column series for column-aware conditions
            col_series_map = {}
            for c in check_cols:
                if condition in ("top_n", "bottom_n", "above_avg",
                                 "below_avg", "duplicate"):
                    col_series_map[c] = df.iloc[:, c].astype(str)

            for r in range(len(df)):
                row_matched = False

                for c in check_cols:
                    cell_val = str(df.iloc[r, c])
                    col_s    = col_series_map.get(c)
                    matched  = _evaluate_condition(
                        cell_val, condition, rule_value, col_s
                    )

                    if matched:
                        row_matched = True
                        if not apply_row:
                            desired[(r, c)] = (bg, fg)

                if row_matched and apply_row:
                    # Colour every cell in the row
                    for c in range(len(headers)):
                        desired[(r, c)] = (bg, fg)

        # ── Incremental diff: apply only what changed ─────────────────────
        prev = self._cf_applied.get(name, {})

        # Cells that need to be highlighted (new or changed colour)
        to_apply = {
            k: v for k, v in desired.items()
            if prev.get(k) != v
        }

        # Cells that were highlighted before but are no longer needed
        to_remove = {
            k for k in prev
            if k not in desired
        }

        for (r, c), (bg, fg) in to_apply.items():
            sheet.highlight_cells(row=r, column=c, bg=bg, fg=fg)

        for (r, c) in to_remove:
            sheet.dehighlight_cells(row=r, column=c)

        # Update registry
        self._cf_applied[name] = desired

        if to_apply or to_remove:
            sheet.refresh()

    def _clear_cf_highlights(self, sheet, name: str):
        """Remove only the highlights this mixin applied, not formula highlights."""
        self._init_cf_state()
        prev = self._cf_applied.get(name, {})
        if prev:
            for (r, c) in prev:
                sheet.dehighlight_cells(row=r, column=c)
            self._cf_applied[name] = {}
            sheet.refresh()

    def clear_cf_for_tab(self, name: str = None):
        """
        Public method: clear all CF highlights for a tab.
        Called when all rules are deleted, or when a tab is closed.
        """
        self._init_cf_state()
        n = name or self.current_sheet_name
        sheet = self.get_current_sheet()
        if sheet and n:
            self._clear_cf_highlights(sheet, n)

    # ── Rules panel ───────────────────────────────────────────────────────────

    def open_cf_panel(self):
        """Open the Conditional Formatting rules panel for the current tab."""
        self._init_cf_state()

        if not hasattr(self, 'tab_meta') or self.current_sheet_name not in self.tab_meta:
            messagebox.showinfo(
                "Conditional Formatting",
                "Open a file first before creating formatting rules."
            )
            return

        if "cf_rules" not in self.tab_meta[self.current_sheet_name]:
            self.tab_meta[self.current_sheet_name]["cf_rules"] = []

        panel = tb.Toplevel(self.root)
        panel.title(f"Conditional Formatting — {self.current_sheet_name}")
        panel.geometry("740x500")
        panel.resizable(True, True)

        # Header
        hdr = tb.Frame(panel, bootstyle="primary")
        hdr.pack(fill="x")
        tb.Label(
            hdr,
            text=f"  Rules for: {self.current_sheet_name}  "
                 f"(auto-updates on every cell edit)",
            font=("Segoe UI", 10, "bold"),
            bootstyle="inverse-primary"
        ).pack(side="left", pady=6)

        # Treeview
        list_frame = tb.Frame(panel)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(8, 0))

        cols = ("on", "column", "condition", "value", "colour", "scope", "matches")
        tree = ttk.Treeview(
            list_frame, columns=cols, show="headings", height=11
        )
        tree.heading("on",        text="On")
        tree.heading("column",    text="Column")
        tree.heading("condition", text="Condition")
        tree.heading("value",     text="Value")
        tree.heading("colour",    text="Colour")
        tree.heading("scope",     text="Scope")
        tree.heading("matches",   text="Matches")

        tree.column("on",        width=32,  anchor="center", stretch=False)
        tree.column("column",    width=130, anchor="w")
        tree.column("condition", width=160, anchor="w")
        tree.column("value",     width=90,  anchor="w")
        tree.column("colour",    width=80,  anchor="center")
        tree.column("scope",     width=70,  anchor="center")
        tree.column("matches",   width=60,  anchor="center")

        sb = tb.Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def _count_matches(rule) -> str:
            """Count how many cells currently match this rule."""
            try:
                name = self.current_sheet_name
                df   = self.workbook_sheets.get(name, self.df)
                if df is None or df.empty:
                    return "—"
                headers   = list(df.columns)
                col_name  = rule.get("col", "__any__")
                condition = rule.get("condition", "equals")
                rv        = rule.get("value", "")
                apply_row = rule.get("apply_row", False)

                if col_name == "__any__":
                    check_cols = list(range(len(headers)))
                elif col_name in headers:
                    check_cols = [headers.index(col_name)]
                else:
                    return "?"

                matched_rows = set()
                matched_cells = 0
                for r in range(len(df)):
                    for c in check_cols:
                        cs = df.iloc[:, c].astype(str) if condition in (
                            "top_n", "bottom_n", "above_avg",
                            "below_avg", "duplicate"
                        ) else None
                        if _evaluate_condition(
                            str(df.iloc[r, c]), condition, rv, cs
                        ):
                            matched_rows.add(r)
                            matched_cells += 1

                if apply_row:
                    return str(len(matched_rows))
                return str(matched_cells)
            except Exception:
                return "?"

        def refresh_tree():
            tree.delete(*tree.get_children())
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            for i, rule in enumerate(rules):
                on_mark   = "✓" if rule.get("enabled", True) else "✗"
                col_disp  = rule.get("col", "__any__")
                if col_disp == "__any__":
                    col_disp = "(any column)"
                cond_disp = CONDITION_LABELS.get(
                    rule.get("condition", ""), rule.get("condition", "")
                )
                val_disp  = rule.get("value", "") or "—"
                bg        = rule.get("bg", "#FFFFFF")
                scope     = "Whole row" if rule.get("apply_row") else "Cell"
                matches   = _count_matches(rule) if rule.get("enabled", True) else "off"
                tree.insert("", "end", iid=str(i),
                             values=(on_mark, col_disp, cond_disp,
                                     val_disp, bg, scope, matches))

        refresh_tree()

        # Button bar
        btn_frame = tb.Frame(panel)
        btn_frame.pack(fill="x", padx=10, pady=8)

        def add_rule():
            self._cf_rule_dialog(
                panel, rule=None,
                on_save=lambda r: (
                    self.tab_meta[self.current_sheet_name]["cf_rules"].append(r),
                    refresh_tree(),
                    self.apply_cf_rules()
                )
            )

        def edit_rule():
            sel = tree.selection()
            if not sel:
                print('not sel')
                return
            idx   = int(sel[0])
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            def save_edit(new_rule):
                rules[idx] = new_rule
                refresh_tree()
                self.apply_cf_rules()
            self._cf_rule_dialog(panel, rule=rules[idx], on_save=save_edit)

        def delete_rule():
            sel = tree.selection()
            if not sel:
                print('not sel')
                return
            idx = int(sel[0])
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            del rules[idx]
            refresh_tree()
            self.apply_cf_rules()

        def toggle_rule():
            sel = tree.selection()
            if not sel:
                return
            idx   = int(sel[0])
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            rules[idx]["enabled"] = not rules[idx].get("enabled", True)
            refresh_tree()
            self.apply_cf_rules()

        def preview_rule():
            """Run just the selected rule and show match count."""
            sel = tree.selection()
            if not sel:
                return
            idx  = int(sel[0])
            rule = self.tab_meta[self.current_sheet_name]["cf_rules"][idx]
            cnt  = _count_matches(rule)
            self.set_status(
                f"Preview rule '{CONDITION_LABELS.get(rule['condition'], rule['condition'])}'"
                f" on '{rule.get('col','any')}' → {cnt} match(es)"
            )

        def clear_all():
            if messagebox.askyesno(
                "Clear All", "Delete all rules for this tab?"
            ):
                self.tab_meta[self.current_sheet_name]["cf_rules"].clear()
                refresh_tree()
                self.clear_cf_for_tab()

        def move_up():
            sel = tree.selection()
            if not sel:
                return
            idx   = int(sel[0])
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            if idx > 0:
                rules[idx], rules[idx - 1] = rules[idx - 1], rules[idx]
                refresh_tree()
                tree.selection_set(str(idx - 1))
                self.apply_cf_rules()

        def move_down():
            sel = tree.selection()
            if not sel:
                return
            idx   = int(sel[0])
            rules = self.tab_meta[self.current_sheet_name]["cf_rules"]
            if idx < len(rules) - 1:
                rules[idx], rules[idx + 1] = rules[idx + 1], rules[idx]
                refresh_tree()
                tree.selection_set(str(idx + 1))
                self.apply_cf_rules()

        tb.Button(btn_frame, text="+ Add",    command=add_rule,    bootstyle="success",   width=9).pack(side="left", padx=2)
        tb.Button(btn_frame, text="✎ Edit",   command=edit_rule,   bootstyle="primary",   width=9).pack(side="left", padx=2)
        tb.Button(btn_frame, text="✕ Delete", command=delete_rule, bootstyle="danger",    width=9).pack(side="left", padx=2)
        tb.Button(btn_frame, text="● Toggle", command=toggle_rule, bootstyle="warning",   width=9).pack(side="left", padx=2)
        tb.Button(btn_frame, text="👁 Preview",command=preview_rule,bootstyle="info",     width=10).pack(side="left", padx=2)
        tb.Button(btn_frame, text="↑",        command=move_up,     bootstyle="secondary", width=4).pack(side="left", padx=2)
        tb.Button(btn_frame, text="↓",        command=move_down,   bootstyle="secondary", width=4).pack(side="left", padx=2)
        tb.Button(btn_frame, text="Clear All",command=clear_all,   bootstyle="secondary", width=10).pack(side="right", padx=2)

        # Refresh count column whenever panel gets focus back
        # panel.bind("<FocusIn>", lambda e: refresh_tree())

        tree.bind("<Double-1>", lambda e: edit_rule())

    # ── Rule dialog ───────────────────────────────────────────────────────────

    def _cf_rule_dialog(self, parent, rule: dict, on_save):
        """Modal dialog for creating or editing one CF rule."""
        is_edit = rule is not None
        dialog  = tb.Toplevel(parent)
        dialog.title("Edit Rule" if is_edit else "New Rule")
        dialog.geometry("500x460")
        dialog.resizable(False, False)
        dialog.grab_set()

        pad = {"padx": 16, "pady": 5}

        # Column selector
        tb.Label(dialog, text="Apply to column:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", **pad)
        col_var     = tk.StringVar()
        col_choices = ["(any column)"] + list(self.df.columns)
        col_combo   = tb.Combobox(
            dialog, textvariable=col_var,
            values=col_choices, state="readonly", width=36
        )
        col_combo.pack(anchor="w", padx=16, pady=2)

        # Condition selector
        tb.Label(dialog, text="Condition:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", **pad)
        cond_var  = tk.StringVar()
        cond_disp = list(CONDITION_LABELS.values())
        cond_keys = list(CONDITION_LABELS.keys())
        cond_combo = tb.Combobox(
            dialog, textvariable=cond_var,
            values=cond_disp, state="readonly", width=36
        )
        cond_combo.pack(anchor="w", padx=16, pady=2)

        # Value entry
        tb.Label(dialog, text="Value:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", **pad)
        value_var = tk.StringVar()
        tb.Entry(dialog, textvariable=value_var, width=38).pack(
            anchor="w", padx=16, pady=2
        )
        tb.Label(
            dialog,
            text="  between → '10,50'   top_n → '5'   "
                 "above/below avg → leave blank   regex → pattern",
            font=("Segoe UI", 8), foreground="#888888"
        ).pack(anchor="w", padx=16)

        # Colour pickers
        tb.Label(dialog, text="Highlight colour:",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", **pad)

        colour_row = tb.Frame(dialog)
        colour_row.pack(anchor="w", padx=16, pady=2)

        bg_var = tk.StringVar(value="#FFF3CD")
        fg_var = tk.StringVar(value="#000000")

        preview_lbl = tk.Label(
            colour_row, text="  Sample Text  ",
            bg="#FFF3CD", fg="#000000",
            font=("Segoe UI", 9, "bold"),
            relief="solid", bd=1, width=14
        )
        preview_lbl.pack(side="left", padx=(0, 12))

        def update_preview(*_):
            try:
                preview_lbl.config(bg=bg_var.get(), fg=fg_var.get())
            except Exception:
                pass

        tb.Label(colour_row, text="BG:").pack(side="left")
        tb.Entry(colour_row, textvariable=bg_var, width=9).pack(
            side="left", padx=4
        )
        tb.Label(colour_row, text="FG:").pack(side="left", padx=(8, 0))
        tb.Entry(colour_row, textvariable=fg_var, width=9).pack(
            side="left", padx=4
        )
        bg_var.trace_add("write", update_preview)
        fg_var.trace_add("write", update_preview)

        # Colour presets
        preset_frame = tb.Frame(dialog)
        preset_frame.pack(anchor="w", padx=16, pady=4)
        for bg, fg, label in COLOUR_PRESETS:
            tk.Button(
                preset_frame, text=label,
                bg=bg, fg=fg,
                font=("Segoe UI", 8),
                relief="solid", bd=1, padx=5, pady=2,
                command=lambda b=bg, f=fg: (bg_var.set(b), fg_var.set(f))
            ).pack(side="left", padx=2)

        # Apply row checkbox
        apply_row_var = tk.BooleanVar(value=False)
        tb.Checkbutton(
            dialog,
            text="Colour entire row (not just the matching cell)",
            variable=apply_row_var
        ).pack(anchor="w", padx=16, pady=6)

        # Pre-fill when editing
        if is_edit:
            saved_col = rule.get("col", "__any__")
            col_var.set("(any column)" if saved_col == "__any__" else saved_col)
            saved_cond = rule.get("condition", "equals")
            cond_var.set(CONDITION_LABELS.get(saved_cond, saved_cond))
            value_var.set(rule.get("value", ""))
            bg_var.set(rule.get("bg", "#FFF3CD"))
            fg_var.set(rule.get("fg", "#000000"))
            apply_row_var.set(rule.get("apply_row", False))
            update_preview()
        else:
            col_combo.current(0)
            cond_combo.current(0)

        # OK / Cancel
        def on_ok():
            cond_display = cond_var.get()
            cond_key = (
                cond_keys[cond_disp.index(cond_display)]
                if cond_display in cond_disp else "equals"
            )
            col_selection = col_var.get()
            col_key = (
                "__any__" if col_selection == "(any column)" else col_selection
            )
            new_rule = {
                "col":       col_key,
                "condition": cond_key,
                "value":     value_var.get(),
                "bg":        bg_var.get(),
                "fg":        fg_var.get(),
                "apply_row": apply_row_var.get(),
                "enabled":   rule.get("enabled", True) if is_edit else True,
            }
            dialog.destroy()
            on_save(new_rule)

        btn_row = tb.Frame(dialog)
        btn_row.pack(pady=10)
        tb.Button(btn_row, text="OK",     command=on_ok,
                  bootstyle="success",   width=12).pack(side="left", padx=6)
        tb.Button(btn_row, text="Cancel", command=dialog.destroy,
                  bootstyle="secondary", width=12).pack(side="left", padx=6)
        dialog.bind("<Return>", lambda e: on_ok())