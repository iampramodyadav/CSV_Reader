"""
mixins/formula_mixin.py
=======================
All formula-related methods for TableEditor.

REQUIREMENTS in table_editor.py
---------------------------------
1.  Import restored:
        from formula_engine import FormulaEngine

2.  In __init__, self.formula_cells = {} must exist (already there).

3.  In _create_sheet_tab(), uncomment:
        if not hasattr(self, 'formula_engine') or self.formula_engine is None:
            self.formula_engine = FormulaEngine(sheet)

4.  In _on_sheet_change(), guard the formula_engine line:
        if self.formula_engine is not None:
            self.formula_engine.sheet = sheet

5.  UIComponents.create_formula_bar() must set these on app:
        app.formula_var     — tk.StringVar()  for the formula bar entry
        app.auto_calc_var   — tk.BooleanVar() for the auto-calc toggle

6.  end_edit_cell is bound to self.on_cell_edit in _create_sheet_tab —
    this mixin provides on_cell_edit, so that binding works automatically
    through Python's MRO.  Do NOT define on_cell_edit in table_editor.py.

WHAT THIS MIXIN PROVIDES
-------------------------
- on_cell_edit()          bound to tksheet end_edit_cell event
- _on_formula_enter()     called when user presses Enter in formula bar
- _on_formula_focus_out() called when formula bar loses focus
- calculate_current_cell() manual calc button
- recalculate_all()       recalc all formulas in current sheet
- insert_formula_template() insert a template string into selected cell
- show_formula_browser()  searchable function catalogue popup
"""

import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb


class FormulaMixin:
    """Mixin: formula bar handling and evaluation using FormulaEngine."""

    # ── Guard helper ──────────────────────────────────────────────────────────

    def _formula_ready(self) -> bool:
        """
        Return True only if formula_engine is initialised.
        Prevents crashes during startup before the first sheet tab is created.
        """
        return (
            hasattr(self, 'formula_engine')
            and self.formula_engine is not None
        )

    def _get_formula_var(self):
        """
        Return self.formula_var safely.
        UIComponents.create_formula_bar() sets this on app.
        """
        return getattr(self, 'formula_var', None)

    def _get_auto_calc(self) -> bool:
        """Return True if auto-calculate is enabled."""
        var = getattr(self, 'auto_calc_var', None)
        if var is None:
            return True   # default: always auto-calc if var not found
        return var.get()

    # ── Shared highlight helper ───────────────────────────────────────────────

    def _apply_formula_highlight(self, sheet, row: int, col: int, result: str):
        """Apply red highlight for errors, yellow for valid results."""
        if str(result).startswith('#'):
            sheet.highlight_cells(row=row, column=col, bg="#FFCCCC", fg="#CC0000")
        else:
            sheet.highlight_cells(row=row, column=col, bg="lightyellow", fg="black")

    # ── Formula bar events ────────────────────────────────────────────────────

    def _on_formula_enter(self, event=None):
        """
        Called when the user presses Enter in the formula bar,
        or when focus leaves the bar.
        """
        if not self._formula_ready():
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return

        selected = sheet.get_currently_selected()
        if not selected:
            return

        row, col  = selected[0], selected[1]
        fvar      = self._get_formula_var()
        if fvar is None:
            return
        formula = fvar.get()

        if formula.startswith('='):
            self._push_undo() 
            # Register the formula
            if self.current_sheet_name not in self.formula_cells:
                self.formula_cells[self.current_sheet_name] = {}
            cell_key = f"{row},{col}"
            self.formula_cells[self.current_sheet_name][cell_key] = formula

            if self._get_auto_calc():
                self.update_dataframe_from_sheet()
                result = self.formula_engine.evaluate_formula(
                    formula, row, col
                )
                sheet.set_cell_data(row, col, str(result), redraw=True)
                self._apply_formula_highlight(sheet, row, col, str(result))
                self.set_status(f"{formula}  →  {result}")
            else:
                sheet.set_cell_data(row, col, formula, redraw=True)
                sheet.highlight_cells(row=row, column=col,
                                      bg="lightblue", fg="black")
                self.set_status("Formula stored — press Calc to evaluate")
        else:
            # Plain text — clear any formula registration
            self._push_undo() 
            sheet.set_cell_data(row, col, formula, redraw=True)
            if self.current_sheet_name in self.formula_cells:
                cell_key = f"{row},{col}"
                self.formula_cells[self.current_sheet_name].pop(cell_key, None)
                sheet.dehighlight_cells(row=row, column=col)

        self.mark_modified()

    def _on_formula_focus_out(self, event=None):
        """Commit formula bar content when focus leaves the bar."""
        if not self._formula_ready():
            return
        sheet = self.get_current_sheet()
        if not sheet:
            return
        selected = sheet.get_currently_selected()
        if not selected:
            return

        row, col  = selected[0], selected[1]
        fvar      = self._get_formula_var()
        if fvar is None:
            return
        formula   = fvar.get()
        cell_key  = f"{row},{col}"
        sheet_formulas  = self.formula_cells.get(self.current_sheet_name, {})
        current_formula = sheet_formulas.get(cell_key, "")
        current_value   = sheet.get_cell_data(row, col)

        if formula != current_formula and formula != current_value:
            self._on_formula_enter()

    # ── Cell edit event (bound to tksheet end_edit_cell) ─────────────────────

    def on_cell_edit(self, event=None):
        """
        Called by tksheet after the user finishes editing a cell directly
        (typing in the cell, not via formula bar).

        Detects if the new value starts with '=' and evaluates it.
        Also updates the formula bar display to match the selected cell.
        """
        self.mark_modified()
        self._filter_sync_edit() #filter row
        if not self._formula_ready():
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return
        selected = sheet.get_currently_selected()
        if not selected:
            return

        row, col = selected[0], selected[1]
        value    = sheet.get_cell_data(row, col)

        if isinstance(value, str) and value.startswith('='):
            if self.current_sheet_name not in self.formula_cells:
                self.formula_cells[self.current_sheet_name] = {}
            cell_key = f"{row},{col}"
            self.formula_cells[self.current_sheet_name][cell_key] = value

            if self._get_auto_calc():
                self.update_dataframe_from_sheet()
                result = self.formula_engine.evaluate_formula(
                    value, row, col
                )
                sheet.set_cell_data(row, col, str(result), redraw=True)
                self._apply_formula_highlight(sheet, row, col, str(result))
                self.set_status(f"{value}  →  {result}")
            else:
                sheet.highlight_cells(row=row, column=col,
                                      bg="lightblue", fg="black")
                self.set_status("Formula entered — press Calc to evaluate")
        else:
            # Plain value typed — remove formula registration if any
            if self.current_sheet_name in self.formula_cells:
                cell_key = f"{row},{col}"
                self.formula_cells[self.current_sheet_name].pop(cell_key, None)
                sheet.dehighlight_cells(row=row, column=col)

        # Update formula bar to show what's in the cell
        fvar = self._get_formula_var()
        if fvar is not None:
            cell_key = f"{row},{col}"
            stored = self.formula_cells.get(
                self.current_sheet_name, {}
            ).get(cell_key)
            fvar.set(stored if stored else (value or ''))

    # ── Manual calculation ────────────────────────────────────────────────────

    def calculate_current_cell(self):
        """Calculate the selected cell (manual-calc mode)."""
        if not self._formula_ready():
            messagebox.showinfo("Formula", "Formula engine not ready.")
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return
        selected = sheet.get_currently_selected()
        if not selected:
            messagebox.showinfo("Info", "Select a cell first.")
            return

        row, col  = selected[0], selected[1]
        cell_key  = f"{row},{col}"
        sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})

        if cell_key not in sheet_formulas:
            messagebox.showinfo("Info", "Selected cell has no formula.")
            return

        formula = sheet_formulas[cell_key]
        self.update_dataframe_from_sheet()
        result = self.formula_engine.evaluate_formula(
            formula, row, col
        )
        sheet.set_cell_data(row, col, str(result), redraw=True)
        self._apply_formula_highlight(sheet, row, col, str(result))
        self.set_status(f"{formula}  →  {result}")
        self.mark_modified()

    def recalculate_all(self):
        """Recalculate every formula in the current sheet."""
        if not self._formula_ready():
            messagebox.showinfo("Formula", "Formula engine not ready.")
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return

        sheet_formulas = self.formula_cells.get(self.current_sheet_name, {})
        if not sheet_formulas:
            messagebox.showinfo("Info", "No formulas in current sheet.")
            return

        # Always sync df before batch recalculation
        self.update_dataframe_from_sheet()

        count  = 0
        errors = 0

        for cell_key, formula in sheet_formulas.items():
            try:
                r, c   = map(int, cell_key.split(','))
                result = self.formula_engine.evaluate_formula(
                    formula, r, c
                )
                sheet.set_cell_data(r, c, str(result), redraw=False)
                self._apply_formula_highlight(sheet, r, c, str(result))
                if str(result).startswith('#'):
                    errors += 1
                count += 1
            except Exception as exc:
                print(f"recalculate_all: error in {cell_key}: {exc}")
                errors += 1

        sheet.refresh()

        msg = f"Recalculated {count} formula(s)"
        if errors:
            msg += f"  —  {errors} error(s)"
        self.set_status(msg)
        messagebox.showinfo("Recalculate All", f"Done.\n\n  Calculated: {count}\n  Errors: {errors}")

    # ── Template insertion ────────────────────────────────────────────────────

    def insert_formula_template(self, template: str):
        """Insert a formula template string into the selected cell."""
        if not self._formula_ready():
            return

        sheet = self.get_current_sheet()
        if not sheet:
            return
        selected = sheet.get_currently_selected()
        if not selected:
            messagebox.showinfo("Info", "Select a cell first.")
            return

        row, col = selected[0], selected[1]

        if self.current_sheet_name not in self.formula_cells:
            self.formula_cells[self.current_sheet_name] = {}
        cell_key = f"{row},{col}"
        self.formula_cells[self.current_sheet_name][cell_key] = template

        if self._get_auto_calc():
            self.update_dataframe_from_sheet()
            result = self.formula_engine.evaluate_formula(
                template, row, col
            )
            sheet.set_cell_data(row, col, str(result), redraw=True)
            self._apply_formula_highlight(sheet, row, col, str(result))
            self.set_status(f"{template}  →  {result}")
        else:
            sheet.set_cell_data(row, col, template, redraw=True)
            sheet.highlight_cells(row=row, column=col,
                                  bg="lightblue", fg="black")
            self.set_status(f"Formula stored: {template}")

        fvar = self._get_formula_var()
        if fvar is not None:
            fvar.set(template)
        self.mark_modified()

    # ── Formula function browser popup ────────────────────────────────────────

    def show_formula_browser(self):
        """Searchable catalogue of all 645+ supported functions."""
        CATALOGUE = {
            "Math": [
                ("SUM",      "=SUM(A1:A10)",             "Sum a range"),
                ("AVERAGE",  "=AVERAGE(A1:A10)",         "Average of a range"),
                ("MAX",      "=MAX(A1:A10)",              "Maximum value"),
                ("MIN",      "=MIN(A1:A10)",              "Minimum value"),
                ("ROUND",    "=ROUND(A1,2)",              "Round to N decimals"),
                ("ABS",      "=ABS(A1)",                  "Absolute value"),
                ("SQRT",     "=SQRT(A1)",                 "Square root"),
                ("POWER",    "=POWER(A1,2)",              "Raise to power"),
                ("MOD",      "=MOD(A1,B1)",               "Remainder"),
                ("INT",      "=INT(A1)",                  "Floor to integer"),
                ("CEILING",  "=CEILING(A1,1)",            "Round up to multiple"),
                ("FLOOR",    "=FLOOR(A1,1)",              "Round down to multiple"),
            ],
            "Statistics": [
                ("COUNT",     "=COUNT(A1:A100)",          "Count numeric cells"),
                ("COUNTA",    "=COUNTA(A1:A100)",         "Count non-empty cells"),
                ("COUNTIF",   "=COUNTIF(A1:A10,\">0\")",  "Count cells meeting condition"),
                ("SUMIF",     "=SUMIF(A1:A10,\">0\",B1:B10)", "Sum if condition met"),
                ("AVERAGEIF", "=AVERAGEIF(A1:A10,\">0\")", "Average if condition met"),
                ("STDEV",     "=STDEV(A1:A10)",           "Standard deviation"),
                ("MEDIAN",    "=MEDIAN(A1:A10)",          "Median value"),
                ("LARGE",     "=LARGE(A1:A10,2)",         "Nth largest value"),
                ("SMALL",     "=SMALL(A1:A10,2)",         "Nth smallest value"),
            ],
            "Logic": [
                ("IF",       "=IF(A1>0,\"yes\",\"no\")",  "Conditional if/then/else"),
                ("IFERROR",  "=IFERROR(A1/B1,\"ERR\")",  "Handle errors"),
                ("AND",      "=AND(A1>0,B1>0)",           "True if all conditions true"),
                ("OR",       "=OR(A1>0,B1>0)",            "True if any condition true"),
                ("NOT",      "=NOT(A1>0)",                "Logical NOT"),
                ("ISBLANK",  "=ISBLANK(A1)",              "True if cell is empty"),
                ("ISERROR",  "=ISERROR(A1)",              "True if cell has error"),
                ("ISNUMBER", "=ISNUMBER(A1)",             "True if cell is numeric"),
            ],
            "Lookup": [
                ("VLOOKUP",  "=VLOOKUP(A1,B:C,2,0)",     "Vertical lookup"),
                ("HLOOKUP",  "=HLOOKUP(A1,B1:Z2,2,0)",   "Horizontal lookup"),
                ("INDEX",    "=INDEX(A1:C10,2,3)",        "Value at row/col in range"),
                ("MATCH",    "=MATCH(A1,B1:B10,0)",       "Position of value in range"),
            ],
            "Text": [
                ("CONCATENATE","=CONCATENATE(A1,\" \",B1)", "Join text"),
                ("LEFT",     "=LEFT(A1,3)",               "First N characters"),
                ("RIGHT",    "=RIGHT(A1,3)",              "Last N characters"),
                ("MID",      "=MID(A1,2,4)",              "Characters from position"),
                ("LEN",      "=LEN(A1)",                  "Length of text"),
                ("UPPER",    "=UPPER(A1)",                "Convert to uppercase"),
                ("LOWER",    "=LOWER(A1)",                "Convert to lowercase"),
                ("TRIM",     "=TRIM(A1)",                 "Remove extra spaces"),
                ("TEXT",     "=TEXT(A1,\"0.00\")",        "Format number as text"),
                ("VALUE",    "=VALUE(A1)",                "Text to number"),
            ],
            "Date": [
                ("TODAY",    "=TODAY()",                  "Today's date"),
                ("NOW",      "=NOW()",                    "Current date and time"),
                ("YEAR",     "=YEAR(A1)",                 "Year from date"),
                ("MONTH",    "=MONTH(A1)",                "Month from date"),
                ("DAY",      "=DAY(A1)",                  "Day from date"),
            ],
        }

        win = tb.Toplevel(self.root)
        win.title("Formula Browser")
        win.geometry("700x500")
        win.resizable(True, True)

        # Header
        hdr = tb.Frame(win, bootstyle="primary")
        hdr.pack(fill="x")
        tb.Label(
            hdr,
            text="  Formula Browser — Excel-compatible functions",
            font=("Segoe UI", 10, "bold"),
            bootstyle="inverse-primary"
        ).pack(side="left", pady=6)

        # Search
        sf = tb.Frame(win)
        sf.pack(fill="x", padx=10, pady=(8, 4))
        tb.Label(sf, text="Search:").pack(side="left")
        search_var = tk.StringVar()
        se = tb.Entry(sf, textvariable=search_var, width=30)
        se.pack(side="left", padx=6)
        se.focus_set()

        # Main split
        main = tb.Frame(win)
        main.pack(fill="both", expand=True, padx=10, pady=4)

        lf = tb.Frame(main)
        lf.pack(side="left", fill="both", expand=True)
        lb = tk.Listbox(lf, font=("Consolas", 9), selectmode="single",
                        activestyle="dotbox")
        lbsb = tb.Scrollbar(lf, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=lbsb.set)
        lb.pack(side="left", fill="both", expand=True)
        lbsb.pack(side="right", fill="y")

        rf = tb.Frame(main, width=240)
        rf.pack(side="right", fill="y", padx=(10, 0))
        rf.pack_propagate(False)

        tb.Label(rf, text="Template:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tmpl_var = tk.StringVar()
        tb.Entry(rf, textvariable=tmpl_var, font=("Consolas", 9),
                 width=28).pack(fill="x", pady=4)

        tb.Label(rf, text="Description:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        desc_txt = tk.Text(rf, height=4, wrap="word", font=("Segoe UI", 9),
                           state="disabled", relief="flat")
        desc_txt.pack(fill="x", pady=4)

        def do_insert():
            tmpl = tmpl_var.get()
            if tmpl:
                fvar = self._get_formula_var()
                if fvar:
                    fvar.set(tmpl)
                self.set_status(
                    f"Template inserted: {tmpl}  — press Enter to evaluate"
                )
                win.destroy()

        tb.Button(rf, text="Insert into Formula Bar",
                  command=do_insert,
                  bootstyle="success").pack(fill="x", pady=6)
        tb.Button(rf, text="Insert & Calculate",
                  command=lambda: (do_insert(), self._on_formula_enter()),
                  bootstyle="primary").pack(fill="x")

        # Build flat list
        all_items = [
            (grp, name, tmpl, desc)
            for grp, funcs in CATALOGUE.items()
            for name, tmpl, desc in funcs
        ]
        displayed = []

        def refresh(*_):
            q = search_var.get().strip().upper()
            lb.delete(0, "end")
            displayed.clear()
            for grp, name, tmpl, desc in all_items:
                if not q or q in name or q in desc.upper():
                    lb.insert("end", f"  {name:<16}  {grp}")
                    displayed.append((grp, name, tmpl, desc))

        def on_select(_=None):
            sel = lb.curselection()
            if not sel:
                return
            _, name, tmpl, desc = displayed[sel[0]]
            tmpl_var.set(tmpl)
            desc_txt.configure(state="normal")
            desc_txt.delete("1.0", "end")
            desc_txt.insert("1.0", desc)
            desc_txt.configure(state="disabled")

        lb.bind("<<ListboxSelect>>", on_select)
        lb.bind("<Double-Button-1>", lambda e: do_insert())
        search_var.trace_add("write", refresh)

        refresh()
        if displayed:
            lb.selection_set(0)
            on_select()