# Table Editor — Keyboard Shortcuts Reference

## File

| Shortcut | Action |
|---|---|
| Ctrl+N | New file (clears session) |
| Ctrl+O | Open file |
| Ctrl+Shift+O | Add file to current session |
| Ctrl+S | Save current tab (silent — no popup) |
| Ctrl+Shift+S | Save ALL modified tabs |
| Ctrl+W | Close current tab |

## Edit

| Shortcut | Action |
|---|---|
| Ctrl+Z | Undo (df-level + cell edits, unified stack) |
| Ctrl+Y | Redo |
| Ctrl+C | Copy selected cells |
| Ctrl+X | Cut selected cells |
| Ctrl+V | Paste |
| Ctrl+A | Select all cells |
| Delete | Clear selected cells |
| Ctrl+R | Autofill selection (pattern detection) |
| Ctrl+Shift+H | Find & Replace across ALL open sheets |
| Ctrl+H | Find & Replace in current sheet (tksheet built-in) |

## Navigation

| Shortcut | Action |
|---|---|
| Arrow keys | Move one cell |
| Shift+Arrow | Extend selection |
| Home | Jump to first column in current row |
| End | Jump to last non-empty cell in current row |
| Ctrl+Home | Jump to A1 (first cell) |
| Ctrl+End | Jump to last cell of last row with data |
| Ctrl+Down | Jump to last non-empty cell in current column |
| Page Up | Scroll up one page |
| Page Down | Scroll down one page |
| Tab | Move right (confirm edit and advance) |
| Ctrl+F | Find (tksheet built-in search) |
| Ctrl+G | Find next |
| Ctrl+Shift+G | Find previous |

## Sheet / Tab

| Shortcut | Action |
|---|---|
| Ctrl+Next | Next tab |
| Ctrl+Prior | Previous tab |

## View

| Shortcut | Action |
|---|---|
| Ctrl+F | Toggle filter bar |
| Ctrl+Shift+V | Toggle split view (Other View) |
| Ctrl+P | Command palette |

## Mouse / Right-Click

| Action | What it does |
|---|---|
| Right-click cell | Context menu |
| Right-click → Extract Selection to New Tab | Creates new tab from selected rows |
| Right-click → Column Stats | Statistics for the selected column |
| Right-click → Autofill | Fill selection with detected pattern |
| Right-click → Rename Sheet | Rename the current tab |
| Right-click → Close This Tab | Close current tab (prompts if unsaved) |
| Right-click → Trim Spaces | Remove leading/trailing whitespace |
| Ctrl+Click | Add to selection |
| Drag column header | Reorder columns |
| Drag row header | Reorder rows |
| Double-click column border | Auto-resize column |
| Double-click row border | Auto-resize row |

## Filter Bar (active when visible)

| Shortcut | Action |
|---|---|
| Escape | Close filter bar (restores all rows) |
| Enter | Re-apply current filter |
| Type text | Live filter — matches as you type |
| * in query | Wildcard (e.g. MB* matches MB1, MBA, MB_root) |
| Regex checkbox | Full regular expression mode |
| Case checkbox | Case-sensitive matching |

## tksheet Built-in (always available)

| Shortcut | Action |
|---|---|
| Ctrl+Space | Select entire column |
| Shift+Space | Select entire row |
| Ctrl+Shift+Space | Select all |

---

## Notes

- **Undo** covers both cell edits (typed in the grid) and dataframe operations (trim, fill nulls, plugins). All operations share one unified stack per tab. Switching tabs gives each tab its own independent undo history.
- **Filter bar edits** are written back to the original data. Editing a filtered row changes the correct row in the full dataset. New rows added while filtered are appended to the end of the original data.
- **Save (Ctrl+S)** is silent — it saves and shows confirmation only in the status bar, no popup. Save All (Ctrl+Shift+S) saves every modified tab in one operation.
- **End key** finds the last non-empty cell in the row. If the row is entirely empty, it jumps to the last column.
- **Ctrl+End** finds the bottom-right cell that contains data across the whole sheet.
