"""
Excel-like Formula Engine
Supports arithmetic, functions, and Python expressions
"""
import re
import math
from typing import Any, List


class FormulaEngine:
    """
    Excel-like formula engine supporting:
    - Basic arithmetic: =A1+B1, =A1*B1
    - Functions: =SUM(A1:A10), =AVERAGE(A1:A10), =COUNT(A1:A10)
    - IF statements: =IF(A1>10, "High", "Low")
    - Python expressions: =PYTHON(A1*2.5)
    """
    
    def __init__(self, sheet_widget):
        self.sheet = sheet_widget
        self.cell_cache = {}
        
    def evaluate_formula(self, formula: str, current_row: int, current_col: int) -> Any:
        """Evaluate a formula and return the result."""
        if not formula or not str(formula).startswith('='):
            return formula
            
        formula = str(formula)[1:].strip()
        
        try:
            cache_key = f"{current_row},{current_col}"
            if cache_key in self.cell_cache:
                return "#CIRCULAR!"
            self.cell_cache[cache_key] = True
            
            formula_upper = formula.upper()
            
            # Route to appropriate handler
            if formula_upper.startswith('PYTHON(') and formula.endswith(')'):
                result = self._eval_python(formula[7:-1], current_row, current_col)
            elif formula_upper.startswith('SUM(') and formula.endswith(')'):
                result = self._eval_sum(formula[4:-1], current_row, current_col)
            elif formula_upper.startswith('AVERAGE(') and formula.endswith(')'):
                result = self._eval_average(formula[8:-1], current_row, current_col)
            elif formula_upper.startswith('COUNT(') and formula.endswith(')'):
                result = self._eval_count(formula[6:-1], current_row, current_col)
            elif formula_upper.startswith('MAX(') and formula.endswith(')'):
                result = self._eval_max(formula[4:-1], current_row, current_col)
            elif formula_upper.startswith('MIN(') and formula.endswith(')'):
                result = self._eval_min(formula[4:-1], current_row, current_col)
            elif formula_upper.startswith('IF(') and formula.endswith(')'):
                result = self._eval_if(formula[3:-1], current_row, current_col)
            elif formula_upper.startswith('CONCAT(') and formula.endswith(')'):
                result = self._eval_concat(formula[7:-1], current_row, current_col)
            elif formula_upper.startswith('CONCATENATE(') and formula.endswith(')'):
                result = self._eval_concat(formula[12:-1], current_row, current_col)
            elif formula_upper.startswith('SQRT(') and formula.endswith(')'):
                result = self._eval_sqrt(formula[5:-1], current_row, current_col)
            elif formula_upper.startswith('POWER(') and formula.endswith(')'):
                result = self._eval_power(formula[6:-1], current_row, current_col)
            elif formula_upper.startswith('ABS(') and formula.endswith(')'):
                result = self._eval_abs(formula[4:-1], current_row, current_col)
            elif formula_upper.startswith('ROUND(') and formula.endswith(')'):
                result = self._eval_round(formula[6:-1], current_row, current_col)
            elif formula_upper.startswith('LEN(') and formula.endswith(')'):
                result = self._eval_len(formula[4:-1], current_row, current_col)
            elif formula_upper.startswith('UPPER(') and formula.endswith(')'):
                result = self._eval_upper(formula[6:-1], current_row, current_col)
            elif formula_upper.startswith('LOWER(') and formula.endswith(')'):
                result = self._eval_lower(formula[6:-1], current_row, current_col)
            elif formula_upper.startswith('COUNTIF(') and formula.endswith(')'):
                result = self._eval_countif(formula[8:-1], current_row, current_col)
            elif formula_upper.startswith('SUMIF(') and formula.endswith(')'):
                result = self._eval_sumif(formula[6:-1], current_row, current_col)
            else:
                result = self._eval_expression(formula, current_row, current_col)
                
            del self.cell_cache[cache_key]
            return result
            
        except Exception as e:
            if cache_key in self.cell_cache:
                del self.cell_cache[cache_key]
            return f"#ERROR: {str(e)}"
    
    def _parse_cell_ref(self, ref: str) -> tuple:
        """Parse cell reference like 'A1' to (row, col)."""
        match = re.match(r'([A-Z]+)(\d+)', ref.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {ref}")
        col_letters, row_num = match.groups()
        col = 0
        for char in col_letters:
            col = col * 26 + (ord(char) - ord('A') + 1)
        return int(row_num) - 1, col - 1
    
    def _parse_range(self, range_str: str) -> List[tuple]:
        """Parse range like 'A1:A10' to list of (row, col) tuples."""
        parts = range_str.split(':')
        if len(parts) != 2:
            raise ValueError(f"Invalid range: {range_str}")
        
        start_row, start_col = self._parse_cell_ref(parts[0])
        end_row, end_col = self._parse_cell_ref(parts[1])
        
        cells = []
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                cells.append((r, c))
        return cells
    
    def _get_cell_value(self, row: int, col: int) -> float:
        """Get numeric value from cell."""
        try:
            val = self.sheet.get_cell_data(row, col)
            if isinstance(val, str) and val.startswith('='):
                val = self.evaluate_formula(val, row, col)
            return float(val) if val else 0
        except:
            return 0
    
    # Statistical Functions
    def _eval_sum(self, args: str, current_row: int, current_col: int) -> float:
        cells = self._parse_range(args.strip())
        return sum(self._get_cell_value(r, c) for r, c in cells)
    
    def _eval_average(self, args: str, current_row: int, current_col: int) -> float:
        cells = self._parse_range(args.strip())
        values = [self._get_cell_value(r, c) for r, c in cells]
        return sum(values) / len(values) if values else 0
    
    def _eval_count(self, args: str, current_row: int, current_col: int) -> int:
        cells = self._parse_range(args.strip())
        return sum(1 for r, c in cells if self.sheet.get_cell_data(r, c))
    
    def _eval_max(self, args: str, current_row: int, current_col: int) -> float:
        cells = self._parse_range(args.strip())
        values = [self._get_cell_value(r, c) for r, c in cells]
        return max(values) if values else 0
    
    def _eval_min(self, args: str, current_row: int, current_col: int) -> float:
        cells = self._parse_range(args.strip())
        values = [self._get_cell_value(r, c) for r, c in cells]
        return min(values) if values else 0
    
    # Logical Functions
    def _eval_if(self, args: str, current_row: int, current_col: int) -> Any:
        parts = [p.strip() for p in args.split(',')]
        if len(parts) != 3:
            raise ValueError("IF requires 3 arguments")
        
        condition = self._eval_expression(parts[0], current_row, current_col)
        if condition:
            return parts[1].strip('"\'')
        else:
            return parts[2].strip('"\'')
    
    def _eval_countif(self, args: str, current_row: int, current_col: int) -> int:
        parts = [p.strip() for p in args.split(',', 1)]
        if len(parts) != 2:
            raise ValueError("COUNTIF requires 2 arguments")
        
        cells = self._parse_range(parts[0])
        criteria = parts[1].strip('"\'')
        
        count = 0
        for r, c in cells:
            value = str(self.sheet.get_cell_data(r, c) or '')
            if criteria.startswith(('>', '<', '=', '!=')):
                try:
                    if eval(f"{float(value)}{criteria}", {"__builtins__": {}}, {}):
                        count += 1
                except:
                    pass
            else:
                if value == criteria:
                    count += 1
        return count
    
    def _eval_sumif(self, args: str, current_row: int, current_col: int) -> float:
        parts = [p.strip() for p in args.split(',')]
        if len(parts) < 2:
            raise ValueError("SUMIF requires at least 2 arguments")
        
        criteria_range = self._parse_range(parts[0])
        criteria = parts[1].strip('"\'')
        sum_range = self._parse_range(parts[2]) if len(parts) > 2 else criteria_range
        
        total = 0
        for i, (r, c) in enumerate(criteria_range):
            value = str(self.sheet.get_cell_data(r, c) or '')
            match = False
            if criteria.startswith(('>', '<', '=', '!=')):
                try:
                    if eval(f"{float(value)}{criteria}", {"__builtins__": {}}, {}):
                        match = True
                except:
                    pass
            else:
                if value == criteria:
                    match = True
            
            if match and i < len(sum_range):
                sr, sc = sum_range[i]
                total += self._get_cell_value(sr, sc)
        
        return total
    
    # Text Functions
    def _eval_concat(self, args: str, current_row: int, current_col: int) -> str:
        parts = [p.strip().strip('"\'') for p in args.split(',')]
        result = []
        for part in parts:
            if re.match(r'[A-Z]+\d+', part):
                r, c = self._parse_cell_ref(part)
                result.append(str(self.sheet.get_cell_data(r, c) or ''))
            else:
                result.append(part)
        return ''.join(result)
    
    def _eval_len(self, args: str, current_row: int, current_col: int) -> int:
        args = args.strip()
        if re.match(r'[A-Z]+\d+', args):
            r, c = self._parse_cell_ref(args)
            value = self.sheet.get_cell_data(r, c) or ''
        else:
            value = args.strip('"\'')
        return len(str(value))
    
    def _eval_upper(self, args: str, current_row: int, current_col: int) -> str:
        args = args.strip()
        if re.match(r'[A-Z]+\d+', args):
            r, c = self._parse_cell_ref(args)
            value = self.sheet.get_cell_data(r, c) or ''
        else:
            value = args.strip('"\'')
        return str(value).upper()
    
    def _eval_lower(self, args: str, current_row: int, current_col: int) -> str:
        args = args.strip()
        if re.match(r'[A-Z]+\d+', args):
            r, c = self._parse_cell_ref(args)
            value = self.sheet.get_cell_data(r, c) or ''
        else:
            value = args.strip('"\'')
        return str(value).lower()
    
    # Math Functions
    def _eval_sqrt(self, args: str, current_row: int, current_col: int) -> float:
        value = self._eval_expression(args.strip(), current_row, current_col)
        return math.sqrt(float(value))
    
    def _eval_power(self, args: str, current_row: int, current_col: int) -> float:
        parts = [p.strip() for p in args.split(',')]
        if len(parts) != 2:
            raise ValueError("POWER requires 2 arguments")
        base = self._eval_expression(parts[0], current_row, current_col)
        exp = self._eval_expression(parts[1], current_row, current_col)
        return pow(float(base), float(exp))
    
    def _eval_abs(self, args: str, current_row: int, current_col: int) -> float:
        value = self._eval_expression(args.strip(), current_row, current_col)
        return abs(float(value))
    
    def _eval_round(self, args: str, current_row: int, current_col: int) -> float:
        parts = [p.strip() for p in args.split(',')]
        number = self._eval_expression(parts[0], current_row, current_col)
        decimals = int(self._eval_expression(parts[1], current_row, current_col)) if len(parts) > 1 else 0
        return round(float(number), decimals)
    
    # Python Expression Evaluation
    def _eval_python(self, expr: str, current_row: int, current_col: int) -> Any:
        """Evaluate Python expression with math functions."""
        expr = self._substitute_refs(expr, current_row, current_col)
        safe_dict = {
            "__builtins__": {},
            "abs": abs, "round": round, "min": min, "max": max, "sum": sum,
            "pow": pow, "len": len,
            "sqrt": math.sqrt, "sin": math.sin, "cos": math.cos, "tan": math.tan,
            "log": math.log, "log10": math.log10, "exp": math.exp,
            "pi": math.pi, "e": math.e, "ceil": math.ceil, "floor": math.floor
        }
        return eval(expr, safe_dict, {})
    
    def _eval_expression(self, expr: str, current_row: int, current_col: int) -> Any:
        """Evaluate simple arithmetic expression."""
        expr = self._substitute_refs(expr, current_row, current_col)
        return eval(expr, {"__builtins__": {}}, {})
    
    def _substitute_refs(self, expr: str, current_row: int, current_col: int) -> str:
        """Replace cell references with their values."""
        def replace_ref(match):
            ref = match.group(0)
            try:
                r, c = self._parse_cell_ref(ref)
                return str(self._get_cell_value(r, c))
            except:
                return ref
        
        return re.sub(r'[A-Z]+\d+', replace_ref, expr)
