"""Translate COUNTIFS and SUMIFS formulas into Pandas snippets."""
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from typing import Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

TARGET_FUNCTIONS = ("SUMIFS", "COUNTIFS")


@dataclass
class FormulaCell:
    sheet: str
    address: str
    formula: str
    worksheet: Worksheet


@dataclass
class FunctionCall:
    func_name: str
    args: List[str]
    raw_text: str
    offset: int

    def __str__(self) -> str:  # pragma: no cover - for debugging
        return f"{self.func_name}({', '.join(self.args)})"


@dataclass
class FilterExpression:
    sheet: str
    expression: str
    warnings: List[str] = field(default_factory=list)


@dataclass
class Translation:
    cell: FormulaCell
    call: FunctionCall
    expression: Optional[str]
    warnings: List[str] = field(default_factory=list)


def iter_formula_cells(ws: Worksheet) -> Iterator[FormulaCell]:
    """Yield every formula cell in a worksheet."""
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                yield FormulaCell(ws.title, cell.coordinate, cell.value, ws)


def strip_array_braces(formula: str) -> str:
    text = formula.strip()
    if text.startswith("{") and text.endswith("}"):
        return text[1:-1]
    return text


def split_top_level_args(arg_string: str) -> List[str]:
    """Split a comma-delimited argument string without breaking nested calls."""
    parts: List[str] = []
    token: List[str] = []
    depth = 0
    in_string = False
    i = 0
    while i < len(arg_string):
        ch = arg_string[i]
        if ch == '"':
            token.append(ch)
            if in_string:
                # Excel escapes quotes by doubling them
                if i + 1 < len(arg_string) and arg_string[i + 1] == '"':
                    token.append('"')
                    i += 2
                    continue
                in_string = False
                i += 1
                continue
            in_string = True
            i += 1
            continue
        if not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(depth - 1, 0)
            elif ch == "," and depth == 0:
                parts.append("".join(token).strip())
                token.clear()
                i += 1
                continue
        token.append(ch)
        i += 1
    if token:
        parts.append("".join(token).strip())
    return parts


def parse_string_literal(token: str) -> Optional[str]:
    if len(token) < 2 or not token.startswith('"') or not token.endswith('"'):
        return None
    inner = token[1:-1]
    return inner.replace('""', '"')


_identifier_tail = re.compile(r"[A-Z0-9_\.]", re.IGNORECASE)


def find_target_calls(formula: str) -> List[FunctionCall]:
    """Locate COUNTIFS/SUMIFS calls within a formula string."""
    cleaned = strip_array_braces(formula)
    upper = cleaned.upper()
    calls: List[FunctionCall] = []
    i = 0
    while i < len(cleaned):
        matched_name = None
        for name in TARGET_FUNCTIONS:
            upper_name = name
            if upper.startswith(upper_name, i):
                prev_idx = i - 1
                if prev_idx >= 0 and _identifier_tail.match(upper[prev_idx]):
                    continue
                next_idx = i + len(name)
                if next_idx >= len(cleaned) or cleaned[next_idx] != "(":
                    continue
                matched_name = name
                break
        if not matched_name:
            i += 1
            continue
        call_text, end_idx = _extract_function_call(cleaned, i, len(matched_name))
        if call_text is None:
            i += len(matched_name)
            continue
        func_name = matched_name
        arg_segment = call_text[len(func_name) + 1 : -1]
        args = split_top_level_args(arg_segment)
        calls.append(FunctionCall(func_name, args, call_text, i))
        i = end_idx
    return calls


def _extract_function_call(formula: str, start: int, name_len: int) -> tuple[str | None, int]:
    idx = start + name_len
    if idx >= len(formula) or formula[idx] != "(":
        return None, idx
    idx += 1
    depth = 1
    in_string = False
    while idx < len(formula):
        ch = formula[idx]
        if ch == '"':
            if in_string and idx + 1 < len(formula) and formula[idx + 1] == '"':
                idx += 2
                continue
            in_string = not in_string
            idx += 1
            continue
        if not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return formula[start : idx + 1], idx + 1
        idx += 1
    return None, idx


def build_header_map(workbook) -> Dict[str, Dict[str, str]]:
    mapping: Dict[str, Dict[str, str]] = {}
    for ws in workbook.worksheets:
        header: Dict[str, str] = {}
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=False), None)
        if header_row:
            for cell in header_row:
                value = cell.value
                if value is None or str(value).strip() == "":
                    continue
                header[cell.column_letter] = str(value)
        mapping[ws.title] = header
    return mapping


def df_reference(sheet: str) -> str:
    return f"dfs[{sheet!r}]"


def split_range_reference(reference: str, default_sheet: str) -> Optional[Tuple[str, str]]:
    token = reference.strip()
    if not token:
        return None
    sheet = default_sheet
    if "!" not in token:
        return sheet, token
    if token[0] == "'":
        idx = 1
        builder: List[str] = []
        while idx < len(token):
            ch = token[idx]
            if ch == "'":
                if idx + 1 < len(token) and token[idx + 1] == "'":
                    builder.append("'")
                    idx += 2
                    continue
                idx += 1
                break
            builder.append(ch)
            idx += 1
        sheet = "".join(builder)
        if idx < len(token) and token[idx] == "!":
            idx += 1
        ref_part = token[idx:]
        return sheet, ref_part
    sheet_part, ref_part = token.split("!", 1)
    return sheet_part, ref_part


COL_RE = re.compile(r"^([A-Z]+)(\d+)?$", re.IGNORECASE)
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$", re.IGNORECASE)


def range_columns(range_token: str) -> Optional[List[str]]:
    token = range_token.replace("$", "")
    if "[" in token:
        return None
    if ":" in token:
        left, right = token.split(":", 1)
    else:
        left = right = token
    left_match = COL_RE.match(left)
    right_match = COL_RE.match(right)
    if not left_match or not right_match:
        return None
    start_idx = column_index_from_string(left_match.group(1))
    end_idx = column_index_from_string(right_match.group(1))
    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx
    return [get_column_letter(idx) for idx in range(start_idx, end_idx + 1)]


def range_to_column(
    range_ref: str,
    header_map: Dict[str, Dict[str, str]],
    default_sheet: str,
) -> Tuple[Optional[Tuple[str, str]], Optional[str]]:
    parsed = split_range_reference(range_ref, default_sheet)
    if not parsed:
        return None, f"Could not parse range reference: {range_ref!r}"
    sheet, ref = parsed
    columns = range_columns(ref)
    if not columns:
        return None, f"Unsupported range reference: {range_ref}"
    if len(columns) != 1:
        return None, f"Only single-column ranges are supported (saw {range_ref})"
    column_letter = columns[0]
    header = header_map.get(sheet, {}).get(column_letter)
    column_name = header if header is not None else column_letter
    return (sheet, column_name), None


def resolve_cell_reference(
    token: str,
    workbook,
    default_sheet: str,
) -> Tuple[bool, Optional[object]]:
    parsed = split_range_reference(token, default_sheet)
    if not parsed:
        return False, None
    sheet, ref = parsed
    if ":" in ref:
        return False, None
    cleaned = ref.replace("$", "")
    cell_match = CELL_RE.match(cleaned)
    if not cell_match:
        return False, None
    sheet_name = sheet
    if sheet_name not in workbook.sheetnames:
        return False, None
    ws = workbook[sheet_name]
    try:
        value = ws[cleaned].value
    except ValueError:
        return False, None
    return True, value


def coerce_literal(text: str) -> object:
    stripped = text.strip()
    if stripped.upper() in {"TRUE", "FALSE"}:
        return stripped.upper() == "TRUE"
    if stripped == "":
        return ""
    try:
        if "." in stripped:
            return float(stripped)
        return int(stripped)
    except ValueError:
        return stripped


def format_literal(value: object) -> str:
    if isinstance(value, str):
        return repr(value)
    if value is None:
        return "None"
    if isinstance(value, bool):
        return "True" if value else "False"
    return repr(value)


def excel_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


OPERATOR_MAP = [
    ("<>", "!="),
    (">=", ">="),
    ("<=", "<="),
    (">", ">"),
    ("<", "<"),
    ("=", "=="),
]

NUMBER_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def contains_wildcard(text: str) -> bool:
    return any(ch in text for ch in ("*", "?"))


def wildcard_to_regex(pattern: str) -> str:
    regex_parts: List[str] = ["^"]
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            regex_parts.append(".*")
        elif ch == "?":
            regex_parts.append(".")
        else:
            regex_parts.append(re.escape(ch))
        i += 1
    regex_parts.append("$")
    return "".join(regex_parts)


def build_filter_from_operator(lhs: str, op: str, payload: str) -> Tuple[Optional[str], List[str], Optional[str]]:
    warnings: List[str] = []
    if contains_wildcard(payload):
        if op not in {"==", "!="}:
            return None, warnings, f"Wildcards with operator {op} are not supported"
        regex = wildcard_to_regex(payload)
        expr = f"{lhs}.astype(str).str.match({regex!r})"
        if op == "!=":
            expr = f"~({expr})"
        return expr, warnings, None
    literal = coerce_literal(payload)
    if literal == "" and op == "!=":
        expr = f"~({lhs}.isna() | ({lhs} == \"\"))"
        return expr, warnings, None
    return f"{lhs} {op} {format_literal(literal)}", warnings, None


def split_concat_tokens(expr: str) -> List[str]:
    parts: List[str] = []
    token: List[str] = []
    depth = 0
    in_string = False
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            token.append(ch)
            if in_string:
                if i + 1 < len(expr) and expr[i + 1] == '"':
                    token.append('"')
                    i += 2
                    continue
                in_string = False
                i += 1
                continue
            in_string = True
            i += 1
            continue
        if not in_string:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(depth - 1, 0)
            elif ch == "&" and depth == 0:
                parts.append("".join(token).strip())
                token.clear()
                i += 1
                continue
        token.append(ch)
        i += 1
    if token:
        parts.append("".join(token).strip())
    return parts


def evaluate_concat_token(
    token: str,
    workbook,
    default_sheet: str,
) -> Tuple[Optional[str], List[str], Optional[str]]:
    warnings: List[str] = []
    stripped = token.strip()
    if stripped == "":
        return "", warnings, None
    literal = parse_string_literal(stripped)
    if literal is not None:
        return literal, warnings, None
    if NUMBER_RE.match(stripped):
        return str(coerce_literal(stripped)), warnings, None
    upper = stripped.upper()
    if upper in {"TRUE", "FALSE"}:
        return ("TRUE" if upper == "TRUE" else "FALSE"), warnings, None
    matched, value = resolve_cell_reference(stripped, workbook, default_sheet)
    if matched:
        return excel_text(value), warnings, None
    return None, warnings, f"Unsupported concatenation token: {token}"


def try_evaluate_concat(
    expr: str,
    workbook,
    default_sheet: str,
) -> Tuple[Optional[str], List[str], Optional[str]]:
    tokens = split_concat_tokens(expr)
    if len(tokens) <= 1:
        return None, [], None
    warnings: List[str] = []
    pieces: List[str] = []
    for token in tokens:
        value, token_warnings, error = evaluate_concat_token(token, workbook, default_sheet)
        warnings.extend(token_warnings)
        if error:
            return None, warnings, error
        if value is None:
            return None, warnings, f"Unable to evaluate concatenation token: {token}"
        pieces.append(value)
    return "".join(pieces), warnings, None


def build_filter_expression_from_string(
    lhs: str,
    value: str,
) -> Tuple[Optional[str], List[str], Optional[str]]:
    warnings: List[str] = []
    for token, py_op in OPERATOR_MAP:
        if value.startswith(token):
            remainder = value[len(token) :]
            expr, extra_warnings, error = build_filter_from_operator(lhs, py_op, remainder)
            warnings.extend(extra_warnings)
            return expr, warnings, error
    expr, extra_warnings, error = build_filter_from_operator(lhs, "==", value)
    warnings.extend(extra_warnings)
    return expr, warnings, error


def build_filter_expression(
    lhs: str,
    criteria: str,
    workbook,
    default_sheet: str,
) -> Tuple[Optional[str], List[str], Optional[str]]:
    crit = criteria.strip()
    warnings: List[str] = []
    if not crit:
        return None, warnings, "Empty criteria"
    concat_value, concat_warnings, concat_error = try_evaluate_concat(crit, workbook, default_sheet)
    warnings.extend(concat_warnings)
    if concat_error:
        return None, warnings, concat_error
    if concat_value is not None:
        expr, extra_warnings, error = build_filter_expression_from_string(lhs, concat_value)
        warnings.extend(extra_warnings)
        return expr, warnings, error
    literal = parse_string_literal(crit)
    if literal is not None:
        expr, extra_warnings, error = build_filter_expression_from_string(lhs, literal)
        warnings.extend(extra_warnings)
        return expr, warnings, error
    if NUMBER_RE.match(crit):
        return f"{lhs} == {coerce_literal(crit)}", warnings, None
    matched, value = resolve_cell_reference(crit, workbook, default_sheet)
    if matched:
        return f"{lhs} == {format_literal(value)}", warnings, None
    if crit.upper() in {"TRUE", "FALSE"}:
        literal = "True" if crit.upper() == "TRUE" else "False"
        return f"{lhs} == {literal}", warnings, None
    return None, warnings, f"Unsupported criteria: {criteria}"


def to_filter_expr(
    range_ref: str,
    criteria: str,
    workbook,
    header_map: Dict[str, Dict[str, str]],
    default_sheet: str,
) -> Tuple[Optional[FilterExpression], Optional[str]]:
    range_result, error = range_to_column(range_ref, header_map, default_sheet)
    if error:
        return None, error
    sheet, column_name = range_result
    lhs = f"{df_reference(sheet)}[{column_name!r}]"
    expression, warnings, crit_error = build_filter_expression(lhs, criteria, workbook, default_sheet)
    if crit_error:
        return None, crit_error
    return FilterExpression(sheet, expression, warnings), None


def combine_filters(filters: List[FilterExpression]) -> str:
    if not filters:
        return "slice(None)"
    joined = " & ".join(f"({flt.expression})" for flt in filters)
    return joined


def build_sumifs_translation(
    cell: FormulaCell,
    call: FunctionCall,
    workbook,
    header_map: Dict[str, Dict[str, str]],
) -> Tuple[Optional[str], List[str]]:
    warnings: List[str] = []
    if not call.args:
        return None, ["SUMIFS requires at least a sum range"]
    sum_range = call.args[0]
    sum_result, error = range_to_column(sum_range, header_map, cell.sheet)
    if error:
        return None, [error]
    sum_sheet, sum_column = sum_result
    criteria_args = call.args[1:]
    if len(criteria_args) % 2 != 0:
        return None, ["SUMIFS expects range/criteria pairs"]
    filters: List[FilterExpression] = []
    for range_ref, criteria in zip(criteria_args[::2], criteria_args[1::2]):
        filt, filt_error = to_filter_expr(range_ref, criteria, workbook, header_map, cell.sheet)
        if filt_error:
            return None, [filt_error]
        if filt is None:
            return None, ["Failed to build filter expression"]
        warnings.extend(filt.warnings)
        filters.append(filt)
    df_ref = df_reference(sum_sheet)
    if filters:
        mask_expr = combine_filters(filters)
        rows_expr = f"{df_ref}.loc[{mask_expr}]"
        expression = f"{rows_expr}[{sum_column!r}].sum()"
    else:
        expression = f"{df_ref}[{sum_column!r}].sum()"
    return expression, warnings


def build_countifs_translation(
    cell: FormulaCell,
    call: FunctionCall,
    workbook,
    header_map: Dict[str, Dict[str, str]],
) -> Tuple[Optional[str], List[str]]:
    warnings: List[str] = []
    if not call.args or len(call.args) % 2 != 0:
        return None, ["COUNTIFS expects an even number of arguments"]
    first_range = call.args[0]
    base_result, error = range_to_column(first_range, header_map, cell.sheet)
    if error:
        return None, [error]
    base_sheet, _ = base_result
    filters: List[FilterExpression] = []
    for range_ref, criteria in zip(call.args[::2], call.args[1::2]):
        filt, filt_error = to_filter_expr(range_ref, criteria, workbook, header_map, cell.sheet)
        if filt_error:
            return None, [filt_error]
        if filt is None:
            return None, ["Failed to build filter expression"]
        warnings.extend(filt.warnings)
        filters.append(filt)
    mask_expr = combine_filters(filters)
    df_ref = df_reference(base_sheet)
    expression = f"{df_ref}.loc[{mask_expr}].shape[0]"
    return expression, warnings


def translate_call(
    cell: FormulaCell,
    call: FunctionCall,
    workbook,
    header_map: Dict[str, Dict[str, str]],
) -> Translation:
    if call.func_name == "SUMIFS":
        expression, warnings = build_sumifs_translation(cell, call, workbook, header_map)
    else:
        expression, warnings = build_countifs_translation(cell, call, workbook, header_map)
    return Translation(cell, call, expression, warnings)


def translate_workbook(workbook) -> Iterator[Translation]:
    header_map = build_header_map(workbook)
    for ws in workbook.worksheets:
        for cell in iter_formula_cells(ws):
            calls = find_target_calls(cell.formula)
            for call in calls:
                yield translate_call(cell, call, workbook, header_map)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="Path to the Excel workbook to inspect")
    args = parser.parse_args(list(argv) if argv is not None else None)

    workbook = load_workbook(args.workbook, data_only=False, read_only=False)
    found = False
    for translation in translate_workbook(workbook):
        found = True
        cell = translation.cell
        call = translation.call
        print(f"{cell.sheet}!{cell.address}: {call.raw_text}")
        if translation.expression:
            print(f"  pandas -> {translation.expression}")
        else:
            print("  pandas -> <unable to build translation>")
        for warning in translation.warnings:
            print(f"  note   -> {warning}")
    if not found:
        print("No COUNTIFS or SUMIFS formulas found.")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
