"""Extract COUNTIFS and SUMIFS formulas from Excel workbooks."""
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from typing import Iterable, Iterator, List, Sequence

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

TARGET_FUNCTIONS = ("SUMIFS", "COUNTIFS")


@dataclass
class FormulaCell:
    sheet: str
    address: str
    formula: str


@dataclass
class FunctionCall:
    func_name: str
    args: List[str]
    raw_text: str
    offset: int

    def __str__(self) -> str:  # pragma: no cover - for debugging
        return f"{self.func_name}({', '.join(self.args)})"


def iter_formula_cells(ws: Worksheet) -> Iterator[FormulaCell]:
    """Yield every formula cell in a worksheet."""
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                yield FormulaCell(ws.title, cell.coordinate, cell.value)


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


def walk_workbook(path: str) -> Iterable[tuple[FormulaCell, Sequence[FunctionCall]]]:
    workbook = load_workbook(path, data_only=False, read_only=False)
    for ws in workbook.worksheets:
        for cell in iter_formula_cells(ws):
            calls = find_target_calls(cell.formula)
            if calls:
                yield cell, calls


def main(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="Path to the Excel workbook to inspect")
    args = parser.parse_args(argv)

    for cell, calls in walk_workbook(args.workbook):
        print(f"{cell.sheet}!{cell.address} -> {cell.formula}")
        for call in calls:
            print(f"  found {call.raw_text}")
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main(None))
