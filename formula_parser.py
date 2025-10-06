"""Basic Excel formula parser producing an abstract syntax tree.

This module relies on :class:`openpyxl.formula.tokenizer.Tokenizer` for lexical
analysis and implements a small recursive-descent parser to turn formulas into a
structured representation that downstream translators can target (pandas,
NumPy, PyTorch, etc.)."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator, List, Optional

from openpyxl.formula.tokenizer import Tokenizer, Token


class FormulaParseError(ValueError):
    """Raised when a formula cannot be parsed into an AST."""


@dataclass
class FormulaNode:
    """Base class for all AST nodes."""


@dataclass
class NumberNode(FormulaNode):
    value: float
    text: str


@dataclass
class StringNode(FormulaNode):
    value: str


@dataclass
class BooleanNode(FormulaNode):
    value: bool


@dataclass
class ErrorNode(FormulaNode):
    value: str


@dataclass
class ReferenceNode(FormulaNode):
    sheet: Optional[str]
    reference: str
    original: str


@dataclass
class NameNode(FormulaNode):
    name: str


@dataclass
class UnaryOpNode(FormulaNode):
    operator: str
    operand: FormulaNode
    postfix: bool = False


@dataclass
class BinaryOpNode(FormulaNode):
    operator: str
    left: FormulaNode
    right: FormulaNode


@dataclass
class FunctionNode(FormulaNode):
    name: str
    args: List[FormulaNode]


PRECEDENCE = {
    '^': 5,
    '*': 4,
    '/': 4,
    '+': 3,
    '-': 3,
    '&': 2,
    '=': 1,
    '<>': 1,
    '<=': 1,
    '>=': 1,
    '<': 1,
    '>': 1,
}

RIGHT_ASSOCIATIVE = {'^'}


class FormulaParser:
    """Recursive-descent / Pratt style parser over openpyxl tokens."""

    def __init__(self, formula: str) -> None:
        cleaned = formula.strip()
        if not cleaned:
            raise FormulaParseError("Empty formula")
        if not cleaned.startswith("="):
            cleaned = "=" + cleaned
        tokenizer = Tokenizer(cleaned)
        self.tokens: List[Token] = [tok for tok in tokenizer.items if tok.type != 'WHITE-SPACE']
        if not self.tokens:
            raise FormulaParseError("Formula produced no tokens")
        self.index = 0

    def parse(self) -> FormulaNode:
        node = self.parse_expression(0)
        if self.current() is not None:
            raise FormulaParseError(f"Unexpected token {self.current().value!r} at end of formula")
        return node

    # Token helpers -----------------------------------------------------

    def current(self) -> Optional[Token]:
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def advance(self) -> Optional[Token]:
        tok = self.current()
        if tok is not None:
            self.index += 1
        return tok

    def expect(self, ttype: str, subtype: Optional[str] = None) -> Token:
        tok = self.current()
        if tok is None:
            raise FormulaParseError(f"Expected token {ttype} but reached end of formula")
        if tok.type != ttype or (subtype is not None and tok.subtype != subtype):
            raise FormulaParseError(f"Expected token {ttype}/{subtype}, got {tok.type}/{tok.subtype}: {tok.value!r}")
        self.index += 1
        return tok

    # Parsing -----------------------------------------------------------

    def parse_expression(self, min_precedence: int) -> FormulaNode:
        node = self.parse_prefix()
        while True:
            tok = self.current()
            if tok is None:
                break
            if tok.type == 'OPERATOR-POSTFIX':
                self.advance()
                node = UnaryOpNode(tok.value, node, postfix=True)
                continue
            if tok.type != 'OPERATOR-INFIX':
                break
            operator = tok.value
            precedence = PRECEDENCE.get(operator)
            if precedence is None or precedence < min_precedence:
                break
            self.advance()
            next_min = precedence if operator in RIGHT_ASSOCIATIVE else precedence + 1
            rhs = self.parse_expression(next_min)
            node = BinaryOpNode(operator, node, rhs)
        return node

    def parse_prefix(self) -> FormulaNode:
        tok = self.current()
        if tok is None:
            raise FormulaParseError("Unexpected end of formula")
        if tok.type == 'OPERAND':
            self.advance()
            return self.parse_operand(tok)
        if tok.type == 'OPERATOR-PREFIX':
            self.advance()
            operand = self.parse_expression(5)
            return UnaryOpNode(tok.value, operand, postfix=False)
        if tok.type == 'PAREN' and tok.subtype == 'OPEN':
            self.advance()
            expr = self.parse_expression(0)
            self.expect('PAREN', 'CLOSE')
            return expr
        if tok.type == 'FUNC' and tok.subtype == 'OPEN':
            return self.parse_function(tok)
        if tok.type == 'ARRAY' and tok.subtype == 'OPEN':
            # Array literals (e.g., {1,2,3}) – treat as name for now.
            self.advance()
            values: List[FormulaNode] = []
            while True:
                values.append(self.parse_expression(0))
                tok = self.current()
                if tok and tok.type == 'SEP' and tok.subtype == 'ROW':
                    self.advance()
                    continue
                if tok and tok.type == 'SEP' and tok.subtype == 'ARG':
                    self.advance()
                    continue
                if tok and tok.type == 'ARRAY' and tok.subtype == 'CLOSE':
                    self.advance()
                    break
                if tok is None:
                    raise FormulaParseError("Unterminated array literal")
            # Represent arrays as function-like node for now
            return FunctionNode('ARRAY', values)
        raise FormulaParseError(f"Unsupported token {tok.type}/{tok.subtype}: {tok.value!r}")

    def parse_function(self, token: Token) -> FormulaNode:
        name = token.value[:-1]  # strip trailing '('
        self.advance()
        args: List[FormulaNode] = []
        # Handle zero-argument functions explicitly
        if self.current() and self.current().type == 'FUNC' and self.current().subtype == 'CLOSE':
            self.advance()
            return FunctionNode(name, args)
        while True:
            args.append(self.parse_expression(0))
            tok = self.current()
            if tok and tok.type == 'SEP' and tok.subtype == 'ARG':
                self.advance()
                continue
            if tok and tok.type == 'FUNC' and tok.subtype == 'CLOSE':
                self.advance()
                break
            if tok is None:
                raise FormulaParseError(f"Function {name} missing closing parenthesis")
            raise FormulaParseError(f"Unexpected token {tok.value!r} inside function {name}")
        return FunctionNode(name, args)

    def parse_operand(self, token: Token) -> FormulaNode:
        subtype = token.subtype
        value = token.value
        if subtype == 'NUMBER':
            try:
                number = float(value)
            except ValueError as exc:
                raise FormulaParseError(f"Invalid numeric literal {value!r}") from exc
            return NumberNode(number, value)
        if subtype == 'TEXT':
            if value.startswith('"') and value.endswith('"'):
                inner = value[1:-1].replace('""', '"')
            else:
                inner = value
            return StringNode(inner)
        if subtype == 'LOGICAL':
            return BooleanNode(value.upper() == 'TRUE')
        if subtype == 'ERROR':
            return ErrorNode(value)
        if subtype == 'RANGE':
            sheet, ref = split_reference(value)
            return ReferenceNode(sheet, ref, value)
        return NameNode(value)


def split_reference(value: str) -> tuple[Optional[str], str]:
    """Split a reference token into sheet and address components."""
    text = value
    sheet: Optional[str] = None
    if '!' in text:
        sheet_part, address = text.split('!', 1)
        sheet = unescape_sheet_name(sheet_part)
        return sheet, address
    return None, text


def unescape_sheet_name(sheet: str) -> str:
    if sheet.startswith("'") and sheet.endswith("'"):
        inner = sheet[1:-1]
        return inner.replace("''", "'")
    return sheet


def parse_formula(formula: str) -> FormulaNode:
    parser = FormulaParser(formula)
    return parser.parse()


# Convenience -----------------------------------------------------------------


def iter_tokens(formula: str) -> Iterator[Token]:
    cleaned = formula.strip()
    if not cleaned.startswith('='):
        cleaned = '=' + cleaned
    tok = Tokenizer(cleaned)
    for item in tok.items:
        if item.type != 'WHITE-SPACE':
            yield item


__all__ = [
    'FormulaNode',
    'NumberNode',
    'StringNode',
    'BooleanNode',
    'ErrorNode',
    'ReferenceNode',
    'NameNode',
    'UnaryOpNode',
    'BinaryOpNode',
    'FunctionNode',
    'FormulaParseError',
    'parse_formula',
    'iter_tokens',
]
