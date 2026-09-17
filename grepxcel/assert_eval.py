"""
Safe expression evaluator for grepxcel assert: rules.

assert: rows in a pattern file can carry cross-field validation expressions:

    assert:  | total == net + vat
    assert:  | discount_pct >= 0 and discount_pct <= 100
    assert:  | inv_date <= delivery_date  | date order check

The expression language is a strict subset of Python:

  Allowed:
    - Names that resolve to extracted field values (numbers / dates / strings)
    - Integer and float literals
    - String literals (single or double quoted)
    - Arithmetic:   +  -  *  /  //  %
    - Comparison:   ==  !=  <  <=  >  >=
    - Boolean:      and  or  not
    - Grouping:     ( )
    - Unary minus / plus

  NOT allowed (raises AssertParseError):
    - Any function call
    - Attribute access  (obj.attr)
    - Subscript / slice (obj[k])
    - Import / exec / eval
    - Walrus operator
    - Any other node type

Values that are None/missing cause the assertion to be **skipped** (the result
is None rather than True/False) — this preserves the semantics that an assertion
can only fail when the relevant fields were actually extracted.

  numeric coercion: if a field value is a string that looks like a number,
  it is cast to int or float before comparison so that ``total == net + vat``
  works with values extracted as strings (the common case with Excel).

Public API:

  parse_assert(expr: str) -> ast.Expression
    Parse and validate; raises AssertParseError on bad syntax.

  evaluate_assert(tree: ast.Expression, fields: dict[str, Any]) -> bool | None
    Evaluate against extracted field values.
    Returns True/False or None (skip — a field was missing/None).

  run_assert(expr: str, fields: dict[str, Any]) -> bool | None
    Convenience wrapper: parse + evaluate.
"""

from __future__ import annotations

import ast
import math
import operator
from typing import Any


class AssertParseError(ValueError):
    """Raised when an assert: expression uses unsupported syntax."""


class AssertScopeError(NameError):
    """Raised when an expression references an unknown field name."""


# ── Allowed AST node types ────────────────────────────────────────────────────

_ALLOWED_NODE_TYPES = frozenset({
    ast.Expression,
    ast.BoolOp, ast.And, ast.Or,
    ast.BinOp,
    ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod,
    ast.UnaryOp, ast.USub, ast.UAdd, ast.Not,
    ast.Compare,
    ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
    ast.Constant,
    ast.Name, ast.Load,
    ast.Attribute,   # dot-notation: loan.term_months → fields['loan.term_months']
})


def _validate_ast(node: ast.AST) -> None:
    """Recursively reject any AST node type not in the allow-list."""
    if type(node) not in _ALLOWED_NODE_TYPES:
        raise AssertParseError(
            f"Unsupported operation '{type(node).__name__}' in assert expression. "
            f"Only arithmetic, comparison, and boolean operations are allowed."
        )
    for child in ast.iter_child_nodes(node):
        _validate_ast(child)


def parse_assert(expr: str) -> ast.Expression:
    """Parse and security-validate an assert: expression.

    Args:
        expr: Expression string, e.g. ``'total == net + vat'``.

    Returns:
        Parsed ``ast.Expression`` ready for evaluation.

    Raises:
        AssertParseError: if the expression has syntax errors or uses
            unsupported node types.
    """
    expr = expr.strip()
    if not expr:
        raise AssertParseError("assert: expression is empty")
    try:
        tree = ast.parse(expr, mode='eval')
    except SyntaxError as exc:
        raise AssertParseError(
            f"Syntax error in assert expression {expr!r}: {exc}"
        ) from exc
    _validate_ast(tree)
    return tree


# ── Numeric coercion ──────────────────────────────────────────────────────────

def _coerce(value: Any) -> Any:
    """Try to coerce a string value to a number for arithmetic comparisons."""
    if isinstance(value, (int, float)) or value is None:
        return value
    s = str(value).strip()
    if not s:
        return None
    # Strip common currency/whitespace noise
    cleaned = s.replace(',', '').strip()
    try:
        iv = int(cleaned)
        return iv
    except ValueError:
        pass
    try:
        fv = float(cleaned)
        return fv
    except ValueError:
        pass
    return value  # keep as string for string comparisons


# ── Recursive evaluator ───────────────────────────────────────────────────────

_BINOPS = {
    ast.Add:      operator.add,
    ast.Sub:      operator.sub,
    ast.Mult:     operator.mul,
    ast.Div:      operator.truediv,
    ast.FloorDiv: operator.floordiv,
    ast.Mod:      operator.mod,
}

_CMPOPS = {
    ast.Eq:    operator.eq,
    ast.NotEq: operator.ne,
    ast.Lt:    operator.lt,
    ast.LtE:   operator.le,
    ast.Gt:    operator.gt,
    ast.GtE:   operator.ge,
}

# Sentinel to propagate "field missing" upwards
_MISSING = object()


def _eval(node: ast.AST, fields: dict[str, Any]) -> Any:
    """Recursively evaluate a validated assert AST node.

    Returns _MISSING if any referenced field is None/missing (so the whole
    expression is treated as skipped rather than failed).
    """
    if isinstance(node, ast.Expression):
        return _eval(node.body, fields)

    if isinstance(node, ast.Constant):
        return node.value

    if isinstance(node, ast.Name):
        name = node.id
        if name not in fields:
            return _MISSING
        val = _coerce(fields[name])
        if val is None:
            return _MISSING
        return val

    if isinstance(node, ast.Attribute):
        # Reconstruct dotted path: loan.term_months → fields['loan.term_months']
        parts = []
        n = node
        while isinstance(n, ast.Attribute):
            parts.append(n.attr)
            n = n.value
        if not isinstance(n, ast.Name):
            raise AssertParseError(
                "Unsupported attribute chain in assert expression — "
                "only simple dotted names (e.g. loan.term_months) are supported."
            )
        parts.append(n.id)
        parts.reverse()
        dotted = '.'.join(parts)
        if dotted not in fields:
            return _MISSING
        val = _coerce(fields[dotted])
        if val is None:
            return _MISSING
        return val

    if isinstance(node, ast.UnaryOp):
        operand = _eval(node.operand, fields)
        if operand is _MISSING:
            return _MISSING
        if isinstance(node.op, ast.USub):
            return -operand
        if isinstance(node.op, ast.UAdd):
            return +operand
        if isinstance(node.op, ast.Not):
            return not operand
        raise AssertParseError(f"Unsupported unary operator: {type(node.op).__name__}")

    if isinstance(node, ast.BinOp):
        left = _eval(node.left, fields)
        right = _eval(node.right, fields)
        if left is _MISSING or right is _MISSING:
            return _MISSING
        op_fn = _BINOPS.get(type(node.op))
        if op_fn is None:
            raise AssertParseError(f"Unsupported binary operator: {type(node.op).__name__}")
        try:
            return op_fn(left, right)
        except (TypeError, ZeroDivisionError):
            return _MISSING   # type mismatch or /0 → skip silently

    if isinstance(node, ast.Compare):
        left = _eval(node.left, fields)
        if left is _MISSING:
            return _MISSING
        result = True
        current = left
        for op, comparator in zip(node.ops, node.comparators):
            right = _eval(comparator, fields)
            if right is _MISSING:
                return _MISSING
            op_fn = _CMPOPS.get(type(op))
            if op_fn is None:
                raise AssertParseError(f"Unsupported comparison operator: {type(op).__name__}")
            try:
                if not op_fn(current, right):
                    result = False
                    break
            except TypeError:
                return _MISSING
            current = right
        return result

    if isinstance(node, ast.BoolOp):
        if isinstance(node.op, ast.And):
            for value in node.values:
                v = _eval(value, fields)
                if v is _MISSING:
                    return _MISSING
                if not v:
                    return False
            return True
        if isinstance(node.op, ast.Or):
            has_missing = False
            for value in node.values:
                v = _eval(value, fields)
                if v is _MISSING:
                    has_missing = True
                    continue
                if v:
                    return True
            return _MISSING if has_missing else False

    raise AssertParseError(f"Unexpected node type during evaluation: {type(node).__name__}")


def evaluate_assert(tree: ast.Expression, fields: dict[str, Any]) -> bool | None:
    """Evaluate a pre-parsed assert expression against extracted field values.

    Args:
        tree:   Validated ``ast.Expression`` from ``parse_assert()``.
        fields: Flat dict mapping field names to extracted values.

    Returns:
        ``True``  — assertion passed.
        ``False`` — assertion failed.
        ``None``  — assertion skipped (a referenced field was missing or null).
    """
    result = _eval(tree, fields)
    if result is _MISSING:
        return None
    return bool(result)


def run_assert(expr: str, fields: dict[str, Any]) -> bool | None:
    """Parse and evaluate an assert expression in one call.

    Args:
        expr:   Expression string.
        fields: Flat dict of extracted field values.

    Returns:
        ``True`` / ``False`` / ``None`` — see ``evaluate_assert()``.

    Raises:
        AssertParseError: if the expression is syntactically invalid or uses
            unsupported operations.
    """
    tree = parse_assert(expr)
    return evaluate_assert(tree, fields)
