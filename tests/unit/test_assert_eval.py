"""Unit tests for grepxcel.assert_eval — safe expression evaluator."""

import pytest

from grepxcel.assert_eval import (
    AssertParseError,
    evaluate_assert,
    parse_assert,
    run_assert,
)


# ── parse_assert — syntax validation ─────────────────────────────────────────

class TestParseAssert:
    def test_empty_raises(self):
        with pytest.raises(AssertParseError, match='empty'):
            parse_assert('')

    def test_blank_raises(self):
        with pytest.raises(AssertParseError, match='empty'):
            parse_assert('   ')

    def test_syntax_error_raises(self):
        with pytest.raises(AssertParseError, match='Syntax error'):
            parse_assert('total ==')

    def test_valid_comparison(self):
        tree = parse_assert('total == net + vat')
        assert tree is not None

    def test_valid_bool_expr(self):
        tree = parse_assert('a > 0 and b < 100')
        assert tree is not None

    def test_function_call_rejected(self):
        with pytest.raises(AssertParseError, match='Unsupported operation'):
            parse_assert('len(name) > 0')

    def test_attribute_access_rejected(self):
        with pytest.raises(AssertParseError, match='Unsupported operation'):
            parse_assert('obj.attr == 1')

    def test_subscript_rejected(self):
        with pytest.raises(AssertParseError, match='Unsupported operation'):
            parse_assert('items[0] == 1')

    def test_list_rejected(self):
        with pytest.raises(AssertParseError, match='Unsupported operation'):
            parse_assert('[1, 2, 3]')

    def test_dict_rejected(self):
        with pytest.raises(AssertParseError, match='Unsupported operation'):
            parse_assert('{"a": 1}')

    def test_walrus_rejected(self):
        with pytest.raises(AssertParseError):
            parse_assert('(x := 5) == 5')

    def test_import_rejected(self):
        with pytest.raises(AssertParseError):
            parse_assert('__import__("os").getcwd()')


# ── evaluate_assert — core evaluation ────────────────────────────────────────

class TestEvaluateAssert:
    def _eval(self, expr, fields):
        tree = parse_assert(expr)
        return evaluate_assert(tree, fields)

    # --- Missing fields produce None (skip) ---

    def test_missing_field_returns_none(self):
        result = self._eval('total == 100', {})
        assert result is None

    def test_none_value_returns_none(self):
        result = self._eval('total == 100', {'total': None})
        assert result is None

    def test_missing_operand_in_arithmetic_returns_none(self):
        result = self._eval('total == net + vat', {'net': 80})
        # vat is missing
        assert result is None

    # --- Arithmetic comparisons ---

    def test_addition_pass(self):
        result = self._eval('total == net + vat', {'total': 110, 'net': 100, 'vat': 10})
        assert result is True

    def test_addition_fail(self):
        result = self._eval('total == net + vat', {'total': 999, 'net': 100, 'vat': 10})
        assert result is False

    def test_subtraction(self):
        result = self._eval('net == total - vat', {'net': 90, 'total': 100, 'vat': 10})
        assert result is True

    def test_multiplication(self):
        result = self._eval('total == qty * price', {'total': 50, 'qty': 5, 'price': 10})
        assert result is True

    def test_division(self):
        result = self._eval('unit_price == total / qty', {'unit_price': 25, 'total': 100, 'qty': 4})
        assert result is True

    def test_floor_division(self):
        result = self._eval('x == y // 3', {'x': 3, 'y': 10})
        assert result is True

    def test_modulo(self):
        result = self._eval('x == y % 3', {'x': 1, 'y': 10})
        assert result is True

    # --- Comparison operators ---

    def test_ne(self):
        assert self._eval('a != b', {'a': 1, 'b': 2}) is True

    def test_lt(self):
        assert self._eval('a < b', {'a': 1, 'b': 2}) is True
        assert self._eval('a < b', {'a': 2, 'b': 1}) is False

    def test_le(self):
        assert self._eval('a <= b', {'a': 1, 'b': 1}) is True
        assert self._eval('a <= b', {'a': 2, 'b': 1}) is False

    def test_gt(self):
        assert self._eval('a > b', {'a': 5, 'b': 3}) is True

    def test_ge(self):
        assert self._eval('a >= b', {'a': 3, 'b': 3}) is True

    # --- Boolean operators ---

    def test_and_both_true(self):
        assert self._eval('a > 0 and b > 0', {'a': 1, 'b': 1}) is True

    def test_and_one_false(self):
        assert self._eval('a > 0 and b > 0', {'a': 0, 'b': 1}) is False

    def test_or_one_true(self):
        assert self._eval('a > 0 or b > 0', {'a': 0, 'b': 1}) is True

    def test_or_all_false(self):
        assert self._eval('a > 0 or b > 0', {'a': 0, 'b': 0}) is False

    def test_not(self):
        assert self._eval('not a > 10', {'a': 5}) is True

    # --- Unary operators ---

    def test_unary_minus(self):
        assert self._eval('a == -b', {'a': -5, 'b': 5}) is True

    def test_unary_plus(self):
        assert self._eval('a == +b', {'a': 5, 'b': 5}) is True

    # --- String coercion ---

    def test_string_number_coerced(self):
        # Values extracted from Excel often come as strings
        result = self._eval('total == net + vat', {'total': '110', 'net': '100', 'vat': '10'})
        assert result is True

    def test_string_with_commas_coerced(self):
        result = self._eval('a > 1000', {'a': '1,500'})
        assert result is True

    def test_string_comparison(self):
        # Non-numeric strings are compared as strings
        result = self._eval('status == "ok"', {'status': 'ok'})
        assert result is True

    # --- Zero division ---

    def test_zero_division_returns_none(self):
        result = self._eval('x == a / b', {'x': 0, 'a': 5, 'b': 0})
        assert result is None

    # --- Chained comparisons ---

    def test_chained_comparison(self):
        result = self._eval('0 <= discount_pct and discount_pct <= 100',
                            {'discount_pct': 15})
        assert result is True

    # --- Constant expressions ---

    def test_constant_pass(self):
        assert self._eval('1 == 1', {}) is True

    def test_constant_fail(self):
        assert self._eval('1 == 2', {}) is False


# ── run_assert — combined parse + evaluate ───────────────────────────────────

class TestRunAssert:
    def test_pass(self):
        assert run_assert('total == net + vat', {'total': 110, 'net': 100, 'vat': 10}) is True

    def test_fail(self):
        assert run_assert('total == net + vat', {'total': 999, 'net': 100, 'vat': 10}) is False

    def test_skip(self):
        assert run_assert('total == net + vat', {'net': 100}) is None

    def test_bad_expr_raises(self):
        with pytest.raises(AssertParseError):
            run_assert('len(x) > 0', {'x': 'hello'})


# ── Integration: pattern_parser parses assert: rows ──────────────────────────

class TestAssertInPatternParser:
    """Verify that assert: rows are correctly parsed and stored on the parser."""

    def _make_csv(self, tmp_path, lines: list[str]) -> str:
        """Write a minimal valid pattern CSV and return its path."""
        path = tmp_path / 'pattern.csv'
        path.write_text('\n'.join(lines))
        return str(path)

    def test_no_assert_rows(self, tmp_path):
        content = [
            'var:,amount,string,.*',
            'START:',
            'cell:next,amount',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        from grepxcel.pattern_parser import PatternParser
        pp = PatternParser()
        pp.parse(path)
        assert pp.assert_rules == []

    def test_single_assert_row(self, tmp_path):
        content = [
            'var:,total,string,.*',
            'var:,net,string,.*',
            'var:,vat,string,.*',
            'assert:,total == net + vat',
            'START:',
            'cell:next,total',
            'cell:next,net',
            'cell:next,vat',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        from grepxcel.pattern_parser import PatternParser
        pp = PatternParser()
        pp.parse(path)
        assert len(pp.assert_rules) == 1
        assert pp.assert_rules[0].expression == 'total == net + vat'
        assert pp.assert_rules[0].message == ''

    def test_assert_with_custom_message(self, tmp_path):
        content = [
            'var:,total,string,.*',
            'var:,net,string,.*',
            'var:,vat,string,.*',
            'assert:,total == net + vat,totals must balance',
            'START:',
            'cell:next,total',
            'cell:next,net',
            'cell:next,vat',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        from grepxcel.pattern_parser import PatternParser
        pp = PatternParser()
        pp.parse(path)
        assert pp.assert_rules[0].message == 'totals must balance'

    def test_invalid_assert_expression_raises(self, tmp_path):
        from grepxcel.pattern_parser import PatternError, PatternParser
        content = [
            'var:,total,string,.*',
            'assert:,len(total) > 0',
            'START:',
            'cell:next,total',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        pp = PatternParser()
        with pytest.raises(PatternError, match='assert:'):
            pp.parse(path)

    def test_empty_assert_expression_raises(self, tmp_path):
        from grepxcel.pattern_parser import PatternError, PatternParser
        content = [
            'var:,total,string,.*',
            'assert:,',
            'START:',
            'cell:next,total',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        pp = PatternParser()
        with pytest.raises(PatternError, match='assert:.*no expression'):
            pp.parse(path)

    def test_multiple_assert_rows(self, tmp_path):
        content = [
            'var:,a,string,.*',
            'var:,b,string,.*',
            'var:,c,string,.*',
            'assert:,a > 0',
            'assert:,b > 0',
            'assert:,c == a + b,totals check',
            'START:',
            'cell:next,a',
            'cell:next,b',
            'cell:next,c',
            'END:',
        ]
        path = self._make_csv(tmp_path, content)
        from grepxcel.pattern_parser import PatternParser
        pp = PatternParser()
        pp.parse(path)
        assert len(pp.assert_rules) == 3
        assert pp.assert_rules[2].message == 'totals check'
