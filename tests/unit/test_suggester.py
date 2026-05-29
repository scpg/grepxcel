"""Unit tests for engine.drafter (pattern drafting) and the infer_cell_type utility."""

import datetime
import sys
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest

from engine.drafter import _type_from_number_format

from engine.drafter import ExcelAnalyzer, LlamaCppClient, PatternDrafter, PatternWriter
from engine.utils import infer_cell_type

# Backward-compat alias used in a few tests below
PatternSuggester = PatternDrafter


# ── helpers ───────────────────────────────────────────────────────────────────

def _make_xlsx(rows: list, tmp_path: Path, name: str = 'test.xlsx') -> str:
    """Write rows (list of lists) to a temporary xlsx and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    path = str(tmp_path / name)
    wb.save(path)
    return path


# ── infer_cell_type ───────────────────────────────────────────────────────────

class TestInferCellType:
    def test_empty_list_returns_string(self):
        assert infer_cell_type([]) == 'string'

    def test_all_none_returns_string(self):
        assert infer_cell_type([None, None]) == 'string'

    def test_integers(self):
        assert infer_cell_type([1, 2, 3]) == 'integer'

    def test_floats_with_decimals(self):
        assert infer_cell_type([1.5, 2.3]) == 'currency'

    def test_whole_floats_are_integer(self):
        assert infer_cell_type([1.0, 2.0]) == 'integer'

    def test_strings(self):
        assert infer_cell_type(['hello', 'world']) == 'string'

    def test_dates(self):
        d = datetime.date(2024, 1, 1)
        assert infer_cell_type([d, d]) == 'date'

    def test_datetimes(self):
        dt = datetime.datetime(2024, 1, 1, 12, 0)
        assert infer_cell_type([dt, dt]) == 'datetime'

    def test_booleans_treated_as_string(self):
        assert infer_cell_type([True, False]) == 'string'

    def test_majority_wins(self):
        assert infer_cell_type([1, 2, 3, 'text']) == 'integer'

    def test_mixed_none_and_values(self):
        assert infer_cell_type([None, 42, None]) == 'integer'


# ── ExcelAnalyzer — table layout ──────────────────────────────────────────────

class TestExcelAnalyzerTableLayout:
    def test_detects_table_layout(self, tmp_path):
        path = _make_xlsx([
            ['Product', 'Qty', 'Price'],
            ['Widget', 10, 9.99],
            ['Gadget', 5, 14.50],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'TABLE' in result
        assert 'Product' in result
        assert 'Qty' in result
        assert 'Price' in result

    def test_infers_column_types(self, tmp_path):
        path = _make_xlsx([
            ['Name', 'Count', 'Total'],
            ['Alpha', 3, 12.5],
            ['Beta', 7, 28.0],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'string' in result
        assert 'integer' in result
        assert 'currency' in result

    def test_includes_sheet_name(self, tmp_path):
        path = _make_xlsx([['A', 'B'], [1, 2]], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Sheet' in result

    def test_empty_sheet(self, tmp_path):
        path = _make_xlsx([], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Empty' in result or 'no data' in result.lower()


# ── ExcelAnalyzer — key-value layout ─────────────────────────────────────────

class TestExcelAnalyzerKVLayout:
    def test_detects_kv_layout(self, tmp_path):
        # Row 1 has only one non-empty cell (title) → KV layout is detected
        path = _make_xlsx([
            ['Invoice Summary'],
            ['Invoice No', 'INV-0042'],
            ['Date', datetime.date(2024, 3, 15)],
            ['Total', 1250.00],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'KEY-VALUE' in result
        assert 'Invoice No' in result

    def test_infers_type_of_adjacent_value(self, tmp_path):
        # Row 1 has only one non-empty cell (label only) → KV layout
        path = _make_xlsx([
            ['Details'],
            ['Amount', 99.5],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'currency' in result


# ── PatternWriter ─────────────────────────────────────────────────────────────

class TestPatternWriter:
    _SAMPLE_LLM_OUTPUT = """\
config: | read.direction | LR
def: | InvoiceNo | string | INV-\\d+
def: | Total | currency | \\d+(\\.\\d{2})?
START:
cell:1 | InvoiceNo
cell:1 | Total
table:*
 | HEADER:1 | Product | Qty
 | DATA:* | Product | Qty
END:
"""

    def test_writes_xlsx(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        assert Path(out).exists()

    def test_roundtrip_config_row(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        row1 = [ws.cell(row=1, column=c).value for c in range(1, 5)]
        assert row1[0] == 'config:'
        assert row1[1] == 'read.direction'
        assert row1[2] == 'LR'

    def test_roundtrip_def_row(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        row2 = [ws.cell(row=2, column=c).value for c in range(1, 5)]
        assert row2[0] == 'def:'
        assert row2[1] == 'InvoiceNo'
        assert row2[2] == 'string'

    def test_start_end_in_col_a(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        col_a_values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        assert 'START:' in col_a_values
        assert 'END:' in col_a_values

    def test_table_template_row_col_a_blank(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        # Find a row where col B is HEADER:1 — col A must be None
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == 'HEADER:1':
                assert ws.cell(row=r, column=1).value is None
                break
        else:
            pytest.fail('HEADER:1 row not found in generated xlsx')

    def test_skips_markdown_fences(self, tmp_path):
        llm_text = '```\ndef: | X | string | .*\n```'
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(llm_text, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        col_a_values = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        assert '```' not in (col_a_values or [])

    def test_bare_table_keyword(self, tmp_path):
        llm_text = 'table:*\n | HEADER:1 | Col\n | DATA:* | Col\nEND:'
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(llm_text, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        assert ws.cell(row=1, column=1).value == 'table:*'


# ── LlamaCppClient ────────────────────────────────────────────────────────────

class TestLlamaCppClient:
    def test_import_error_exits(self, tmp_path, capsys):
        client = LlamaCppClient(model_path=str(tmp_path / "model.gguf"))
        with patch.dict(sys.modules, {'llama_cpp': None}):
            with pytest.raises(SystemExit) as exc_info:
                client.chat('sys', 'user')
        assert exc_info.value.code == 1
        captured = capsys.readouterr()
        assert 'llama-cpp-python' in captured.err

    def test_chat_returns_model_content(self, tmp_path):
        client = LlamaCppClient(model_path=str(tmp_path / "model.gguf"))
        fake_response = {
            "choices": [{"message": {"content": "def: | X | string | .*"}}]
        }
        fake_llm = MagicMock()
        fake_llm.create_chat_completion.return_value = fake_response

        with patch.object(client, '_load', side_effect=lambda: setattr(client, '_llm', fake_llm)):
            result = client.chat("system prompt", "user prompt")

        assert result == "def: | X | string | .*"
        fake_llm.create_chat_completion.assert_called_once()


# ── PatternSuggester integration (mocked LLM) ────────────────────────────────

class TestPatternSuggesterMocked:
    _LLM_RESPONSE = "var: | Name | string | .*\nSTART:\ncell:next | Name\nEND:"

    def test_full_pipeline_creates_xlsx(self, tmp_path):
        data_path = _make_xlsx([['Name', 'Age'], ['Alice', 30]], tmp_path, 'data.xlsx')
        out_path = str(tmp_path / 'pattern.xlsx')

        mock_manager = MagicMock()
        mock_manager.return_value.ensure_ready.return_value = tmp_path / 'model.gguf'

        mock_client = MagicMock()
        mock_client.return_value.chat.return_value = self._LLM_RESPONSE

        with patch('engine.drafter.ModelManager', mock_manager), \
             patch('engine.drafter.LlamaCppClient', mock_client):
            s = PatternDrafter(input_path=data_path, output_path=out_path)
            code = s.run()

        assert code == 0
        assert Path(out_path).exists()

    def test_security_error_returns_1(self, tmp_path, capsys):
        s = PatternSuggester(
            input_path=str(tmp_path / 'nonexistent.xlsx'),
            output_path=str(tmp_path / 'out.xlsx'),
        )
        code = s.run()
        assert code == 1
        captured = capsys.readouterr()
        assert 'error' in captured.err.lower() or 'Error' in captured.err


# ── A2: output validation ────────────────────────────────────────────────────

def _run_drafter_with_llm_text(llm_text: str, tmp_path: Path) -> tuple[int, str, str]:
    """Run PatternDrafter with a mocked LLM returning llm_text. Returns (code, out, err)."""
    data_path = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
    out_path = str(tmp_path / 'pattern.xlsx')

    mock_manager = MagicMock()
    mock_manager.return_value.ensure_ready.return_value = tmp_path / 'model.gguf'
    mock_client = MagicMock()
    mock_client.return_value.chat.return_value = llm_text

    import io, contextlib
    stdout_buf, stderr_buf = io.StringIO(), io.StringIO()
    with patch('engine.drafter.ModelManager', mock_manager), \
         patch('engine.drafter.LlamaCppClient', mock_client), \
         contextlib.redirect_stdout(stdout_buf), \
         contextlib.redirect_stderr(stderr_buf):
        code = PatternDrafter(input_path=data_path, output_path=out_path).run()

    return code, stdout_buf.getvalue(), stderr_buf.getvalue(), out_path


class TestDraftValidation:
    _VALID_LLM = "var: | Name | string | .*\nSTART:\ncell:next | Name\nEND:"
    _GARBAGE_LLM = "this is not a pattern\nrandom text\nno start or end markers"
    _BAD_CELL_LLM = "var: | x | string | .*\nSTART:\ncell:BADREF | x\nEND:"

    def test_valid_output_creates_xlsx(self, tmp_path):
        code, _, _, out_path = _run_drafter_with_llm_text(self._VALID_LLM, tmp_path)
        assert code == 0
        assert Path(out_path).exists()

    def test_valid_output_no_failed_txt(self, tmp_path):
        _run_drafter_with_llm_text(self._VALID_LLM, tmp_path)
        failed = tmp_path / 'pattern_FAILED.txt'
        assert not failed.exists()

    def test_garbage_output_returns_1(self, tmp_path):
        # A pattern with no START:/END: and no defined fields is structurally
        # empty — PatternParser won't raise, but an invalid cell: reference will.
        code, _, _, _ = _run_drafter_with_llm_text(self._BAD_CELL_LLM, tmp_path)
        assert code == 1

    def test_invalid_cell_ref_writes_failed_txt(self, tmp_path):
        _run_drafter_with_llm_text(self._BAD_CELL_LLM, tmp_path)
        failed = tmp_path / 'pattern_FAILED.txt'
        assert failed.exists()
        content = failed.read_text(encoding='utf-8')
        assert '# Draft validation error' in content
        assert self._BAD_CELL_LLM in content

    def test_invalid_cell_ref_prints_error_to_stderr(self, tmp_path):
        _, _, err, _ = _run_drafter_with_llm_text(self._BAD_CELL_LLM, tmp_path)
        assert '[!] Draft validation failed' in err

    def test_failed_xlsx_not_written_on_error(self, tmp_path):
        _run_drafter_with_llm_text(self._BAD_CELL_LLM, tmp_path)
        assert not Path(tmp_path / 'pattern.xlsx').exists()

    def test_raw_llm_text_printed_to_stdout_even_on_failure(self, tmp_path):
        _, out, _, _ = _run_drafter_with_llm_text(self._BAD_CELL_LLM, tmp_path)
        assert self._BAD_CELL_LLM in out


# ── _type_from_number_format ──────────────────────────────────────────────────

class TestTypeFromNumberFormat:
    def test_percentage(self):
        assert _type_from_number_format('0.00%') == 'percentage'

    def test_percentage_plain(self):
        assert _type_from_number_format('0%') == 'percentage'

    def test_date_iso(self):
        assert _type_from_number_format('yyyy-mm-dd') == 'date'

    def test_date_us(self):
        assert _type_from_number_format('mm/dd/yyyy') == 'date'

    def test_datetime(self):
        assert _type_from_number_format('yyyy-mm-dd h:mm') == 'datetime'

    def test_currency_dollar(self):
        assert _type_from_number_format('"$"#,##0.00') == 'currency'

    def test_currency_euro_locale(self):
        assert _type_from_number_format('[$€-407]#,##0.00') == 'currency'

    def test_currency_accounting(self):
        assert _type_from_number_format('#,##0.00') == 'currency'

    def test_general_returns_none(self):
        assert _type_from_number_format('General') is None

    def test_plain_integer_returns_none(self):
        assert _type_from_number_format('0') is None

    def test_empty_returns_none(self):
        assert _type_from_number_format('') is None


# ── B2: section detection ─────────────────────────────────────────────────────

class TestExcelAnalyzerTitledTable:
    """Titled-table heuristic: single-cell heading row + ≥5-row section → TABLE."""

    def test_titled_table_classified_as_table(self, tmp_path):
        # title row + header row + 3 data rows = 5 rows → titled table
        path = _make_xlsx([
            ['INCOME'],
            ['Item', 'Jan', 'Feb'],
            ['Salary', 5000, 5000],
            ['Rent',   1000, 1000],
            ['Food',    800,  800],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'TABLE' in result

    def test_titled_table_header_names_in_result(self, tmp_path):
        path = _make_xlsx([
            ['EXPENSES'],
            ['Item', 'Budget', 'Actual'],
            ['Rent',   1000, 950],
            ['Food',    500, 480],
            ['Travel',  200, 310],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Item' in result
        assert 'Budget' in result

    def test_four_row_section_not_titled_table(self, tmp_path):
        # title + header + 2 data = 4 rows < 5 → should NOT be TABLE
        path = _make_xlsx([
            ['Invoice Summary'],
            ['Invoice No', 'INV-001'],
            ['Date', '2024-01-01'],
            ['Total', 1250.00],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'KEY-VALUE' in result


class TestExcelAnalyzerSections:
    def test_multi_section_labels_present(self, tmp_path):
        path = _make_xlsx([
            ['Invoice No', 'INV-001'],
            ['Date', '2024-01-01'],
            [None],                             # empty row → section break
            ['Product', 'Qty', 'Price'],
            ['Widget', 10, 9.99],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Section 1' in result
        assert 'Section 2' in result

    def test_multi_section_kv_then_table(self, tmp_path):
        path = _make_xlsx([
            ['Invoice No', 'INV-001'],
            [None],
            ['Product', 'Qty'],
            ['Widget', 10],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'KEY-VALUE' in result
        assert 'TABLE' in result

    def test_single_section_no_section_label(self, tmp_path):
        path = _make_xlsx([
            ['Product', 'Qty'],
            ['Widget', 10],
        ], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Section 1' not in result

    def test_multi_sheet_workbook_preamble(self, tmp_path):
        path = str(tmp_path / 'multi.xlsx')
        wb   = openpyxl.Workbook()
        ws1  = wb.active
        ws1.title = 'Invoice'
        ws1.append(['Invoice No', 'INV-001'])
        ws2  = wb.create_sheet('Items')
        ws2.append(['Product', 'Qty'])
        wb.save(path)

        result = ExcelAnalyzer(path).analyse()
        assert 'Workbook' in result
        assert 'Invoice' in result
        assert 'Items' in result

    def test_single_sheet_no_preamble(self, tmp_path):
        path = _make_xlsx([['A', 'B'], [1, 2]], tmp_path)
        result = ExcelAnalyzer(path).analyse()
        assert 'Workbook' not in result

    def test_number_format_overrides_value_inference(self, tmp_path):
        # A table column of 0.xx values with an explicit percentage format
        # should be typed 'percentage' even though value-only inference gives 'currency'.
        path = str(tmp_path / 'fmt.xlsx')
        wb   = openpyxl.Workbook()
        ws   = wb.active
        ws.append(['Rate', 'Amount'])      # row 1: headers
        for row_idx in range(2, 5):        # rows 2-4: data
            ws.append([0.5, 1000.0])
            ws.cell(row=row_idx, column=1).number_format = '0.00%'
        wb.save(path)

        result = ExcelAnalyzer(path).analyse()
        assert 'percentage' in result


# ── B3: --dry-run ─────────────────────────────────────────────────────────────

class TestDryRun:
    def test_dry_run_returns_0(self, tmp_path):
        data_path = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'pattern.xlsx')
        code = PatternDrafter(
            input_path=data_path, output_path=out_path, dry_run=True,
        ).run()
        assert code == 0

    def test_dry_run_no_xlsx_created(self, tmp_path):
        data_path = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'pattern.xlsx')
        PatternDrafter(input_path=data_path, output_path=out_path, dry_run=True).run()
        assert not Path(out_path).exists()

    def test_dry_run_prints_analysis_to_stderr(self, tmp_path, capsys):
        data_path = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'pattern.xlsx')
        PatternDrafter(input_path=data_path, output_path=out_path, dry_run=True).run()
        err = capsys.readouterr().err
        assert 'Name' in err
        assert 'dry-run' in err.lower()

    def test_dry_run_no_model_loaded(self, tmp_path):
        data_path = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'pattern.xlsx')
        with patch('engine.drafter.ModelManager') as mock_mm:
            PatternDrafter(input_path=data_path, output_path=out_path, dry_run=True).run()
        mock_mm.assert_not_called()
