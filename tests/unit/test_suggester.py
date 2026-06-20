"""Unit tests for grepxcel.drafter (pattern drafting) and the infer_cell_type utility."""

import datetime
import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import openpyxl
import pytest

from grepxcel.drafter import _type_from_number_format

from grepxcel.drafter import ClaudeBackend, ExcelAnalyzer, GeminiBackend, LlamaCppClient, OpenAICompatBackend, PatternDrafter, PatternWriter
from grepxcel.utils import infer_cell_type

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

    def test_stamps_pattern_version_first(self, tmp_path):
        """Generated patterns get an explicit pattern.version stamped as row 1."""
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        ws = openpyxl.load_workbook(out).active
        row1 = [ws.cell(row=1, column=c).value for c in range(1, 4)]
        assert row1 == ['config:', 'pattern.version', '1']

    def test_roundtrip_config_row(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        wb = openpyxl.load_workbook(out)
        ws = wb.active
        # read.direction is now row 2 (row 1 is the stamped pattern.version)
        row2 = [ws.cell(row=2, column=c).value for c in range(1, 5)]
        assert row2[0] == 'config:'
        assert row2[1] == 'read.direction'
        assert row2[2] == 'LR'

    def test_no_duplicate_version_when_model_emits_one(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        llm = 'config: | pattern.version | 1\n' + self._SAMPLE_LLM_OUTPUT
        PatternWriter().write(llm, out)
        ws = openpyxl.load_workbook(out).active
        versions = [ws.cell(row=r, column=2).value for r in range(1, ws.max_row + 1)]
        assert versions.count('pattern.version') == 1

    def test_roundtrip_def_row(self, tmp_path):
        out = str(tmp_path / 'out.xlsx')
        PatternWriter().write(self._SAMPLE_LLM_OUTPUT, out)
        ws = openpyxl.load_workbook(out).active
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == 'def:' and \
               ws.cell(row=r, column=2).value == 'InvoiceNo':
                assert ws.cell(row=r, column=3).value == 'string'
                break
        else:
            pytest.fail('def: InvoiceNo row not found')

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
        col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
        assert 'table:*' in col_a


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
        data_path    = _make_xlsx([['Name', 'Age'], ['Alice', 30]], tmp_path, 'data.xlsx')
        out_path     = str(tmp_path / 'pattern.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = self._LLM_RESPONSE

        code = PatternDrafter(
            input_path=data_path, output_path=out_path, backend=mock_backend,
        ).run()

        assert code == 0
        assert Path(out_path).exists()

    def test_security_error_returns_1(self, tmp_path, capsys):
        # Inject a backend so the local-deps fail-fast is bypassed — the security
        # error (missing file) is raised during analysis regardless of backend,
        # which is what this test verifies. (Without a backend, on a machine
        # lacking the local deps, the dep-check would short-circuit first.)
        s = PatternSuggester(
            input_path=str(tmp_path / 'nonexistent.xlsx'),
            output_path=str(tmp_path / 'out.xlsx'),
            backend=MagicMock(),
        )
        code = s.run()
        assert code == 1
        captured = capsys.readouterr()
        assert 'error' in captured.err.lower() or 'Error' in captured.err


# ── A2: output validation ────────────────────────────────────────────────────

def _run_drafter_with_llm_text(llm_text: str, tmp_path: Path) -> tuple[int, str, str]:
    """Run PatternDrafter with a backend stub returning llm_text. Returns (code, out, err, path)."""
    import contextlib, io
    data_path    = _make_xlsx([['Name'], ['Alice']], tmp_path, 'data.xlsx')
    out_path     = str(tmp_path / 'pattern.xlsx')
    mock_backend = MagicMock()
    mock_backend.chat.return_value = llm_text

    stdout_buf, stderr_buf = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(stdout_buf), contextlib.redirect_stderr(stderr_buf):
        code = PatternDrafter(
            input_path=data_path, output_path=out_path, backend=mock_backend,
        ).run()

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

    def test_duration_elapsed_hours(self):
        assert _type_from_number_format('[h]:mm') == 'duration'

    def test_duration_elapsed_hours_seconds(self):
        assert _type_from_number_format('[hh]:mm:ss') == 'duration'

    def test_duration_elapsed_minutes(self):
        assert _type_from_number_format('[mm]:ss') == 'duration'

    def test_time_clock(self):
        assert _type_from_number_format('h:mm') == 'time'

    def test_time_clock_am_pm_locale(self):
        # Real Excel clock format — contains [$...] locale tag, must NOT be currency.
        assert _type_from_number_format('[$-409]h:mm\\ AM/PM;@') == 'time'

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
        with patch('grepxcel.drafter.ModelManager') as mock_mm:
            PatternDrafter(input_path=data_path, output_path=out_path, dry_run=True).run()
        mock_mm.assert_not_called()


# ── C1: LLMBackend protocol ───────────────────────────────────────────────────

from grepxcel.drafter import LLMBackend  # noqa: E402


class TestLLMBackendProtocol:
    def test_llamacppclient_satisfies_protocol(self):
        """LlamaCppClient must implement LLMBackend structurally."""
        from grepxcel.drafter import LlamaCppClient
        assert isinstance(LlamaCppClient('dummy.gguf'), LLMBackend)

    def test_plain_object_with_chat_satisfies_protocol(self):
        class MyBackend:
            def chat(self, system: str, user: str) -> str:
                return 'ok'
        assert isinstance(MyBackend(), LLMBackend)

    def test_object_without_chat_does_not_satisfy_protocol(self):
        class NotABackend:
            pass
        assert not isinstance(NotABackend(), LLMBackend)

    def test_backend_injection_bypasses_model_manager(self, tmp_path):
        """Providing backend= must skip ModelManager entirely."""
        data_path    = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path     = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        with patch('grepxcel.drafter.ModelManager') as mock_mm:
            PatternDrafter(
                input_path=data_path, output_path=out_path, backend=mock_backend,
            ).run()
        mock_mm.assert_not_called()

    def test_backend_chat_called_with_system_and_user_prompts(self, tmp_path):
        data_path    = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path     = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        PatternDrafter(
            input_path=data_path, output_path=out_path, backend=mock_backend,
        ).run()
        mock_backend.chat.assert_called_once()
        system_arg, user_arg = mock_backend.chat.call_args.args
        assert 'grepxcel' in system_arg.lower()
        assert 'pattern' in system_arg.lower()
        assert 'analyse' in user_arg.lower() or 'analysis' in user_arg.lower()


# ── System-prompt content coverage ───────────────────────────────────────────

class TestSystemPromptCoverage:
    """The system prompt must teach the LLM the current instruction set and types,
    so drafted patterns use features that actually exist in the parser."""

    def test_documents_dir_instruction(self):
        from grepxcel.drafter import _SYSTEM_PROMPT
        assert 'dir:LR' in _SYSTEM_PROMPT
        assert 'dir:TD' in _SYSTEM_PROMPT

    def test_documents_seek_instruction(self):
        from grepxcel.drafter import _SYSTEM_PROMPT
        assert 'seek:' in _SYSTEM_PROMPT

    def test_documents_time_and_duration_types(self):
        from grepxcel.drafter import _SYSTEM_PROMPT
        assert 'time' in _SYSTEM_PROMPT
        assert 'duration' in _SYSTEM_PROMPT

    def test_does_not_teach_legacy_syntax(self):
        from grepxcel.drafter import _SYSTEM_PROMPT
        assert 'def:' not in _SYSTEM_PROMPT
        assert 'cell:1 ' not in _SYSTEM_PROMPT


# ── C2: ClaudeBackend ─────────────────────────────────────────────────────────

class TestClaudeBackend:
    def test_satisfies_llm_backend_protocol(self):
        from grepxcel.drafter import LLMBackend
        assert isinstance(ClaudeBackend(), LLMBackend)

    def test_default_model(self):
        assert ClaudeBackend()._model == 'claude-haiku-4-5-20251001'

    def test_custom_model(self):
        assert ClaudeBackend(model='claude-opus-4-8')._model == 'claude-opus-4-8'

    def test_chat_calls_anthropic_with_correct_args(self):
        mock_anthropic = MagicMock()
        mock_msg = MagicMock()
        mock_msg.content = [MagicMock(text='pattern output')]
        mock_anthropic.Anthropic.return_value.messages.create.return_value = mock_msg

        with patch.dict('sys.modules', {'anthropic': mock_anthropic}):
            result = ClaudeBackend().chat('sys prompt', 'user prompt')

        assert result == 'pattern output'
        create_call = mock_anthropic.Anthropic.return_value.messages.create
        create_call.assert_called_once()
        kwargs = create_call.call_args.kwargs
        assert kwargs['system'] == 'sys prompt'
        assert kwargs['messages'] == [{'role': 'user', 'content': 'user prompt'}]
        assert kwargs['model'] == 'claude-haiku-4-5-20251001'

    def test_missing_anthropic_exits(self, capsys):
        with patch.dict('sys.modules', {'anthropic': None}):
            import importlib
            import builtins
            real_import = builtins.__import__
            def mock_import(name, *args, **kwargs):
                if name == 'anthropic':
                    raise ImportError('No module named anthropic')
                return real_import(name, *args, **kwargs)
            with patch('builtins.__import__', side_effect=mock_import):
                with pytest.raises(SystemExit) as exc_info:
                    ClaudeBackend().chat('sys', 'user')
        assert exc_info.value.code == 1


@pytest.mark.skip(reason="Gemini backend disabled — planned for a future release")
class TestGeminiBackend:
    def test_satisfies_llm_backend_protocol(self):
        from grepxcel.drafter import LLMBackend
        assert isinstance(GeminiBackend(), LLMBackend)

    def test_default_model(self):
        assert GeminiBackend()._model == 'gemini-2.0-flash'

    def test_custom_model(self):
        assert GeminiBackend(model='gemini-2.5-pro')._model == 'gemini-2.5-pro'

    def test_missing_google_genai_exits(self):
        import builtins
        real_import = builtins.__import__
        def mock_import(name, *args, **kwargs):
            if name == 'google' or name.startswith('google.'):
                raise ImportError('No module named google')
            return real_import(name, *args, **kwargs)
        with patch('builtins.__import__', side_effect=mock_import):
            with pytest.raises(SystemExit) as exc_info:
                GeminiBackend().chat('sys', 'user')
        assert exc_info.value.code == 1


# ── OpenAI-compatible server backend ─────────────────────────────────────────

class TestOpenAICompatBackend:
    def test_satisfies_llm_backend_protocol(self):
        from grepxcel.drafter import LLMBackend
        assert isinstance(OpenAICompatBackend(), LLMBackend)

    def test_default_base_url(self):
        assert OpenAICompatBackend()._base_url == 'http://localhost:1234/v1'

    def test_custom_base_url(self):
        b = OpenAICompatBackend(base_url='http://myhost:8080/v1')
        assert b._base_url == 'http://myhost:8080/v1'

    def test_chat_refuses_non_http_base_url(self):
        """A non-http(s) base_url (e.g. file://) must be refused before any
        client/network call — guards against SSRF to internal endpoints."""
        b = OpenAICompatBackend(base_url='file:///etc/passwd')
        with pytest.raises(ValueError, match='http'):
            b.chat('sys', 'user')

    def test_chat_returns_model_content(self):
        mock_openai = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [MagicMock(message=MagicMock(content='pattern output'))]
        mock_completion.usage = MagicMock(prompt_tokens=100, completion_tokens=50)
        mock_openai.OpenAI.return_value.chat.completions.create.return_value = mock_completion
        mock_openai.OpenAI.return_value.models.list.return_value = MagicMock(
            data=[MagicMock(id='test-model')]
        )

        with patch.dict('sys.modules', {'openai': mock_openai}):
            result = OpenAICompatBackend().chat('sys prompt', 'user prompt')

        assert result == 'pattern output'

    def test_chat_uses_specified_model(self):
        mock_openai = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [MagicMock(message=MagicMock(content='ok'))]
        mock_completion.usage = MagicMock(prompt_tokens=10, completion_tokens=5)
        mock_openai.OpenAI.return_value.chat.completions.create.return_value = mock_completion

        with patch.dict('sys.modules', {'openai': mock_openai}):
            OpenAICompatBackend(model='my-model').chat('sys', 'user')

        create_call = mock_openai.OpenAI.return_value.chat.completions.create
        assert create_call.call_args.kwargs['model'] == 'my-model'

    def test_auto_discovers_model_when_none_specified(self):
        mock_openai = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [MagicMock(message=MagicMock(content='ok'))]
        mock_completion.usage = MagicMock(prompt_tokens=10, completion_tokens=5)
        mock_openai.OpenAI.return_value.chat.completions.create.return_value = mock_completion
        mock_openai.OpenAI.return_value.models.list.return_value = MagicMock(
            data=[MagicMock(id='discovered-model')]
        )

        with patch.dict('sys.modules', {'openai': mock_openai}):
            OpenAICompatBackend().chat('sys', 'user')

        create_call = mock_openai.OpenAI.return_value.chat.completions.create
        assert create_call.call_args.kwargs['model'] == 'discovered-model'

    def test_no_models_loaded_raises_runtime_error(self):
        mock_openai = MagicMock()
        mock_openai.OpenAI.return_value.models.list.return_value = MagicMock(data=[])

        with patch.dict('sys.modules', {'openai': mock_openai}):
            with pytest.raises(RuntimeError, match='No models loaded'):
                OpenAICompatBackend().chat('sys', 'user')

    def test_last_cost_reports_zero_dollars(self):
        mock_openai = MagicMock()
        mock_completion = MagicMock()
        mock_completion.choices = [MagicMock(message=MagicMock(content='ok'))]
        mock_completion.usage = MagicMock(prompt_tokens=100, completion_tokens=50)
        mock_openai.OpenAI.return_value.chat.completions.create.return_value = mock_completion

        with patch.dict('sys.modules', {'openai': mock_openai}):
            b = OpenAICompatBackend(model='test-model')
            b.chat('sys', 'user')

        cost = b.last_cost()
        assert cost is not None
        assert cost.input_tokens == 100
        assert cost.output_tokens == 50
        assert cost.total_cost_usd == 0.0


# ── C3: --backend CLI flag ────────────────────────────────────────────────────

class TestBackendCLIFlag:
    def test_default_backend_is_local(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', 'data.xlsx'])
        assert args.backend == 'local'

    def test_backend_claude_accepted(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', '--backend', 'claude', 'data.xlsx'])
        assert args.backend == 'claude'

    @pytest.mark.skip(reason="Gemini backend disabled — planned for a future release")
    def test_backend_gemini_accepted(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', '--backend', 'gemini', 'data.xlsx'])
        assert args.backend == 'gemini'

    def test_backend_server_accepted(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', '--backend', 'server', 'data.xlsx'])
        assert args.backend == 'server'

    def test_server_url_flag_parsed(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args([
            'draft', '--backend', 'server',
            '--server-url', 'http://myhost:8080/v1',
            'data.xlsx',
        ])
        assert args.server_url == 'http://myhost:8080/v1'

    def test_server_url_defaults_to_lmstudio(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', 'data.xlsx'])
        assert args.server_url == 'http://localhost:1234/v1'

    def test_server_model_flag_parsed(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args([
            'draft', '--backend', 'server',
            '--server-model', 'qwen2.5-coder-7b',
            'data.xlsx',
        ])
        assert args.server_model == 'qwen2.5-coder-7b'

    def test_server_model_defaults_to_none(self):
        from grepxcel.cli import _build_parser
        args = _build_parser().parse_args(['draft', 'data.xlsx'])
        assert args.server_model is None

    def test_server_backend_wires_openai_compat(self, tmp_path, capsys):
        data_path = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        with patch('grepxcel.drafter.OpenAICompatBackend',
                   return_value=mock_backend) as mock_cls:
            from grepxcel.cli import _build_parser, _run_draft
            args = _build_parser().parse_args([
                'draft', '--backend', 'server', data_path, '-o', out_path,
            ])
            _run_draft(args)
        mock_cls.assert_called_once()
        mock_backend.chat.assert_called_once()

    def test_backend_invalid_rejected(self):
        from grepxcel.cli import _build_parser
        with pytest.raises(SystemExit):
            _build_parser().parse_args(['draft', '--backend', 'openai', 'data.xlsx'])

    @pytest.mark.skip(reason="Gemini backend disabled — planned for a future release")
    def test_gemini_backend_emits_privacy_warning_and_is_used(self, tmp_path, capsys):
        """Re-enable when the Gemini backend ships (flip _GEMINI_ENABLED)."""
        data_path    = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path     = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        with patch('grepxcel.drafter.GeminiBackend', return_value=mock_backend) as mock_cls:
            from grepxcel.cli import _build_parser, _run_draft
            args = _build_parser().parse_args([
                'draft', '--backend', 'gemini', data_path, '-o', out_path,
            ])
            _run_draft(args)
        mock_cls.assert_called_once()
        mock_backend.chat.assert_called_once()
        assert 'Google' in capsys.readouterr().err

    def test_gemini_backend_is_disabled_with_future_release_message(self, tmp_path, capsys):
        """ACTIVE guard: --backend gemini is recognised but disabled — it must
        print a 'future release' notice, exit non-zero, and run NO inference."""
        data_path = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'out.xlsx')
        from grepxcel.cli import _build_parser, _run_draft
        args = _build_parser().parse_args([
            'draft', '--backend', 'gemini', data_path, '-o', out_path,
        ])
        rc  = _run_draft(args)
        err = capsys.readouterr().err
        assert rc == 1
        assert 'future release' in err.lower()
        assert not os.path.exists(out_path)  # no draft written

    def test_claude_backend_emits_privacy_warning(self, tmp_path, capsys):
        """--backend claude must print the privacy notice to stderr."""
        data_path    = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path     = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        # ClaudeBackend is imported lazily inside _run_draft — patch at the source
        with patch('grepxcel.drafter.ClaudeBackend', return_value=mock_backend):
            from grepxcel.cli import _build_parser, _run_draft
            args = _build_parser().parse_args([
                'draft', '--backend', 'claude', data_path, '-o', out_path,
            ])
            _run_draft(args)
        captured = capsys.readouterr()
        assert 'Anthropic' in captured.err

    def test_local_backend_uses_no_claude_backend(self, tmp_path):
        """--backend local must never instantiate ClaudeBackend."""
        data_path = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out_path  = str(tmp_path / 'out.xlsx')
        with patch('grepxcel.drafter.ClaudeBackend') as mock_cls, \
             patch('grepxcel.drafter.ModelManager') as mock_mm:
            mock_mm.return_value.ensure_ready.side_effect = RuntimeError('no model')
            from grepxcel.cli import _build_parser, _run_draft
            args = _build_parser().parse_args(['draft', 'data.xlsx', '-o', out_path])
            try:
                _run_draft(args)
            except RuntimeError:
                pass
        mock_cls.assert_not_called()


# ── Local-backend dependency check (fail fast + clear message) ─────────────────

class TestLocalBackendDepCheck:
    def test_fails_fast_with_guidance_when_local_deps_missing(self, tmp_path, capsys):
        """No cloud backend + deps missing → fail BEFORE analysis, with guidance."""
        data = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out  = str(tmp_path / 'out.xlsx')
        with patch('grepxcel.drafter._missing_local_deps',
                   return_value=['llama-cpp-python', 'huggingface_hub']):
            rc = PatternDrafter(input_path=data, output_path=out).run()
        err = capsys.readouterr().err
        assert rc == 1
        assert "grepxcel[suggest]" in err            # tells the user what to install
        assert "Analysing Excel structure" not in err  # failed fast, no analysis
        assert not os.path.exists(out)                  # nothing written

    def test_dry_run_does_not_require_local_deps(self, tmp_path):
        """--dry-run never touches the model, so missing local deps must be fine."""
        data = _make_xlsx([['Product', 'Qty'], ['Widget', 3]], tmp_path, 'data.xlsx')
        out  = str(tmp_path / 'out.xlsx')
        with patch('grepxcel.drafter._missing_local_deps',
                   return_value=['llama-cpp-python', 'huggingface_hub']):
            rc = PatternDrafter(input_path=data, output_path=out, dry_run=True).run()
        assert rc == 0

    def test_cloud_backend_does_not_require_local_deps(self, tmp_path):
        """A provided backend bypasses the local-deps check entirely."""
        data = _make_xlsx([['X'], [1]], tmp_path, 'data.xlsx')
        out  = str(tmp_path / 'out.xlsx')
        mock_backend = MagicMock()
        mock_backend.chat.return_value = (
            "var: | x | integer | .*\nSTART:\ncell:next | x\nEND:"
        )
        with patch('grepxcel.drafter._missing_local_deps',
                   return_value=['llama-cpp-python', 'huggingface_hub']):
            rc = PatternDrafter(
                input_path=data, output_path=out, backend=mock_backend,
            ).run()
        assert rc == 0
        mock_backend.chat.assert_called_once()
