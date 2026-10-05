"""Unit tests for grepxcel.profile — the `profile` subcommand's logic."""
import io
import json
import os
import zipfile

import openpyxl
import pytest

from grepxcel.profile import (
    ProfileCell, SheetProfile, filter_cells, group_by_type, has_error,
    profile_sheet, profile_workbook, render_full, render_summary,
    run_profile, to_json_records, write_colored_xlsx, write_json,
)
from grepxcel.cell_taxonomy import classify_value
from grepxcel.engine import scan_image_cells
from grepxcel.color import MARK_FAIL, MARK_INFO, MARK_WARN
from grepxcel.security import SecurityError

_FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'fixtures')
_FIXTURE_24 = os.path.join(_FIXTURES_DIR, '24_type_tests', '24_type_tests_data.xlsx')
_FIXTURE_01 = os.path.join(_FIXTURES_DIR, '01_simple_invoice', '01_simple_invoice_data.xlsx')
_FIXTURE_19 = os.path.join(_FIXTURES_DIR, '19_blood_pressure_tracker', '19_blood_pressure_tracker_data.xlsx')


def _write_error_workbook(path: str) -> str:
    """A minimal workbook with one genuine Excel error cell (#DIV/0!) and one
    clean cell. Fixture 24's only 'error'-looking cells are actually IMAGE()
    rich-value cells (see TestProfileWorkbookRealFixtures below), so exit-code
    / colored-xlsx tests that need a real error use this instead."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = '#DIV/0!'
    ws['A2'] = 'clean value'
    wb.save(path)
    return ws.title


def _cell(ref, value, semantic='string', flags=None, sheet='Sheet1'):
    profile = classify_value(value)
    profile.semantic_type = semantic
    profile.flags = flags or []
    return ProfileCell(ref=ref, sheet=sheet, profile=profile)


# ── grouping / filtering (pure logic, no file I/O) ──────────────────────────

class TestGroupByType:
    def test_groups_by_storage_and_semantic(self):
        cells = [
            _cell('A1', 'x', 'string'),
            _cell('A2', 'y', 'string'),
            _cell('B1', 1, 'integer'),
        ]
        cells[0].profile.storage_type = 's'
        cells[1].profile.storage_type = 's'
        cells[2].profile.storage_type = 'n'
        groups = group_by_type(cells)
        assert groups[('s', 'string')] == [cells[0], cells[1]]
        assert groups[('n', 'integer')] == [cells[2]]

    def test_empty_input(self):
        assert group_by_type([]) == {}


class TestFilterCells:
    def test_none_mode_returns_everything(self):
        cells = [_cell('A1', 1), _cell('A2', 2, flags=['error'])]
        assert filter_cells(cells, None) == cells

    def test_errors_mode(self):
        ok = _cell('A1', 1)
        err = _cell('A2', '#DIV/0!', flags=['error'])
        assert filter_cells([ok, err], 'errors') == [err]

    def test_issues_mode_includes_errors_and_text_forced_numeric(self):
        ok = _cell('A1', 1)
        err = _cell('A2', '#DIV/0!', flags=['error'])
        tfn = _cell('A3', '00123', flags=['text_forced_numeric'])
        result = filter_cells([ok, err, tfn], 'issues')
        assert result == [err, tfn]

    def test_unknown_mode_raises(self):
        with pytest.raises(ValueError):
            filter_cells([], 'bogus')


class TestHasError:
    def test_true_when_any_error(self):
        cells = [_cell('A1', 1), _cell('A2', '#N/A', flags=['error'])]
        assert has_error(cells)

    def test_false_when_none(self):
        assert not has_error([_cell('A1', 1), _cell('A2', 'x')])

    def test_false_for_empty_list(self):
        assert not has_error([])


# ── rendering ────────────────────────────────────────────────────────────────

class TestRendering:
    def test_summary_shows_counts_and_sample_refs(self):
        cells = [_cell(f'A{i}', i, 'integer') for i in range(1, 6)]
        for c in cells:
            c.profile.storage_type = 'n'
        text = render_summary(cells)
        assert 'n/integer' in text
        assert '5' in text
        assert '+2 more' in text  # sample limit is 3

    def test_summary_empty(self):
        assert 'no cells' in render_summary([]).lower()

    def test_full_lists_every_cell(self):
        cells = [_cell(f'A{i}', i, 'integer') for i in range(1, 6)]
        for c in cells:
            c.profile.storage_type = 'n'
        text = render_full(cells)
        for i in range(1, 6):
            assert f'A{i}' in text

    def test_full_shows_flags(self):
        c = _cell('A1', '#DIV/0!', flags=['error'])
        text = render_full([c])
        assert 'error' in text

    def test_summary_disambiguates_refs_across_sheets(self):
        """Same coordinate on two different sheets must not collapse into
        one ambiguous sample ref (the bug: `A1, B1` gave no clue that A1
        and B1 might belong to different sheets)."""
        cells = [
            _cell('A1', 1, 'integer', sheet='Sheet1'),
            _cell('A1', 2, 'integer', sheet='Sheet2'),
        ]
        text = render_summary(cells)
        assert 'Sheet1!A1' in text
        assert 'Sheet2!A1' in text

    def test_summary_single_sheet_refs_have_no_prefix(self):
        cells = [_cell('A1', 1, 'integer', sheet='Sheet1')]
        text = render_summary(cells)
        assert 'Sheet1!A1' not in text
        assert 'A1' in text

    def test_full_disambiguates_refs_across_sheets(self):
        cells = [
            _cell('A1', 1, 'integer', sheet='Sheet1'),
            _cell('A1', 2, 'integer', sheet='Sheet2'),
        ]
        text = render_full(cells)
        assert 'Sheet1!A1' in text
        assert 'Sheet2!A1' in text

    def test_error_cells_get_fail_mark_not_warn(self):
        """Errors must use the red/FAIL severity, not amber/WARN — matches
        the semantic palette in color.py (MARK_FAIL = error, MARK_WARN =
        lesser issue)."""
        c = _cell('A1', '#DIV/0!', semantic='error', flags=['error'])
        text = render_summary([c], color=True)
        assert MARK_FAIL in text
        assert MARK_WARN not in text

    def test_rich_value_cells_get_info_mark_not_fail(self):
        """A richData IMAGE() cell is informational, not an error or warning."""
        c = _cell('A1', '#VALUE!', semantic='image', flags=['rich_value'])
        text = render_summary([c], color=True)
        assert MARK_INFO in text
        assert MARK_FAIL not in text
        assert MARK_WARN not in text


# ── JSON export ──────────────────────────────────────────────────────────────

class TestJsonExport:
    def test_to_json_records_shape(self):
        sp = SheetProfile(sheet='Sheet1', cells=[_cell('A1', 1, 'integer')])
        records = to_json_records([sp], None)
        assert len(records) == 1
        r = records[0]
        assert r['ref'] == 'A1'
        assert r['sheet'] == 'Sheet1'
        assert 'flags' in r and 'raw_value' in r and 'number_format' in r

    def test_to_json_records_respects_filter(self):
        sp = SheetProfile(sheet='Sheet1', cells=[
            _cell('A1', 1, 'integer'),
            _cell('A2', '#N/A', flags=['error']),
        ])
        records = to_json_records([sp], 'errors')
        assert len(records) == 1
        assert records[0]['ref'] == 'A2'

    def test_write_json_roundtrip(self, tmp_path):
        sp = SheetProfile(sheet='Sheet1', cells=[_cell('A1', 1, 'integer')])
        out = str(tmp_path / 'out.json')
        write_json([sp], None, out)
        with open(out, encoding='utf-8') as fh:
            data = json.load(fh)
        assert data[0]['ref'] == 'A1'


# ── real-file integration: profile_sheet / profile_workbook ────────────────

class TestProfileWorkbookRealFixtures:
    def test_fixture_24_image_cells_not_misreported_as_errors(self):
        """An IMAGE()-formula cell must never be reported as an Excel error.

        Excel caches such a cell's own formula result as the literal string
        '#VALUE!'; the real content lives in the richData chain that
        scan_image_cells() walks, not in the cached scalar. Reading only the
        scalar is what once made profile call them errors.

        The refs are derived from scan_image_cells() rather than hard-coded.
        They used to be listed as M20/E13/H13, which tied the test to one
        revision of the fixture: a round-trip through a spreadsheet editor
        strips richData (no editor outside Excel preserves it), and those three
        cells then hold a bare '#VALUE!' with no chain behind them — genuinely
        errors, so a blanket "this file has no error cells" no longer describes
        the file. Deriving the refs keeps the actual guard at full strength and
        immune to that churn.
        """
        image_by_sheet = scan_image_cells(_FIXTURE_24)
        checked = 0

        for sp in profile_workbook(_FIXTURE_24):
            profiled = {c.ref: c for c in sp.cells}
            # Only image cells that carry a cached scalar reach the profiler at
            # all: one with no cached value is skipped as an empty cell, so it
            # cannot be misreported and is not what this guards. Intersecting
            # keeps the test honest about which cells are actually at risk.
            at_risk = set(image_by_sheet.get(sp.sheet, {})) & set(profiled)

            for ref in sorted(at_risk):
                cell = profiled[ref]
                assert cell.profile.semantic_type == 'image', (
                    f'{sp.sheet}!{ref} is an IMAGE() cell but profile called it '
                    f'{cell.profile.semantic_type!r} — the cached #VALUE! scalar '
                    f'was read instead of the richData chain'
                )
                assert 'rich_value' in cell.profile.flags
                assert 'error' not in cell.profile.flags, (
                    f'{sp.sheet}!{ref} is an IMAGE() cell flagged as an Excel error'
                )
                checked += 1

            # The converse, per sheet: nothing flagged as an error is an image.
            errors = {c.ref for c in sp.cells if 'error' in c.profile.flags}
            misreported = errors & set(image_by_sheet.get(sp.sheet, {}))
            assert not misreported, (
                f'{sp.sheet}: image cells misreported as errors: {sorted(misreported)}'
            )

        assert checked >= 3, (
            f'only {checked} image cell(s) carried a cached value — the fixture '
            f'no longer exercises this guard. A round-trip through a spreadsheet '
            f'editor strips richData; restore the fixture from git.'
        )

    def test_fixture_24_all_sheets_by_default(self):
        sheets = profile_workbook(_FIXTURE_24)
        assert len(sheets) == 2
        assert {sp.sheet for sp in sheets} == {'Sheet1', 'Sheet1_2'}

    def test_fixture_24_sheet_narrowing(self):
        sheets = profile_workbook(_FIXTURE_24, sheet_name='Sheet1')
        assert len(sheets) == 1
        assert sheets[0].sheet == 'Sheet1'

    def test_fixture_24_sheet_narrowing_by_index(self):
        sheets = profile_workbook(_FIXTURE_24, sheet_name='1')
        assert sheets[0].sheet == 'Sheet1_2'

    def test_fixture_24_finds_currency_and_percentage(self):
        sheets = profile_workbook(_FIXTURE_24)
        all_cells = [c for sp in sheets for c in sp.cells]
        semantics = {c.profile.semantic_type for c in all_cells}
        assert 'currency' in semantics
        assert 'percentage' in semantics
        assert 'url' in semantics
        assert 'duration' in semantics

    def test_fixture_24_formula_cells_flagged(self):
        sheets = profile_workbook(_FIXTURE_24)
        all_cells = [c for sp in sheets for c in sp.cells]
        formula_cells = [c for c in all_cells if 'formula' in c.profile.flags]
        assert formula_cells  # fixture 24 has a FUNCTION-header column

    def test_fixture_01_small_and_clean(self):
        sheets = profile_workbook(_FIXTURE_01)
        all_cells = [c for sp in sheets for c in sp.cells]
        assert not has_error(all_cells)
        assert len(all_cells) == 16  # 5 rows x 6 cols, matches lint's reported extent


# ── merged cells ─────────────────────────────────────────────────────────────

class TestMergedCells:
    def test_merged_member_flagged_and_not_double_counted(self, tmp_path):
        path = str(tmp_path / 'merged.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'merged value'
        ws.merge_cells('A1:C1')
        ws['A2'] = 'plain'
        wb.save(path)

        sheets = profile_workbook(path)
        all_cells = [c for sp in sheets for c in sp.cells]
        refs = {c.ref for c in all_cells}
        # A1 (real value) and A2 (unrelated) both present; B1/C1 (merged
        # members with no value of their own) must not appear as blanks.
        assert 'A1' in refs
        assert 'A2' in refs
        assert 'B1' not in refs
        assert 'C1' not in refs


# ── non-worksheet sheets (chart sheets) ─────────────────────────────────────

class TestChartsheets:
    def test_workbook_with_chartsheet_does_not_crash(self):
        """Fixture 19 has a real chart sheet alongside its data sheet.
        wb.sheetnames/wb[name] includes it, but a Chartsheet has no
        .max_row/.max_column/.iter_rows() — profile_workbook must skip it
        (mirroring lint.py's wb.worksheets iteration) rather than crash."""
        sheets = profile_workbook(_FIXTURE_19)
        assert sheets  # at least the real data sheet was profiled
        assert all(sp.cells for sp in sheets if sp.sheet == 'Blutdruckwerte')

    def test_explicit_chartsheet_name_raises_clear_error(self):
        wb = openpyxl.load_workbook(_FIXTURE_19)
        chart_names = [n for n in wb.sheetnames if n not in {ws.title for ws in wb.worksheets}]
        assert chart_names, 'fixture 19 is expected to have a chart sheet'
        with pytest.raises(SecurityError, match='chart sheet'):
            profile_workbook(_FIXTURE_19, sheet_name=chart_names[0])


# ── row/column caps ──────────────────────────────────────────────────────────

class TestSizeCaps:
    def test_truncated_flag_when_sheet_exceeds_max_rows(self, tmp_path):
        path = str(tmp_path / 'big.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        for i in range(1, 11):
            ws.cell(row=i, column=1, value=i)
        wb.save(path)

        sheets = profile_workbook(path, max_rows=5, max_cols=10)
        assert sheets[0].truncated
        assert len(sheets[0].cells) == 5

    def test_not_truncated_when_within_caps(self, tmp_path):
        path = str(tmp_path / 'small.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'x'
        wb.save(path)

        sheets = profile_workbook(path, max_rows=100, max_cols=100)
        assert not sheets[0].truncated


# ── run_profile: exit codes and CLI-level behaviour ─────────────────────────

class TestRunProfileExitCodes:
    def test_exit_1_when_errors_present(self, tmp_path):
        path = str(tmp_path / 'error.xlsx')
        _write_error_workbook(path)
        code = run_profile([path], out=io.StringIO(), quiet=True)
        assert code == 1

    def test_exit_0_when_clean(self):
        code = run_profile([_FIXTURE_01], out=io.StringIO(), quiet=True)
        assert code == 0

    def test_exit_code_independent_of_filter(self, tmp_path):
        # The filter controls what's printed, not the exit code — matches
        # lint's convention (verbose printing more detail doesn't change
        # whether it exits 0 or 1).
        path = str(tmp_path / 'error.xlsx')
        _write_error_workbook(path)
        buf = io.StringIO()
        code_filtered = run_profile([path], out=buf, errors_only=True, quiet=True)
        code_unfiltered = run_profile([path], out=io.StringIO(), quiet=True)
        assert code_filtered == code_unfiltered == 1

    def test_no_excel_files_in_empty_dir_exits_0(self, tmp_path):
        code = run_profile([str(tmp_path)], out=io.StringIO())
        assert code == 0


class TestInvisibleDataAdvisories:
    """profile shares lint's blind spot for Data Model / external connection
    / external link / embedded object data -- same check, reused directly
    from lint.py rather than duplicated (see _invisible_data_advisories)."""

    def _inject_zip_part(self, tmp_path, part_name: str) -> str:
        path = str(tmp_path / 'data.xlsx')
        wb = openpyxl.Workbook()
        wb.active['A1'] = 'hello'
        wb.save(path)
        with zipfile.ZipFile(path, 'a') as zf:
            zf.writestr(part_name, b'fake')
        return path

    def test_no_signal_by_default(self, tmp_path):
        path = str(tmp_path / 'plain.xlsx')
        wb = openpyxl.Workbook()
        wb.active['A1'] = 'hello'
        wb.save(path)
        buf = io.StringIO()
        run_profile([path], out=buf)
        assert 'Data Model' not in buf.getvalue()

    def test_flags_data_model_in_default_output(self, tmp_path):
        path = self._inject_zip_part(tmp_path, 'xl/model/item1.data')
        buf = io.StringIO()
        run_profile([path], out=buf)
        assert 'Data Model' in buf.getvalue()

    def test_flags_data_model_in_verbose_output(self, tmp_path):
        path = self._inject_zip_part(tmp_path, 'xl/model/item1.data')
        buf = io.StringIO()
        run_profile([path], out=buf, verbose=True)
        assert 'Data Model' in buf.getvalue()

    def test_quiet_mode_stays_one_line(self, tmp_path):
        """Quiet mode's contract is one summary line per file -- the
        advisory must not break that, even when a signal is present."""
        path = self._inject_zip_part(tmp_path, 'xl/model/item1.data')
        buf = io.StringIO()
        run_profile([path], out=buf, quiet=True)
        lines = [l for l in buf.getvalue().splitlines() if l.strip()]
        assert len(lines) == 1

    def test_flagged_in_xlsx_output_mode_too(self, tmp_path):
        path = self._inject_zip_part(tmp_path, 'xl/embeddings/oleObject1.bin')
        out_path = str(tmp_path / 'report.xlsx')
        buf = io.StringIO()
        run_profile([path], out=buf, fmt='xlsx', output=out_path)
        assert 'embedded object' in buf.getvalue()


class TestRunProfileOutput:
    def test_writes_json_file(self, tmp_path):
        out_path = str(tmp_path / 'profile.json')
        run_profile([_FIXTURE_01], out=io.StringIO(), fmt='json', output=out_path)
        assert os.path.exists(out_path)
        with open(out_path, encoding='utf-8') as fh:
            data = json.load(fh)
        assert len(data) == 16

    def test_writes_colored_xlsx(self, tmp_path):
        src_path = str(tmp_path / 'error.xlsx')
        sheet_title = _write_error_workbook(src_path)
        out_path = str(tmp_path / 'profile.xlsx')
        run_profile([src_path], out=io.StringIO(), fmt='xlsx', output=out_path)
        assert os.path.exists(out_path)
        wb = openpyxl.load_workbook(out_path)
        # A1's error cell should carry the error fill.
        cell = wb[sheet_title]['A1']
        assert cell.fill.fgColor.rgb == '00FFC7CE'

    def test_refuses_overwrite_without_force(self, tmp_path):
        out_path = str(tmp_path / 'profile.json')
        with open(out_path, 'w') as fh:
            fh.write('{}')
        code = run_profile([_FIXTURE_01], out=io.StringIO(), fmt='json', output=out_path)
        assert code == 1  # refused

    def test_force_allows_overwrite(self, tmp_path):
        out_path = str(tmp_path / 'profile.json')
        with open(out_path, 'w') as fh:
            fh.write('{}')
        code = run_profile([_FIXTURE_01], out=io.StringIO(), fmt='json', output=out_path, force=True)
        assert code == 0
        with open(out_path, encoding='utf-8') as fh:
            data = json.load(fh)
        assert len(data) == 16

    def test_multi_file_xlsx_output_requires_directory(self, tmp_path):
        out_path = str(tmp_path / 'single.xlsx')
        code = run_profile([_FIXTURE_01, _FIXTURE_24], out=io.StringIO(),
                           fmt='xlsx', output=out_path)
        assert code == 1

    def test_multi_file_output_as_directory(self, tmp_path):
        out_dir = str(tmp_path / 'out')
        run_profile([_FIXTURE_01, _FIXTURE_24], out=io.StringIO(),
                    fmt='json', output=out_dir)
        assert os.path.exists(os.path.join(out_dir, '01_simple_invoice_data.json'))
        assert os.path.exists(os.path.join(out_dir, '24_type_tests_data.json'))
