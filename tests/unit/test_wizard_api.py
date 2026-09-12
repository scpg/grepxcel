"""Unit tests for the grepxcel web wizard FastAPI backend (wizard_api.py).

Uses FastAPI's built-in TestClient (sync wrapper around httpx) so no real
server or browser is needed.  Tests are skipped automatically when FastAPI /
uvicorn are not installed (same pattern as test_wizard_tui.py for Textual).
"""
from __future__ import annotations

import csv
import io
from pathlib import Path

import openpyxl
import pytest

# ── Optional import guard ──────────────────────────────────────────────────

try:
    from fastapi.testclient import TestClient
    from grepxcel.wizard_api import create_app, _STATE
    _API_OK = True
except ImportError:
    _API_OK = False

_skip_no_api = pytest.mark.skipif(
    not _API_OK,
    reason='fastapi / uvicorn not installed (pip install "grepxcel[web]")',
)

# ── Fixture helpers ────────────────────────────────────────────────────────

FIXTURE_01 = Path(__file__).parent.parent / 'fixtures/01_simple_invoice/01_simple_invoice_data.xlsx'


def _make_client(tmp_path: Path, cells: dict | None = None) -> 'TestClient':
    """Create a TestClient backed by a minimal xlsx workbook."""
    if cells is None:
        cells = {
            (1, 1): 'Invoice No:',
            (1, 2): 'AB123456',
            (1, 3): 'Date:',
            (1, 4): '2026-01-15',
            (2, 1): 'Amount:',
            (2, 2): 1234.56,
        }
    xlsx_path = str(tmp_path / 'test.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    for (r, c), v in cells.items():
        ws.cell(row=r, column=c, value=v)
    wb.save(xlsx_path)

    app = create_app(xlsx_path)
    return TestClient(app)


def _make_real_client() -> 'TestClient':
    """Create a TestClient backed by fixture 01 (real file)."""
    assert FIXTURE_01.exists(), f'Fixture not found: {FIXTURE_01}'
    app = create_app(str(FIXTURE_01))
    return TestClient(app)


# ── Tests: page + basic state ──────────────────────────────────────────────

@_skip_no_api
class TestIndexPage:
    def test_returns_html(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.get('/')
        assert r.status_code == 200
        assert 'text/html' in r.headers['content-type']

    def test_html_contains_grid_element(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.get('/')
        assert 'grid-table' in r.text

    def test_html_contains_sheet_name(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.get('/')
        assert 'Sheet1' in r.text


@_skip_no_api
class TestSheetEndpoint:
    def test_returns_rows(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/sheet').json()
        assert 'rows' in d
        assert len(d['rows']) > 0

    def test_col_letters_present(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/sheet').json()
        assert 'col_letters' in d
        assert d['col_letters'][0] == 'A'

    def test_cell_has_expected_keys(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/sheet').json()
        cell = d['rows'][0][0]  # A1
        for key in ('ref', 'row', 'col', 'value', 'choice', 'empty'):
            assert key in cell, f'Missing key: {key}'

    def test_a1_value(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/sheet').json()
        a1 = d['rows'][0][0]
        assert a1['ref'] == 'A1'
        assert a1['value'] == 'Invoice No:'
        assert not a1['empty']

    def test_max_row_max_col(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/sheet').json()
        assert d['max_row'] >= 1
        assert d['max_col'] >= 1


@_skip_no_api
class TestStateEndpoint:
    def test_returns_config_and_stats(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/state').json()
        assert 'config' in d
        assert 'stats' in d
        assert 'choices' in d

    def test_default_direction_is_LR(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/state').json()
        assert d['config']['direction'] == 'LR'

    def test_stats_start_at_zero(self, tmp_path):
        client = _make_client(tmp_path)
        stats = client.get('/api/state').json()['stats']
        assert stats['L'] == 0
        assert stats['V'] == 0
        assert stats['classified'] == 0

    def test_total_counts_nonempty_cells(self, tmp_path):
        client = _make_client(tmp_path, {(1, 1): 'A', (1, 2): 'B', (2, 1): None})
        stats = client.get('/api/state').json()['stats']
        assert stats['total'] == 2  # None cell not counted


@_skip_no_api
class TestClassifyLabel:
    def test_classify_L_ok(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'invoice_no_label', 'type': 'string',
                       'match_mode': '(default)', 'match': 'Invoice No:', 'notes': ''},
        })
        assert r.status_code == 200
        d = r.json()
        assert d['ok']
        assert d['ref'] == 'A1'
        assert d['action'] == 'L'

    def test_classify_L_updates_stats(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'lbl', 'type': 'string', 'match_mode': '(default)',
                       'match': 'Invoice No:', 'notes': ''},
        })
        stats = client.get('/api/state').json()['stats']
        assert stats['L'] == 1
        assert stats['classified'] == 1

    def test_classify_L_stored_in_choices(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'my_label', 'type': 'string', 'match_mode': '(default)',
                       'match': 'Invoice No:', 'notes': ''},
        })
        choices = client.get('/api/state').json()['choices']
        assert 'A1' in choices
        assert choices['A1']['choice'] == 'L'
        assert choices['A1']['name'] == 'my_label'

    def test_classify_L_with_regexp_mode(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'lbl', 'type': 'string', 'match_mode': 'regexp',
                       'match': 'Invoice.*', 'notes': ''},
        })
        assert r.status_code == 200
        choices = client.get('/api/state').json()['choices']
        assert choices['A1']['lbl_mode'] == 'regexp'


@_skip_no_api
class TestClassifyValue:
    def test_classify_V_ok(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'invoice_no', 'type': 'string',
                       'match_mode': '(default)', 'match': '.*',
                       'modifiers': 'none', 'notes': ''},
        })
        assert r.status_code == 200
        assert r.json()['ok']

    def test_classify_V_with_not_null_modifier(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'invoice_no', 'type': 'string',
                       'match_mode': '(default)', 'match': r'[A-Z]{2}\d+',
                       'modifiers': 'not-null', 'notes': ''},
        })
        choices = client.get('/api/state').json()['choices']
        assert choices['B1']['col_a_extra'] == 'not-null'

    def test_classify_V_combo_modifier(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'f', 'type': 'string', 'match_mode': '(default)',
                       'match': '.*', 'modifiers': 'not-null:trim-whitespace', 'notes': ''},
        })
        choices = client.get('/api/state').json()['choices']
        assert 'not-null' in choices['B1']['col_a_extra']
        assert 'trim-whitespace' in choices['B1']['col_a_extra']


@_skip_no_api
class TestClassifyOther:
    def test_classify_C(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'A1', 'action': 'C', 'fields': {'name': 'section_header'},
        })
        assert r.status_code == 200
        choices = client.get('/api/state').json()['choices']
        assert choices['A1']['choice'] == 'C'
        assert choices['A1']['name'] == 'section_header'

    def test_classify_I(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'A1', 'action': 'I', 'fields': {},
        })
        assert r.status_code == 200
        choices = client.get('/api/state').json()['choices']
        assert choices['A1']['choice'] == 'I'

    def test_classify_T(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={
            'ref': 'A1', 'action': 'T', 'fields': {'name': 'my_table', 'notes': ''},
        })
        assert r.status_code == 200
        choices = client.get('/api/state').json()['choices']
        assert choices['A1']['choice'] == 'T'

    def test_classify_CLEAR(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L', 'fields': {'name': 'lbl', 'type': 'string',
            'match_mode': '(default)', 'match': 'x', 'notes': ''},
        })
        r = client.post('/api/classify', json={'ref': 'A1', 'action': 'CLEAR', 'fields': {}})
        assert r.status_code == 200
        choices = client.get('/api/state').json()['choices']
        assert 'A1' not in choices

    def test_classify_invalid_action_400(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/classify', json={'ref': 'A1', 'action': 'X', 'fields': {}})
        assert r.status_code == 400

    def test_classify_with_note(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'f', 'type': 'string', 'match_mode': '(default)',
                       'match': '.*', 'modifiers': 'none', 'notes': 'My test note'},
        })
        notes = client.get('/api/state').json()['notes']
        assert notes.get('B1') == 'My test note'


@_skip_no_api
class TestCellDetail:
    def test_cell_detail_existing(self, tmp_path):
        client = _make_client(tmp_path)
        # Classify first
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'inv_label', 'type': 'string',
                       'match_mode': '(default)', 'match': 'Invoice No:', 'notes': ''},
        })
        d = client.get('/api/cell/A1').json()
        assert d['ref'] == 'A1'
        assert d['choice'] == 'L'
        assert d['name'] == 'inv_label'
        assert d['raw'] == 'Invoice No:'

    def test_cell_detail_unclassified(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/cell/A1').json()
        assert d['choice'] == ''
        assert d['inferred_type'] == 'string'

    def test_cell_detail_number_type(self, tmp_path):
        client = _make_client(tmp_path)
        d = client.get('/api/cell/B2').json()
        assert d['inferred_type'] == 'number'

    def test_cell_detail_invalid_ref(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.get('/api/cell/ZZZ9999')
        # Should succeed (no explicit 404 for out-of-range but valid format)
        assert r.status_code in (200, 404)


@_skip_no_api
class TestUndo:
    def test_undo_reverses_classification(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'lbl', 'type': 'string', 'match_mode': '(default)',
                       'match': 'Invoice No:', 'notes': ''},
        })
        assert client.get('/api/state').json()['stats']['L'] == 1
        r = client.post('/api/undo')
        assert r.status_code == 200
        assert r.json()['ok']
        assert client.get('/api/state').json()['stats']['L'] == 0

    def test_undo_nothing(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/undo')
        assert r.status_code == 200
        assert not r.json()['ok']

    def test_undo_multiple_times(self, tmp_path):
        client = _make_client(tmp_path)
        for action, ref in [('L', 'A1'), ('V', 'B1'), ('C', 'A2')]:
            fields = {'name': 'x', 'type': 'string', 'match_mode': '(default)', 'match': 'y', 'notes': ''}
            if action == 'V':
                fields['modifiers'] = 'none'
            client.post('/api/classify', json={'ref': ref, 'action': action, 'fields': fields})
        assert client.get('/api/state').json()['stats']['classified'] == 3
        client.post('/api/undo')
        assert client.get('/api/state').json()['stats']['classified'] == 2
        client.post('/api/undo')
        assert client.get('/api/state').json()['stats']['classified'] == 1


@_skip_no_api
class TestConfig:
    def test_save_config_ok(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.post('/api/config', json={
            'direction': 'TD', 'template': True,
            'ignore_case': True, 'trim_whitespace': True,
            'currency_sign': '$', 'lbl_match': 'regexp', 'var_match': '',
            'empty_aliases': ['N/A', '—'],
        })
        assert r.status_code == 200
        assert r.json()['ok']

    def test_config_persists_in_state(self, tmp_path):
        client = _make_client(tmp_path)
        client.post('/api/config', json={
            'direction': 'TD', 'template': False,
            'ignore_case': False, 'trim_whitespace': False,
            'currency_sign': '$', 'lbl_match': '', 'var_match': '',
            'empty_aliases': [],
        })
        cfg = client.get('/api/state').json()['config']
        assert cfg['direction'] == 'TD'
        assert cfg['currency_sign'] == '$'


@_skip_no_api
class TestPreview:
    def _classify_some(self, client):
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'invoice_no_label', 'type': 'string',
                       'match_mode': '(default)', 'match': 'Invoice No:', 'notes': ''},
        })
        client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'invoice_no', 'type': 'string',
                       'match_mode': '(default)', 'match': '.*',
                       'modifiers': 'none', 'notes': ''},
        })

    def test_preview_returns_csv(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        r = client.get('/api/preview')
        assert r.status_code == 200
        assert 'text/plain' in r.headers['content-type']

    def test_preview_csv_is_parseable(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        text = client.get('/api/preview').text
        rows = list(csv.reader(io.StringIO(text)))
        assert len(rows) > 0

    def test_preview_contains_lbl_row(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        text = client.get('/api/preview').text
        assert 'lbl:' in text or 'invoice_no_label' in text

    def test_preview_contains_var_row(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        text = client.get('/api/preview').text
        assert 'var:' in text or 'invoice_no' in text

    def test_preview_config_header(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        text = client.get('/api/preview').text
        assert 'config:' in text

    def test_preview_empty_when_no_classifications(self, tmp_path):
        client = _make_client(tmp_path)
        r = client.get('/api/preview')
        assert r.status_code == 200
        # Should at least return config row
        assert 'config:' in r.text or r.text.strip() == ''


@_skip_no_api
class TestSave:
    def _classify_some(self, client):
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'invoice_no_label', 'type': 'string',
                       'match_mode': '(default)', 'match': 'Invoice No:', 'notes': ''},
        })

    def test_save_creates_file(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        r = client.post('/api/save', json={})
        assert r.status_code == 200
        d = r.json()
        assert d['ok']
        assert Path(d['saved_to']).exists()

    def test_save_file_is_valid_csv(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        r = client.post('/api/save', json={})
        saved = Path(r.json()['saved_to'])
        text = saved.read_text(encoding='utf-8')
        rows = list(csv.reader(io.StringIO(text)))
        assert len(rows) > 0

    def test_save_custom_filename(self, tmp_path):
        client = _make_client(tmp_path)
        self._classify_some(client)
        r = client.post('/api/save', json={'filename': 'my_output.csv'})
        assert r.status_code == 200
        d = r.json()
        assert d['saved_to'].endswith('my_output.csv')
        assert Path(d['saved_to']).exists()


@_skip_no_api
class TestWithRealFixture:
    """Tests against the real fixture 01 file."""

    def test_fixture01_sheet_loads(self):
        if not FIXTURE_01.exists():
            pytest.skip('fixture 01 not found')
        client = _make_real_client()
        d = client.get('/api/sheet').json()
        assert d['max_row'] >= 1
        assert d['max_col'] >= 1
        # Fixture 01 has Invoice No: in A1
        a1 = d['rows'][0][0]
        assert a1['value'] == 'Invoice No:'

    def test_fixture01_classify_and_preview(self):
        if not FIXTURE_01.exists():
            pytest.skip('fixture 01 not found')
        client = _make_real_client()
        client.post('/api/classify', json={
            'ref': 'A1', 'action': 'L',
            'fields': {'name': 'invoice_no_label', 'type': 'string',
                       'match_mode': '(default)', 'match': 'Invoice No:', 'notes': ''},
        })
        client.post('/api/classify', json={
            'ref': 'B1', 'action': 'V',
            'fields': {'name': 'invoice_no', 'type': 'string',
                       'match_mode': '(default)', 'match': r'[A-Z]{2}\d+',
                       'modifiers': 'not-null', 'notes': ''},
        })
        text = client.get('/api/preview').text
        assert 'invoice_no_label' in text
        assert 'invoice_no' in text
        assert 'not-null' in text
