#!/usr/bin/env python3
"""
End-to-end verification of the 'grepxcel draft' feature (LLM-assisted pattern drafter).

Run from the project root:
    python3 scripts/verify_draft.py

No model download required — all LLM calls are stubbed.
The script exits with code 0 if every test passes, 1 otherwise.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _bootstrap import ensure_venv; ensure_venv()

import datetime
import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).parent.parent))

import openpyxl

from grepxcel.utils import infer_cell_type
from grepxcel.suggester import ExcelAnalyzer, PatternWriter, PatternSuggester
from grepxcel.model_manager import ModelManager, MODEL_FILENAME, _CHECK_INTERVAL


# ── terminal colours ──────────────────────────────────────────────────────────

GREEN = "\033[32m"
RED   = "\033[31m"
RESET = "\033[0m"
BOLD  = "\033[1m"

passed = 0
failed = 0


def ok(label: str) -> None:
    global passed
    passed += 1
    print(f"  {GREEN}PASS{RESET}  {label}")


def fail(label: str, reason: str) -> None:
    global failed
    failed += 1
    print(f"  {RED}FAIL{RESET}  {label}")
    print(f"         → {reason}")


def section(title: str) -> None:
    print(f"\n{BOLD}{title}{RESET}")
    print("─" * 60)


def check(label: str, condition: bool, reason: str = "") -> None:
    if condition:
        ok(label)
    else:
        fail(label, reason or "assertion failed")


# ── xlsx helpers ──────────────────────────────────────────────────────────────

def make_table_xlsx(tmp: Path) -> str:
    """Invoice-style table: header row + 3 data rows."""
    path = str(tmp / "table_invoice.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "InvoiceLines"
    ws.append(["Product",   "Qty", "Unit Price",   "Line Total"])
    ws.append(["Widget A",  10,    9.99,            99.90])
    ws.append(["Gadget B",  2,     49.50,           99.00])
    ws.append(["Sprocket C",5,     12.00,           60.00])
    wb.save(path)
    return path


def make_kv_xlsx(tmp: Path) -> str:
    """Key-value invoice header (scattered cells, single title in row 1)."""
    path = str(tmp / "kv_invoice.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"
    # Row 1 must have only ONE non-empty cell to trigger KV heuristic.
    # Use non-whole floats so infer_cell_type returns 'currency', not 'integer'.
    ws.append(["Invoice Summary"])
    ws.append(["Invoice No", "INV-0042"])
    ws.append(["Date",       datetime.date(2024, 6, 15)])
    ws.append(["Due Date",   datetime.date(2024, 7, 15)])
    ws.append(["Subtotal",   1250.50])
    ws.append(["VAT",        250.10])
    ws.append(["Total",      1500.60])
    wb.save(path)
    return path


def make_mixed_xlsx(tmp: Path) -> str:
    """More realistic file: KV header + table body on same sheet."""
    path = str(tmp / "mixed_report.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append(["Q1 Sales Report"])           # single cell → KV heuristic
    ws.append(["Region", "UK"])
    ws.append(["Period",  "Jan-Mar 2024"])
    ws.append([])
    ws.append(["Category", "Units", "Revenue"])
    ws.append(["Hardware",  120,    14400.00])
    ws.append(["Software",   45,     9000.00])
    wb.save(path)
    return path


# ── 1. infer_cell_type ────────────────────────────────────────────────────────

section("1. infer_cell_type()")

cases = [
    ("empty list → string",   [],                          "string"),
    ("all None → string",     [None, None],                "string"),
    ("integers",              [1, 2, 3],                   "integer"),
    ("whole floats → integer",[1.0, 2.0],                  "integer"),
    ("decimals → currency",   [9.99, 49.50],               "currency"),
    ("strings",               ["hello", "world"],          "string"),
    ("booleans → string",     [True, False],               "string"),
    ("date objects",          [datetime.date(2024,1,1)],   "date"),
    ("datetime objects",      [datetime.datetime(2024,1,1,12,0)], "datetime"),
    ("majority wins",         [1, 2, 3, "text"],           "integer"),
    ("None ignored",          [None, 42, None],            "integer"),
]

for label, values, expected in cases:
    result = infer_cell_type(values)
    check(f"{label}  ({expected!r})", result == expected,
          f"got {result!r}")


# ── 2. ExcelAnalyzer — table layout ──────────────────────────────────────────

section("2. ExcelAnalyzer — table layout")

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    path = make_table_xlsx(tmp)
    result = ExcelAnalyzer(path).analyse()

    check("detects TABLE keyword",         "TABLE"       in result, repr(result[:200]))
    check("sheet name present",            "InvoiceLines" in result, repr(result[:200]))
    check("header 'Product' listed",       "Product"     in result, repr(result[:200]))
    check("header 'Unit Price' listed",    "Unit Price"  in result, repr(result[:200]))
    check("column type 'string' present",  "string"      in result, repr(result[:200]))
    check("column type 'integer' present", "integer"     in result, repr(result[:200]))
    check("column type 'currency' present","currency"    in result, repr(result[:200]))
    check("data row count mentioned",      "3" in result or "rows" in result.lower(),
          repr(result[:200]))


# ── 3. ExcelAnalyzer — key-value layout ──────────────────────────────────────

section("3. ExcelAnalyzer — KV layout")

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    path = make_kv_xlsx(tmp)
    result = ExcelAnalyzer(path).analyse()

    check("detects KEY-VALUE keyword", "KEY-VALUE"  in result, repr(result[:300]))
    check("'Invoice No' pair found",   "Invoice No" in result, repr(result[:300]))
    check("'INV-0042' value present",  "INV-0042"   in result, repr(result[:300]))
    check("date type inferred",        "date"        in result, repr(result[:300]))
    check("currency type inferred",    "currency"    in result, repr(result[:300]))


# ── 4. ExcelAnalyzer — empty sheet ───────────────────────────────────────────

section("4. ExcelAnalyzer — edge cases")

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    # Empty sheet
    empty_path = str(tmp / "empty.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Empty"
    wb.save(empty_path)
    result = ExcelAnalyzer(empty_path).analyse()
    check("empty sheet: 'Empty' or 'no data' in result",
          "Empty" in result or "no data" in result.lower(),
          repr(result))

    # Single-column file
    single_col_path = str(tmp / "single.xlsx")
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.append(["Label"])
    ws2.append(["Value"])
    ws2.append(["Data"])
    wb2.save(single_col_path)
    result2 = ExcelAnalyzer(single_col_path).analyse()
    check("single-column: does not crash", True)  # just verifying no exception


# ── 5. ModelManager — state management (no network) ──────────────────────────

section("5. ModelManager — state (no network)")

with tempfile.TemporaryDirectory() as td:
    cache = Path(td)
    m = ModelManager(cache_dir=cache)

    # No state yet
    check("_load_state() on missing file → {}", m._load_state() == {})

    # _check_due with empty state
    check("_check_due({}) → True", m._check_due({}) is True)

    # Recent check → not due
    recent = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=1)
    not_due = m._check_due({"last_check": recent.isoformat()})
    check("_check_due (1 h ago) → False", not_due is False)

    # Old check → due
    old = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(seconds=_CHECK_INTERVAL + 60)
    due = m._check_due({"last_check": old.isoformat()})
    check("_check_due (>24 h ago) → True", due is True)

    # Corrupt timestamp → due
    check("_check_due (corrupt) → True",
          m._check_due({"last_check": "not-a-date"}) is True)

    # Save and reload
    m._save_state("deadbeef")
    state = m._load_state()
    check("_save_state/_load_state roundtrip commit_hash",
          state.get("commit_hash") == "deadbeef")
    check("_save_state/_load_state roundtrip filename",
          state.get("filename") == MODEL_FILENAME)
    check("saved state: last_check present", "last_check" in state)

    # After save, check should NOT be due (freshly saved)
    check("_check_due after fresh save → False",
          m._check_due(state) is False)

    # Corrupt state.json
    (cache / "state.json").write_text("{{broken json", encoding="utf-8")
    check("_load_state() on corrupt JSON → {}", m._load_state() == {})

    # GREPXCEL_MODEL_DIR env var
    custom_dir = str(cache / "custom_cache")
    os.environ["GREPXCEL_MODEL_DIR"] = custom_dir
    from grepxcel.model_manager import default_cache_dir
    check("GREPXCEL_MODEL_DIR overrides cache dir",
          default_cache_dir() == Path(custom_dir))
    del os.environ["GREPXCEL_MODEL_DIR"]

    # Network failure must not advance last_check
    old_state_time = (
        datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=2)
    ).isoformat()
    m._save_state.__func__  # verify it exists
    # Manually write old state
    import json as _json
    (cache / "state.json").write_text(
        _json.dumps({"last_check": old_state_time, "commit_hash": "abc", "filename": MODEL_FILENAME}),
        encoding="utf-8",
    )
    with patch.object(m, "_remote_commit", return_value=None):
        m._maybe_update(verbose=False)
    reloaded = m._load_state()
    saved_time = datetime.datetime.fromisoformat(reloaded["last_check"])
    elapsed = (datetime.datetime.now(datetime.timezone.utc) - saved_time).total_seconds()
    check("network failure: last_check NOT advanced",
          elapsed > _CHECK_INTERVAL,
          f"elapsed={elapsed:.0f}s expected >{_CHECK_INTERVAL}s")


# ── 6. PatternWriter ──────────────────────────────────────────────────────────

section("6. PatternWriter — xlsx round-trip")

SAMPLE_LLM_OUTPUT = """\
config: | read.direction | LR
def: | InvoiceNo | string | INV-\\d+
def: | Total | currency | \\d+(\\.\\d{2})?
def: | Product | string | .*
def: | Qty | integer | \\d+
def: | OrderDate | date |
START:
cell:1 | InvoiceNo
cell:1 | Total
cell:1 | OrderDate
table:*
 | HEADER:1 | Product | Qty
 | DATA:* | Product | Qty
END:
"""

with tempfile.TemporaryDirectory() as td:
    out = str(Path(td) / "pattern.xlsx")
    PatternWriter().write(SAMPLE_LLM_OUTPUT, out)
    check("output file created", Path(out).exists())

    wb = openpyxl.load_workbook(out)
    ws = wb.active

    # config row
    r1 = [ws.cell(row=1, column=c).value for c in range(1, 5)]
    check("row 1 col A = 'config:'",        r1[0] == "config:")
    check("row 1 col B = 'read.direction'", r1[1] == "read.direction")
    check("row 1 col C = 'LR'",             r1[2] == "LR")

    # def rows
    r2 = [ws.cell(row=2, column=c).value for c in range(1, 5)]
    check("row 2 col A = 'def:'",      r2[0] == "def:")
    check("row 2 col B = 'InvoiceNo'", r2[1] == "InvoiceNo")
    check("row 2 col C = 'string'",    r2[2] == "string")

    # START / END in col A
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    check("START: in col A", "START:" in col_a)
    check("END: in col A",   "END:"   in col_a)

    # table template row: col A must be blank
    header_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=2).value == "HEADER:1":
            header_row = r
            break
    if header_row:
        check("HEADER:1 row: col A is None",
              ws.cell(row=header_row, column=1).value is None)
        check("HEADER:1 row: col C = 'Product'",
              ws.cell(row=header_row, column=3).value == "Product")
    else:
        fail("HEADER:1 row present", "row not found in xlsx")

    # Markdown fences must be skipped
    llm_with_fences = "```\ndef: | X | string | .*\n```"
    out2 = str(Path(td) / "fences.xlsx")
    PatternWriter().write(llm_with_fences, out2)
    wb2 = openpyxl.load_workbook(out2)
    ws2 = wb2.active
    col_a2 = [ws2.cell(row=r, column=1).value for r in range(1, ws2.max_row + 1)]
    check("markdown fences not written to col A",
          "```" not in (col_a2 or []))

    # Bare table keyword (no pipes)
    bare = "table:*\n | HEADER:1 | Col\n | DATA:* | Col\nEND:"
    out3 = str(Path(td) / "bare.xlsx")
    PatternWriter().write(bare, out3)
    wb3 = openpyxl.load_workbook(out3)
    ws3 = wb3.active
    check("bare 'table:*' in col A",
          ws3.cell(row=1, column=1).value == "table:*")


# ── 7. PatternSuggester — mocked LLM, full pipeline ──────────────────────────

section("7. PatternSuggester — full pipeline (mocked LLM)")

LLM_RESPONSE = """\
def: | Product | string | .*
def: | Qty | integer | \\d+
def: | UnitPrice | currency | \\d+\\.\\d{2}
START:
table:*
 | HEADER:1 | Product | Qty | UnitPrice
 | DATA:* | Product | Qty | UnitPrice
END:
"""

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    data_path = make_table_xlsx(tmp)
    out_path  = str(tmp / "suggested.xlsx")

    mock_manager = MagicMock()
    mock_manager.return_value.ensure_ready.return_value = tmp / "dummy_model.gguf"
    mock_client  = MagicMock()
    mock_client.return_value.chat.return_value = LLM_RESPONSE

    with patch("grepxcel.suggester.ModelManager", mock_manager), \
         patch("grepxcel.suggester.LlamaCppClient", mock_client):
        s    = PatternSuggester(input_path=data_path, output_path=out_path)
        code = s.run()

    check("PatternSuggester.run() returns 0", code == 0, f"exit code = {code}")
    check("suggested pattern.xlsx created",   Path(out_path).exists())

    # Verify the xlsx has reasonable content
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    check("pattern.xlsx: 'def:' rows present",  "def:"   in col_a)
    check("pattern.xlsx: 'START:' present",      "START:" in col_a)
    check("pattern.xlsx: 'END:' present",        "END:"   in col_a)
    check("pattern.xlsx: 'table:*' present",     "table:*" in col_a)

    # Security error path: missing file → exit code 1
    s2   = PatternSuggester(input_path=str(tmp / "nonexistent.xlsx"),
                            output_path=str(tmp / "out2.xlsx"))
    code2 = s2.run()
    check("missing input file → exit code 1", code2 == 1, f"exit code = {code2}")


# ── 8. Round-trip: suggest → extract ─────────────────────────────────────────

section("8. Round-trip: PatternWriter output → grepxcel extract")

# The engine validates HEADER rows in strict mode: each cell must pass type
# validation against its field definition.  The Excel header row contains text
# labels ("Qty") which would fail an integer-type check.  Use IGNORE for all
# HEADER columns so the engine simply anchors on the header row without
# validating its values.  DATA rows are validated with strict=False (warnings
# only), so integer/currency validation still exercises the right code paths.
#
# Engine table result structure: tables[i] has keys 'headers', 'data', 'footers'
# (not 'rows').

ROUNDTRIP_PATTERN = """\
def: | Product | string | .*
def: | Qty | integer | \\d+
def: | UnitPrice | currency | .*
def: | LineTotal | currency | .*
START:
table:*
 | HEADER:1 | IGNORE | IGNORE | IGNORE | IGNORE
 | DATA:* | Product | Qty | UnitPrice | LineTotal
END:
"""

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    data_path    = make_table_xlsx(tmp)
    pattern_path = str(tmp / "pattern.xlsx")
    PatternWriter().write(ROUNDTRIP_PATTERN, pattern_path)

    from grepxcel.engine import Engine
    from grepxcel.logger import Logger, VerbosityLevel

    logger = Logger(level=VerbosityLevel.NORMAL)
    try:
        result = Engine().process(
            pattern_path, data_path, logger=logger,
            max_file_mb=5, max_uncompressed_mb=50,
            max_cell_len=1000,
        )
        tables = result.get("tables", [])
        check("extract: at least one table returned",
              len(tables) >= 1,
              f"tables={tables!r}")

        if tables:
            rows = tables[0].get("data", [])   # engine key is 'data', not 'rows'
            check("extract: table has 3 data rows",
                  len(rows) == 3,
                  f"got {len(rows)} rows: {rows!r}")

            first = rows[0] if rows else {}
            check("extract: 'Product' field present in row",
                  "Product" in first,
                  f"keys: {list(first.keys())}")
            check("extract: 'Product' value is 'Widget A'",
                  first.get("Product") == "Widget A",
                  f"got {first.get('Product')!r}")
            check("extract: 'Qty' field is 10",
                  first.get("Qty") == 10,
                  f"got {first.get('Qty')!r}")
    except Exception as exc:
        fail("extract: Engine().process() raised", str(exc))
    finally:
        logger.close()


# ── 9. CLI integration: grepxcel extract via subprocess ──────────────────────

section("9. CLI smoke test — grepxcel extract (subprocess)")

import subprocess

with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    data_path    = make_table_xlsx(tmp)
    pattern_path = str(tmp / "pattern.xlsx")
    PatternWriter().write(ROUNDTRIP_PATTERN, pattern_path)

    result = subprocess.run(
        ["grepxcel", "extract", "-p", pattern_path, str(data_path)],
        capture_output=True, text=True,
    )
    check("CLI exit code = 0", result.returncode == 0,
          f"stderr: {result.stderr[:300]}")

    try:
        payload = json.loads(result.stdout)
        tables  = payload.get("tables", [])
        check("CLI stdout is valid JSON",    True)
        check("CLI JSON has 'tables' key",   "tables" in payload)
        check("CLI JSON tables non-empty",   len(tables) >= 1,
              f"tables={tables!r}")
        if tables:
            rows = tables[0].get("data", [])   # engine key is 'data', not 'rows'
            check("CLI JSON: 3 rows extracted", len(rows) == 3,
                  f"got {len(rows)}")
    except json.JSONDecodeError as exc:
        fail("CLI stdout is valid JSON", str(exc))


# ── 10. grepxcel suggest --help (no crash) ───────────────────────────────────

section("10. CLI smoke test — grepxcel draft --help")

result = subprocess.run(
    ["grepxcel", "draft", "--help"],
    capture_output=True, text=True,
)
check("grepxcel draft --help exits 0",   result.returncode == 0,
      f"stderr: {result.stderr[:200]}")
check("help mentions 'FILE'",            "FILE"    in result.stdout, result.stdout[:300])
check("help mentions '--output'",        "--output" in result.stdout, result.stdout[:300])
check("help mentions '--verbose'",       "--verbose" in result.stdout, result.stdout[:300])


# ── summary ───────────────────────────────────────────────────────────────────

total = passed + failed
print(f"\n{'═' * 60}")
print(f"  {BOLD}Results: {passed}/{total} passed{RESET}", end="")
if failed:
    print(f"   {RED}{failed} FAILED{RESET}")
else:
    print(f"   {GREEN}all good{RESET}")
print(f"{'═' * 60}\n")

sys.exit(0 if failed == 0 else 1)
