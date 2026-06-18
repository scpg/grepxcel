"""
Pattern drafting: analyse a target Excel file and use a locally-running
GGUF model (via llama-cpp-python) to produce a starter grepxcel pattern file.

No data leaves the process during inference — the model runs entirely in-process.
The only network activity is the one-time model download (and daily update checks)
handled by ModelManager.
"""

import os
import platform
import subprocess  # nosec B404 — only fixed-argv, shell-free GPU probes (see below)
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol, runtime_checkable

import openpyxl

from .model_manager import MODEL_CHAT_FORMAT, ModelManager
from .pattern_parser import PatternError, PatternParser
from .security import SecurityError, validate_file
from .utils import infer_cell_type, is_empty


# ── Prompts ───────────────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
You are an expert at generating grepxcel pattern files.

A grepxcel pattern file is an Excel workbook that describes how to extract structured data
from another Excel file. Each row uses up to 4 columns (A, B, C, D).
In your output, separate columns with ' | ' (space-pipe-space).

─── SECTIONS (in order) ───────────────────────────────────────────────────────

1. Config rows (optional):
   config: | read.direction | LR          (or TD for top-down column scanning)
   config: | currency.sign  | €
   config: | ignore.case    | no          (yes = case-insensitive regex matching)

2. Label definitions — anchor cells, NEVER written to output JSON:
   lbl: | FieldName | type | regex

   Use lbl: for literal text that marks where a value lives (e.g. "Invoice No:",
   column headers like "Product", "Qty").  These are matched for position only. They will allow you to confirm that the analysis and extraction of data is being done correctly and that the file being processed respects the structure that has been defined it should have.

3. Variable definitions — extracted to the output JSON:
   var: | field.name | type | regex

   Use var: for every value you want to capture.
   Dot notation creates nested JSON: po.number → {"po": {"number": ...}}
   All var: field names in one table DATA row must share the same group prefix.

   Types:  string  integer  currency  percentage  date  datetime
   Regex:  Python re.fullmatch pattern.  Use .* to match anything.
           Leave blank for date/datetime (type check only, no regex needed).

   Examples:
     var: | po.number    | string     | PO-\\d{4,8}
     var: | inv.total    | currency   | \\d+(\\.\\d{1,2})?
     var: | line.qty     | integer    | \\d+
     var: | line.margin  | percentage |
     var: | inv.date     | date       |

   Important notes:
     - integers can be negative too — include a leading -? in the regex if needed.
     - same applies for currency amounts — they often can be negative (e.g. credit notes, negative adjustments).


4. Extraction sequence between START: and END:
   Important note: Headers are "in general" associated to lables, and data cells to variables.  The pattern must reflect this association by referencing the lbl: names in the HEADER: row and the var: names in the DATA: row. (there are exceptions to these rules, but following them will make the pattern easier to understand and maintain).

   For scattered key-value cells — two addressing modes:

     cell:next | FieldName     (next non-empty cell in scan order → FieldName)
     cell:next | IGNORE        (skip the next non-empty cell — use for lbl: anchors)
     cell:B5   | FieldName     (jump directly to cell B5 → FieldName)
     cell:B5   | IGNORE        (jump to B5 and discard it)

   Prefer cell:next for sequential layouts.
   Use cell:B5 (A1-notation) when you know the exact address — makes the pattern
   self-documenting and independent of scan order.
   Absolute references must appear in forward reading order (LR: row then column).

   To reposition the cursor without reading a cell — use seek:

     seek:G5                (move cursor to G5; the next cell:next starts from G5)

   seek: is for sheets where you must jump backward (or forward past a gap) to
   a new region after reading some scattered absolute cells. It does NOT read the
   target cell — it only sets the cursor position. Use it sparingly; prefer
   cell:next for the common sequential case and cell:B5 for isolated static cells.

   For repeating tables:
     table:*                   (bare keyword, no pipe, starts a table block)
       | HEADER:1 | ColA | ColB | ColC
       | DATA:*   | ColA | ColB | ColC
       | FOOTER:1 | ColA | ColB | ColC
     (Table template rows have a blank column A — start the line with ' | ')

   Row type suffixes:  :1 (exactly one)  :* (greedy)  :{n,m} (bounded: min n, max m total rows)
   Column keywords:  FieldName  IGNORE  EMPTY

   For fixed-slot templates (pre-allocated empty rows before the footer), use DATA:{n,m}
   with one or more SKIP_IF rows to silently skip empty rows:
     table:1
       | HEADER:1  | col_desc | col_qty
       | SKIP_IF   | EMPTY    | IGNORE
       | DATA:{0,15} | line.description | line.qty
       | FOOTER:1  | lbl_total | inv.total
   SKIP_IF uses EMPTY (cell must be null) and IGNORE (don't check). A row matching
   any SKIP_IF condition is silently excluded from output but still counts toward {n,m}.
   SKIP_IF is only valid with DATA:{n,m}.

─── HOW TO READ THE ANALYSIS ───────────────────────────────────────────────────

KEY-VALUE sections list lines shaped like:
    - LABEL 'Invoice No:'  →  VALUE 'AB123456' [string]

Each LABEL → VALUE pair becomes THREE coordinated rows in your pattern:
  1. lbl: | <label_name>     | string | <exact LABEL text>   ← defines the anchor
  2. var: | <group>.<field>  | <type> | <regex>              ← defines the value
  3. inside START:/END:, two sequence steps that alternate:
        cell:next | <label_name>        (consume the label cell — never output)
        cell:next | <group>.<field>     (consume the value cell — goes to output)

WHY: the scanner walks cells left-to-right. The label cell comes first and must be
consumed by its lbl: anchor (or IGNORE) so the cursor lands on the value next.
The label is matched for POSITION ONLY and never appears in the output JSON;
only var: fields appear in the output.

TABLE sections list columns shaped like:
    - HEADER 'Datum' (→ lbl:)  →  DATA [datetime] (→ var:)  samples: ...

For each column produce:
  1. lbl: | <col_name>      | string | <exact HEADER text>   ← the column header anchor
  2. var: | <group>.<field> | <type> | <regex>              ← the column's data
Then build the table block:
    table:*
     | HEADER:1 | <col_name_1> | <col_name_2> | ...   (all lbl: names, in column order)
     | DATA:*   | <group>.f1   | <group>.f2   | ...   (all var: names, same order)
The HEADER: row references ONLY lbl: names; the DATA: row references ONLY var:
names. Never put a var: name in the HEADER: row or an lbl: name in the DATA: row.

If a column's DATA is empty (the analysis says "use IGNORE in the DATA: row"),
still put its lbl: in the HEADER: row, but write IGNORE at that position in the
DATA: row so the column alignment is preserved.

─── FIELD NAMING ───────────────────────────────────────────────────────────────

Group semantically related values under a shared dot-prefix. Pick the prefix from
what the value MEANS, not from the label text next to it:
  - invoice header fields → inv.number, inv.date, inv.due_date
  - monetary amounts      → amount.net, amount.vat, amount.gross
  - party / contact info  → client.name, client.email
  - repeating line items  → line.description, line.qty, line.price
Use consistent, conventional names. Group every monetary total under one prefix
(e.g. amount.*), not scattered across unrelated groups.

─── KEY DESIGN RULES ──────────────────────────────────────────────────────────

- Use lbl: for label cells ("Invoice No:", "Total:", column headers).
- Use var: for the values that follow those labels.
- In HEADER rows, use lbl: field names.  In DATA rows, use var: field names.
- Give var: fields a dot-notation name: group.field (e.g. po.number, line.qty).
- Use percentage for cells that hold a percentage (e.g. 0.625 representing 62.5%).

─── OUTPUT RULES ──────────────────────────────────────────────────────────────

- Output ONLY the pattern rows. No explanations, no markdown fences, no comments.
- One row per line, columns separated by ' | '.
- Table template rows must start with ' | ' (blank column A).
- START: and END: are bare keywords with no pipe separator.

─── EXAMPLE OUTPUT ────────────────────────────────────────────────────────────

config: | read.direction | LR
lbl: | inv_label | string | Invoice No:
lbl: | total_label | string | Total
lbl: | col_product | string | Product
lbl: | col_qty | string | Qty
var: | inv.number | string | INV-\\d+
var: | inv.total | currency | \\d+(\\.\\d{2})?
var: | inv.date | date |
var: | line.product | string | .*
var: | line.qty | integer | \\d+
START:
cell:next | inv_label
cell:next | inv.number
cell:next | total_label
cell:next | inv.total
cell:next | inv.date
table:*
 | HEADER:1 | col_product | col_qty
 | DATA:* | line.product | line.qty
END:
"""

_USER_PROMPT_TEMPLATE = """\
Analyse the Excel structure below and produce a grepxcel pattern file that would extract its key data.

{analysis}

Generate the pattern file now:
"""


# ── Number-format type hint ───────────────────────────────────────────────────

def _type_from_number_format(fmt: str) -> str | None:
    """Derive a grepxcel type from an Excel number format string.

    Conservative — returns None for ambiguous formats so value-based inference
    can serve as the fallback.
    """
    if not fmt or fmt in ('General', '@', '0', '#,##0'):
        return None
    if '%' in fmt:
        return 'percentage'
    fmt_lower = fmt.lower()
    if 'y' in fmt_lower:                          # year token → date family
        return 'datetime' if ('h' in fmt_lower and ':' in fmt) else 'date'
    if any(c in fmt for c in ('$', '€', '£', '¥', '₹')) or '[$' in fmt:
        return 'currency'
    if '#,##0.00' in fmt or ('0.00' in fmt and '#' in fmt):
        return 'currency'
    return None


# ── Excel analyser ────────────────────────────────────────────────────────────

class ExcelAnalyzer:
    """Reads an xlsx file and returns a plain-text structural description for the LLM.

    Two-pass load strategy:
      Pass 1 (data_only=False): formula detection + number-format strings.
      Pass 2 (data_only=True):  cached cell values for type inference and samples.
    """

    _MAX_SAMPLE_ROWS = 100   # rows used for type inference
    _MAX_SAMPLE_COLS = 20    # columns described per section
    _DISPLAY_SAMPLES = 3     # value examples shown per column in the output text
    _ROW_TOLERANCE   = 1.2   # if actual_rows <= MAX × tolerance, include all rows

    def __init__(
        self,
        path: str,
        sheet=None,
        max_file_mb: float = 5.0,
        max_uncompressed_mb: float = 50.0,
    ):
        self.path                = path
        self.sheet               = sheet
        self.max_file_mb         = max_file_mb
        self.max_uncompressed_mb = max_uncompressed_mb

    def analyse(self) -> str:
        validate_file(self.path, self.max_file_mb, self.max_uncompressed_mb)
        wb_raw = openpyxl.load_workbook(self.path, data_only=False)
        wb_val = openpyxl.load_workbook(self.path, data_only=True)
        try:
            preamble = self._workbook_preamble(wb_val)
            ws_raw   = self._select_sheet(wb_raw)
            ws_val   = self._select_sheet(wb_val)
            body     = self._analyse_sheet(ws_raw, ws_val)
        finally:
            wb_raw.close()
            wb_val.close()
        return (preamble + body) if preamble else body

    def _select_sheet(self, wb):
        sheet = self.sheet
        if sheet is None:
            return wb.active

        n = len(wb.worksheets)
        available = ', '.join(wb.sheetnames)

        def _by_index(idx: int):
            # Support negative indices; bound-check with a clear message.
            if -n <= idx < n:
                return wb.worksheets[idx]
            raise ValueError(
                f'Sheet index {idx} is out of range — the workbook has {n} '
                f'sheet(s) (valid: 0..{n - 1}). Available sheets: {available}'
            )

        if isinstance(sheet, int):
            return _by_index(sheet)
        if sheet in wb.sheetnames:
            return wb[sheet]                       # name match wins (incl. "2025")
        if str(sheet).lstrip('-').isdigit():       # numeric string, no such name → index
            return _by_index(int(sheet))
        raise ValueError(
            f'Sheet {sheet!r} not found. Available sheets: {available}'
        )

    def _workbook_preamble(self, wb) -> str:
        """List all sheets with dimensions. Empty for single-sheet workbooks."""
        sheets = wb.worksheets
        if len(sheets) <= 1:
            return ''
        target_title = self._select_sheet(wb).title
        lines = [f'Workbook: {len(sheets)} sheets']
        for i, ws in enumerate(sheets, 1):
            r = ws.max_row or 0
            c = ws.max_column or 0
            marker = ' ← target' if ws.title == target_title else ''
            lines.append(f"  Sheet {i}: '{ws.title}' — {r} rows × {c} columns{marker}")
        return '\n'.join(lines) + '\n\n'

    def _analyse_sheet(self, ws_raw, ws_val) -> str:
        max_row = ws_val.max_row or 0
        max_col = ws_val.max_column or 0

        header = [
            f"Sheet: '{ws_val.title}'",
            f'Dimensions: {max_row} rows × {max_col} columns',
        ]

        if max_row == 0:
            return '\n'.join(header) + '\n\nEmpty sheet — no data found.'

        # Apply tolerance: include all rows when count is only slightly above the cap.
        load_rows = (
            max_row if max_row <= int(self._MAX_SAMPLE_ROWS * self._ROW_TOLERANCE)
            else self._MAX_SAMPLE_ROWS
        )
        load_cols = min(max_col, self._MAX_SAMPLE_COLS)

        formula_cells, number_formats = self._cell_metadata(ws_raw, load_rows, load_cols)

        rows = list(ws_val.iter_rows(
            min_row=1, max_row=load_rows, max_col=load_cols,
            values_only=True,
        ))

        if not rows or all(all(v is None for v in r) for r in rows):
            return '\n'.join(header) + '\n\nEmpty sheet — no data found.'

        notes = []
        if max_row > load_rows:
            notes.append(f'(analysis covers first {load_rows} of {max_row} rows)')
        if max_col > self._MAX_SAMPLE_COLS:
            notes.append(f'(first {self._MAX_SAMPLE_COLS} of {max_col} columns shown)')

        sections = self._split_into_sections(rows)
        if not sections:
            return '\n'.join(header) + '\n\nEmpty sheet — no data found.'

        lines   = header + notes
        is_multi = len(sections) > 1

        for sec_idx, (sec_start, sec_rows) in enumerate(sections, 1):
            sec_end = sec_start + len(sec_rows)
            label   = (
                f'Section {sec_idx} (rows {sec_start + 1}–{sec_end})'
                if is_multi else None
            )
            non_empty_first = [v for v in sec_rows[0] if not is_empty(v)]

            # Titled table: a single-cell heading row (e.g. "INCOME", "EXPENSES")
            # followed by a real table.  Requires ≥ 5 rows in total so that ordinary
            # KV blocks (title + 3 label-value pairs = 4 rows) are not misclassified.
            if len(non_empty_first) == 1 and len(sec_rows) >= 5:
                non_empty_second = [v for v in sec_rows[1] if not is_empty(v)]
                if len(non_empty_second) >= 2:
                    combined_label = (
                        f'{label}, title: {non_empty_first[0]!r}'
                        if label else None
                    )
                    lines += self._describe_table_section(
                        sec_rows[1:], sec_start + 1, formula_cells, number_formats,
                        combined_label)
                    continue

            # TABLE requires >= 2 rows (header + at least one data row)
            # and >= 2 non-empty values in the first row.  A grid whose first row
            # is mostly colon-terminated labels is really a key-value block laid
            # out across columns, not a table — route it to the KV describer.
            if (len(non_empty_first) >= 2 and len(sec_rows) >= 2
                    and not self._looks_like_kv_grid(sec_rows[0])):
                lines += self._describe_table_section(
                    sec_rows, sec_start, formula_cells, number_formats, label)
            else:
                lines += self._describe_kv_section(
                    sec_rows, sec_start, formula_cells, number_formats, label)

        return '\n'.join(lines)

    def _looks_like_kv_grid(self, first_row) -> bool:
        """True if a grid's first row is mostly colon-terminated labels.

        Column headers ("Date", "Amount", "Receipt No") rarely end with ':',
        whereas key-value labels ("Employee:", "Department:") usually do. When
        at least half the non-empty string cells end with ':', the section is a
        KV block written across columns rather than a real table.
        """
        cells = [v for v in first_row if not is_empty(v) and isinstance(v, str)]
        if len(cells) < 2:
            return False
        colon = sum(1 for c in cells if c.rstrip().endswith(':'))
        return colon >= len(cells) / 2

    def _cell_metadata(self, ws, max_row: int, max_col: int) -> tuple[set, dict]:
        """Return (formula_cells, number_formats) using 1-based (row, col) keys."""
        formula_cells  = set()
        number_formats = {}
        for cell_row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in cell_row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    formula_cells.add((cell.row, cell.column))
                fmt = getattr(cell, 'number_format', None)
                if fmt and fmt not in ('General', '@'):
                    number_formats[(cell.row, cell.column)] = fmt
        return formula_cells, number_formats

    def _split_into_sections(self, rows: list) -> list[tuple[int, list]]:
        """Split rows at runs of fully-empty rows. Returns [(0based_start, rows), ...]."""
        sections: list[tuple[int, list]] = []
        current:  list                   = []
        start = 0
        for i, row in enumerate(rows):
            if all(v is None or is_empty(v) for v in row):
                if current:
                    sections.append((start, current))
                    current = []
            else:
                if not current:
                    start = i
                current.append(row)
        if current:
            sections.append((start, current))
        return sections

    def _col_type(
        self, col_idx: int, sec_start: int,
        data_rows: list, number_formats: dict,
    ) -> tuple[str, str | None]:
        """Return (type_str, format_hint_or_None) for one column."""
        vals = [r[col_idx] for r in data_rows if col_idx < len(r)]
        # First non-None format for this column in any data row (1-based indexing)
        fmt = next(
            (number_formats[(sec_start + 2 + ri, col_idx + 1)]
             for ri in range(len(data_rows))
             if (sec_start + 2 + ri, col_idx + 1) in number_formats),
            number_formats.get((sec_start + 1, col_idx + 1)),  # fall back to header row
        )
        return (_type_from_number_format(fmt) or infer_cell_type(vals)), fmt

    def _describe_table_section(
        self, rows, sec_start, formula_cells, number_formats, label,
    ) -> list:
        headers     = list(rows[0])
        data_rows   = rows[1:]
        non_empty_h = [(c, h) for c, h in enumerate(headers) if not is_empty(h)
                       ][:self._MAX_SAMPLE_COLS]

        prefix = (f'{label}: TABLE layout' if label
                  else 'Layout: TABLE (first row appears to be column headers)')
        lines  = [
            '', prefix,
            f'Columns ({len(non_empty_h)}) — for EACH column define an lbl: for its '
            f'HEADER text and a var: for its DATA values. Reference the lbl: names in '
            f'the HEADER: row and the var: names in the DATA: row (same column order). '
            f'A column whose DATA is empty has no var: — put its lbl: in the HEADER: '
            f'row and IGNORE at that position in the DATA: row.',
        ]

        for col_idx, header in non_empty_h:
            col_type, fmt = self._col_type(col_idx, sec_start, data_rows, number_formats)
            is_formula    = any(
                (sec_start + 2 + ri, col_idx + 1) in formula_cells
                for ri in range(len(data_rows))
            )
            vals    = [r[col_idx] for r in data_rows
                       if col_idx < len(r) and not is_empty(r[col_idx])]
            samples = [repr(v) for v in vals[:self._DISPLAY_SAMPLES]]
            if vals:
                line = (f"  - HEADER '{header}' (→ lbl:)  →  "
                        f"DATA [{col_type}] (→ var:)  samples: {', '.join(samples)}")
            else:
                line = (f"  - HEADER '{header}' (→ lbl:)  →  "
                        f"DATA (empty — use IGNORE in the DATA: row)")
            if fmt:
                line += f'  (format: {fmt})'
            if is_formula:
                line += '  [formula]'
            lines.append(line)

        if data_rows:
            lines.append(f'\nData rows: {len(data_rows)}')
        return lines

    def _describe_kv_section(
        self, rows, sec_start, formula_cells, number_formats, label,
    ) -> list:
        prefix = (f'{label}: KEY-VALUE layout' if label
                  else 'Layout: KEY-VALUE (scattered cells, not a standard table)')
        lines  = [
            '', prefix,
            'Label → Value pairs '
            '(LABEL marks position → define with lbl: and skip with cell:next | IGNORE; '
            'VALUE is captured → define with var: and read with cell:next):',
        ]
        count  = 0

        # KV cells alternate LABEL, VALUE across each row. Consume the non-empty
        # cells two at a time: the first is the label, the second is its value.
        # This avoids emitting a spurious pair for every value→next-label adjacency.
        for ri, row in enumerate(rows):
            global_row = sec_start + ri + 1          # 1-based sheet row
            non_empty  = [(ci, v) for ci, v in enumerate(row) if not is_empty(v)]

            i = 0
            while i < len(non_empty):
                lbl_ci, lbl_val = non_empty[i]
                if i + 1 < len(non_empty):
                    val_ci, val_val = non_empty[i + 1]
                    fmt      = number_formats.get((global_row, val_ci + 1))
                    val_type = (_type_from_number_format(fmt)
                                or infer_cell_type([val_val]))
                    is_fml   = (global_row, val_ci + 1) in formula_cells
                    fml_note = '  [formula]' if is_fml else ''
                    lines.append(
                        f"  - LABEL '{lbl_val}'  →  VALUE {repr(val_val)} "
                        f"[{val_type}]{fml_note}"
                    )
                    i += 2
                else:
                    lines.append(f"  - LABEL '{lbl_val}'  →  (no value beside it)")
                    i += 1
                count += 1
                if count >= 50:
                    lines.append('  ... (additional pairs omitted)')
                    return lines
        return lines


# ── Hardware detection ────────────────────────────────────────────────────────

_CUDA_INSTALL  = "CMAKE_ARGS=\"-DGGML_CUDA=on\"  pip install llama-cpp-python --force-reinstall"
_METAL_INSTALL = "CMAKE_ARGS=\"-DGGML_METAL=on\" pip install llama-cpp-python --force-reinstall"
_VULKAN_INSTALL= "CMAKE_ARGS=\"-DGGML_VULKAN=on\" pip install llama-cpp-python --force-reinstall"


def _detect_gpu() -> tuple[int, str | None, str | None]:
    """
    Probe the local machine for GPU/NPU acceleration and return
    (n_gpu_layers, backend_label, install_hint).

    Priority:
      1. GREPXCEL_GPU_LAYERS env var — user override, no messages printed.
      2. NVIDIA GPU via nvidia-smi (CUDA).
      3. Apple Silicon via platform check (Metal).
      4. Any Vulkan-capable GPU via vulkaninfo (Windows/Linux fallback).
      5. CPU-only fallback (n_gpu_layers=0).

    NPU note: Intel NPU, Qualcomm Hexagon, and Apple ANE are not yet supported
    by llama.cpp in mainstream builds.  Apple Silicon users get equivalent
    acceleration through Metal (step 3 above).
    """
    override = os.environ.get("GREPXCEL_GPU_LAYERS")
    if override is not None:
        try:
            return int(override), None, None
        except ValueError:
            pass

    # NVIDIA CUDA (works on Linux, Windows, and WSL2 with NVIDIA WSL2 driver)
    try:
        proc = subprocess.run(  # nosec — B603/B607: fixed argv, no shell, no user input
            ["nvidia-smi", "--query-gpu=name", "--format=csv,noheader"],
            capture_output=True, text=True, timeout=5,
        )
        if proc.returncode == 0:
            gpu_name = proc.stdout.strip().splitlines()[0]
            return -1, f"NVIDIA GPU ({gpu_name}) — CUDA", _CUDA_INSTALL
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    # Apple Silicon Metal (macOS only — ANE is bypassed; Metal is the fast path)
    if platform.system() == "Darwin":
        import struct
        if struct.calcsize("P") == 8 and platform.machine() == "arm64":
            return -1, "Apple Silicon — Metal", _METAL_INSTALL
        # Intel Mac: Metal exists but GPU offload yields little benefit
        return 0, None, None

    # Vulkan (any GPU on Windows/Linux — AMD, Intel, NVIDIA without CUDA driver)
    try:
        proc = subprocess.run(  # nosec — B603/B607: fixed argv, no shell, no user input
            ["vulkaninfo", "--summary"],
            capture_output=True, text=True, timeout=5,
        )
        if proc.returncode == 0 and "GPU" in proc.stdout:
            for line in proc.stdout.splitlines():
                if "deviceName" in line:
                    gpu_name = line.split("=", 1)[-1].strip()
                    return -1, f"Vulkan GPU ({gpu_name})", _VULKAN_INSTALL
            return -1, "Vulkan GPU", _VULKAN_INSTALL
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        pass

    # AMD XDNA NPU (Ryzen AI — Phoenix / Hawk Point / Strix Point)
    # Present in Ryzen 7040, 8040, and AI 300 series. Not accessible from
    # llama-cpp-python; requires ONNX Runtime + VitisAI EP on Windows.
    # Detected here only to inform the user rather than silently falling back
    # to CPU with no explanation.
    # Works on Windows-native (platform.processor()), Linux, and WSL2
    # (/proc/cpuinfo fallback when platform.processor() returns empty).
    try:
        import re
        _cpu = platform.processor()
        if not _cpu:  # Linux / WSL2 — platform.processor() is often empty
            try:
                with open('/proc/cpuinfo', encoding='utf-8') as _f:
                    for _line in _f:
                        if 'model name' in _line.lower():
                            _cpu = _line.split(':', 1)[-1].strip()
                            break
            except OSError:
                pass
        if _cpu and 'ryzen' in _cpu.lower() and re.search(
            r'[78][04]\d{2}|Ryzen AI', _cpu, re.IGNORECASE
        ):
            print(
                '  NPU detected: AMD XDNA (Ryzen AI) — not yet supported by '
                'llama-cpp-python.\n'
                '  Inference will run on CPU.  For faster results use:\n'
                '    grepxcel draft --backend claude  (or --backend gemini)',
                file=sys.stderr,
            )
    except Exception:  # noqa: BLE001  # nosec B110 — best-effort info message; never crash inference
        pass

    return 0, None, None


# ── Backend protocol ──────────────────────────────────────────────────────────

@dataclass
class CostRecord:
    """Token usage and USD cost for one API call.

    For metered APIs (Claude, Gemini) the *_cost_usd fields carry the dollar
    cost. For quota-based services (GitHub Models, included with a
    subscription) the dollar cost is 0 and consumption is reported instead via
    the optional rate-limit fields below (remaining / limit for requests and
    tokens, as returned in the provider's response headers)."""
    model:        str
    input_tokens: int
    output_tokens: int
    input_cost_usd:  float
    output_cost_usd: float
    # Optional quota signal (GitHub Models) — None for metered backends.
    rate_remaining_requests: int | None = None
    rate_limit_requests:     int | None = None
    rate_remaining_tokens:   int | None = None
    rate_limit_tokens:       int | None = None

    @property
    def total_cost_usd(self) -> float:
        return self.input_cost_usd + self.output_cost_usd

    def __str__(self) -> str:
        base = (
            f'{self.model}  '
            f'in={self.input_tokens} out={self.output_tokens}  '
            f'cost=${self.total_cost_usd:.6f} '
            f'(in=${self.input_cost_usd:.6f} + out=${self.output_cost_usd:.6f})'
        )
        if self.rate_remaining_requests is not None:
            base += (
                f'  quota: {self.rate_remaining_requests}/{self.rate_limit_requests} req, '
                f'{self.rate_remaining_tokens}/{self.rate_limit_tokens} tok remaining'
            )
        return base


@runtime_checkable
class LLMBackend(Protocol):
    """Minimal interface every inference backend must satisfy.

    Implement this to add new backends (Claude API, OpenAI, …) without
    touching PatternDrafter.  Pass an instance via PatternDrafter(backend=…).
    """

    def chat(self, system: str, user: str) -> str:
        """Send system + user messages; return the model's text response."""
        ...


# ── Local LLM client ──────────────────────────────────────────────────────────

class LlamaCppClient:
    """
    Runs inference entirely in-process using llama-cpp-python.
    The model is loaded lazily on the first chat() call and stays loaded
    for the lifetime of this instance — making it safe to reuse across
    multiple requests in a web service.
    """

    def __init__(self, model_path: str):
        self.model_path = model_path
        self._llm       = None

    def _load(self) -> None:
        if self._llm is not None:
            return
        try:
            from llama_cpp import Llama
        except ImportError:
            print(
                "Error: the local draft backend needs the 'suggest' extra "
                "(llama-cpp-python + huggingface_hub), which isn't installed.\n"
                "Fix:   pip install 'grepxcel[suggest]'\n"
                "       (or: python3 scripts/install_llm_deps.py  — autodetects GPU)\n"
                "Or skip the local model with a cloud backend: "
                "grepxcel draft --backend claude ...",
                file=sys.stderr,
            )
            sys.exit(1)

        n_gpu_layers, backend, install_hint = _detect_gpu()

        if backend:
            print(f"  Hardware: {backend}", file=sys.stderr)
            layers_label = "all layers" if n_gpu_layers == -1 else f"{n_gpu_layers} layers"
            print(f"  GPU offload: {layers_label}", file=sys.stderr)
            print(
                f"  Note: GPU acceleration requires a GPU-compiled build of llama-cpp-python.\n"
                f"  If inference seems slow, reinstall with:\n"
                f"    {install_hint}",
                file=sys.stderr,
            )
        else:
            print("  Hardware: CPU (no GPU detected — set GREPXCEL_GPU_LAYERS to override)",
                  file=sys.stderr)

        kwargs: dict = dict(
            model_path   = self.model_path,
            n_ctx        = 4096,
            n_threads    = os.cpu_count() or 4,
            n_gpu_layers = n_gpu_layers,
            verbose      = False,
        )
        if MODEL_CHAT_FORMAT is not None:
            kwargs['chat_format'] = MODEL_CHAT_FORMAT
        self._llm = Llama(**kwargs)

    def chat(self, system: str, user: str) -> str:
        self._load()
        response = self._llm.create_chat_completion(
            messages=[
                {"role": "system", "content": system},
                {"role": "user",   "content": user},
            ],
            temperature = 0.1,
            max_tokens  = 2048,
        )
        return response["choices"][0]["message"]["content"]


# ── Claude API backend ────────────────────────────────────────────────────────

class ClaudeBackend:
    """Sends inference requests to the Claude API (anthropic SDK).

    Requires ANTHROPIC_API_KEY to be set in the environment.
    Only the Excel structure description (column types, sample values, labels)
    is transmitted — the raw file bytes never leave the machine.

    Pricing (USD per 1M tokens, as of 2025-05):
      claude-haiku-4-5:   input $0.80   output $4.00
      claude-sonnet-4-5:  input $3.00   output $15.00
      claude-opus-4-5:    input $15.00  output $75.00
    """

    # USD per 1M tokens (source: platform.claude.com/docs/en/about-claude/models/overview)
    _PRICING: dict[str, tuple[float, float]] = {
        # Current models
        'claude-opus-4-8':              (5.00,  25.00),
        'claude-haiku-4-5-20251001':    (1.00,   5.00),
        'claude-haiku-4-5':             (1.00,   5.00),
        'claude-sonnet-4-6':            (3.00,  15.00),
        # Legacy models still available
        'claude-sonnet-4-5-20250929':   (3.00,  15.00),
        'claude-sonnet-4-5':            (3.00,  15.00),
        'claude-opus-4-7':              (5.00,  25.00),
        'claude-opus-4-6':              (5.00,  25.00),
        'claude-opus-4-5-20251101':     (5.00,  25.00),
        'claude-opus-4-5':              (5.00,  25.00),
        'claude-opus-4-1-20250805':    (15.00,  75.00),
        'claude-opus-4-1':             (15.00,  75.00),
    }

    def __init__(self, model: str = 'claude-haiku-4-5-20251001'):
        self._model     = model
        self._last_cost: CostRecord | None = None

    def last_cost(self) -> CostRecord | None:
        return self._last_cost

    def chat(self, system: str, user: str) -> str:
        try:
            import anthropic
        except ImportError:
            print(
                "Error: 'anthropic' package is not installed.\n"
                "Fix:   pip install anthropic",
                file=sys.stderr,
            )
            sys.exit(1)
        # Honor a corporate CA bundle / OS trust store behind a TLS-inspection
        # proxy (httpx ignores REQUESTS_CA_BUNDLE on its own). See proxy_support.
        from .proxy_support import make_httpx_client
        _http = make_httpx_client()
        _kw = {'http_client': _http} if _http is not None else {}
        client = anthropic.Anthropic(**_kw)  # reads ANTHROPIC_API_KEY from env
        msg = client.messages.create(
            model=self._model,
            max_tokens=2048,
            system=system,
            messages=[{'role': 'user', 'content': user}],
        )
        in_tok  = msg.usage.input_tokens
        out_tok = msg.usage.output_tokens
        in_p, out_p = self._PRICING.get(self._model, (0.0, 0.0))
        self._last_cost = CostRecord(
            model=self._model,
            input_tokens=in_tok,
            output_tokens=out_tok,
            input_cost_usd=in_tok  * in_p  / 1_000_000,
            output_cost_usd=out_tok * out_p / 1_000_000,
        )
        return msg.content[0].text


# ── Gemini API backend ────────────────────────────────────────────────────────

# Feature flag: the Gemini backend is implemented but DISABLED — planned for a
# future release. The full implementation below is kept intact; flip this to True
# (and re-enable the skipped Gemini tests) to ship it.
_GEMINI_ENABLED = False


class GeminiUnavailableError(RuntimeError):
    """Raised when the (disabled) Gemini backend is invoked programmatically."""


class GeminiBackend:
    """Sends inference requests to the Google Gemini API (google-genai SDK).

    DISABLED — planned for a future release. chat() refuses to run while
    _GEMINI_ENABLED is False; the implementation is preserved for that release.

    Requires GOOGLE_API_KEY to be set in the environment.
    Only the Excel structure description is transmitted — raw file bytes
    never leave the machine.

    Pricing (USD per 1M tokens, as of 2025-05, prompts ≤200K tokens):
      gemini-2.0-flash:       input $0.10   output $0.40
      gemini-2.5-flash:       input $0.15   output $0.60  (non-thinking)
      gemini-2.5-pro:         input $1.25   output $10.00 (≤200K)
    """

    _PRICING: dict[str, tuple[float, float]] = {
        'gemini-2.0-flash':          (0.10,  0.40),
        'gemini-2.0-flash-001':      (0.10,  0.40),
        'gemini-2.5-flash-preview-05-20': (0.15, 0.60),
        'gemini-2.5-flash':          (0.15,  0.60),
        'gemini-2.5-pro-preview-05-06':   (1.25, 10.00),
        'gemini-2.5-pro':            (1.25, 10.00),
    }

    def __init__(self, model: str = 'gemini-2.0-flash'):
        self._model     = model
        self._last_cost: CostRecord | None = None

    def last_cost(self) -> CostRecord | None:
        return self._last_cost

    def chat(self, system: str, user: str) -> str:
        if not _GEMINI_ENABLED:
            raise GeminiUnavailableError(
                'The Gemini backend is planned for a future release and is not '
                'yet available. Use the local or Claude backend instead.'
            )
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            print(
                "Error: 'google-genai' package is not installed.\n"
                "Fix:   pip install google-genai",
                file=sys.stderr,
            )
            sys.exit(1)
        # NOTE: when this backend is enabled, wire corporate-proxy CA trust here
        # too — the google-genai SDK takes an http_options transport; adapt
        # proxy_support.make_httpx_client() to it (see proxy_support).
        client = genai.Client()  # reads GOOGLE_API_KEY from env
        response = client.models.generate_content(
            model=self._model,
            contents=user,
            config=types.GenerateContentConfig(
                system_instruction=system,
                max_output_tokens=2048,
                temperature=0.1,
            ),
        )
        in_tok  = response.usage_metadata.prompt_token_count or 0
        out_tok = response.usage_metadata.candidates_token_count or 0
        in_p, out_p = self._PRICING.get(self._model, (0.0, 0.0))
        self._last_cost = CostRecord(
            model=self._model,
            input_tokens=in_tok,
            output_tokens=out_tok,
            input_cost_usd=in_tok  * in_p  / 1_000_000,
            output_cost_usd=out_tok * out_p / 1_000_000,
        )
        return response.text


# ── GitHub Models backend ─────────────────────────────────────────────────────

class GitHubModelsBackend:
    """Sends inference requests to GitHub Models (OpenAI-compatible endpoint).

    Access is included with a GitHub account / Copilot subscription, so there
    is no per-token dollar cost — consumption is governed by rate limits
    instead. Those limits (requests + tokens, with the remaining amounts) are
    returned in x-ratelimit-* response headers and captured into the
    CostRecord so usage stays visible.

    Requires GITHUB_TOKEN in the environment, with the 'Models: read'
    fine-grained permission. Model ids are namespaced, e.g. 'openai/gpt-4o',
    'meta/llama-3.3-70b-instruct', 'deepseek/deepseek-v3-0324'.
    Only the Excel structure description is transmitted — raw bytes never leave
    the machine.
    """

    ENDPOINT = 'https://models.github.ai/inference'

    def __init__(self, model: str = 'openai/gpt-4o-mini'):
        self._model     = model
        self._last_cost: CostRecord | None = None

    def last_cost(self) -> CostRecord | None:
        return self._last_cost

    @staticmethod
    def _int_header(headers, name: str) -> int | None:
        try:
            v = headers.get(name)
            return int(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    def chat(self, system: str, user: str) -> str:
        try:
            from openai import OpenAI
        except ImportError:
            print(
                "Error: the GitHub Models backend needs the 'openai' package.\n"
                "Fix:   pip install openai",
                file=sys.stderr,
            )
            sys.exit(1)
        token = os.environ.get('GITHUB_TOKEN')
        if not token:
            raise RuntimeError(
                'GITHUB_TOKEN is not set. Create a fine-grained token with the '
                "'Models: read' permission at https://github.com/settings/tokens "
                'and add GITHUB_TOKEN=... to your environment or .env file.'
            )
        # Corporate-proxy CA trust (see proxy_support / ClaudeBackend above).
        from .proxy_support import make_httpx_client
        _http = make_httpx_client()
        _kw = {'http_client': _http} if _http is not None else {}
        client = OpenAI(base_url=self.ENDPOINT, api_key=token, **_kw)
        raw = client.chat.completions.with_raw_response.create(
            model=self._model,
            messages=[
                {'role': 'system', 'content': system},
                {'role': 'user',   'content': user},
            ],
            temperature=0.1,
            max_tokens=2048,
        )
        completion = raw.parse()
        usage = completion.usage
        in_tok  = getattr(usage, 'prompt_tokens', 0) or 0
        out_tok = getattr(usage, 'completion_tokens', 0) or 0
        h = raw.headers
        self._last_cost = CostRecord(
            model=self._model,
            input_tokens=in_tok,
            output_tokens=out_tok,
            input_cost_usd=0.0,   # included in the subscription quota
            output_cost_usd=0.0,
            rate_remaining_requests=self._int_header(h, 'x-ratelimit-remaining-requests'),
            rate_limit_requests=self._int_header(h, 'x-ratelimit-limit-requests'),
            rate_remaining_tokens=self._int_header(h, 'x-ratelimit-remaining-tokens'),
            rate_limit_tokens=self._int_header(h, 'x-ratelimit-limit-tokens'),
        )
        return completion.choices[0].message.content


# ── OpenAI-compatible server backend ──────────────────────────────────────

class OpenAICompatBackend:
    """Sends inference to any OpenAI-compatible API server.

    Works with LM Studio, Ollama, vLLM, text-generation-inference, or any
    server exposing /v1/chat/completions.  Data stays local unless the user
    explicitly points at a remote URL.
    """

    def __init__(
        self,
        base_url: str = 'http://localhost:1234/v1',
        model: str | None = None,
        api_key: str = 'not-needed',
    ):
        self._base_url = base_url
        self._model    = model
        self._api_key  = api_key
        self._last_cost: CostRecord | None = None

    def last_cost(self) -> CostRecord | None:
        return self._last_cost

    def chat(self, system: str, user: str) -> str:
        try:
            from openai import OpenAI
        except ImportError:
            print(
                "Error: the server backend needs the 'openai' package.\n"
                "Fix:   pip install openai",
                file=sys.stderr,
            )
            sys.exit(1)
        client = OpenAI(base_url=self._base_url, api_key=self._api_key)
        model = self._model or self._resolve_model(client)
        completion = client.chat.completions.create(
            model=model,
            messages=[
                {'role': 'system', 'content': system},
                {'role': 'user',   'content': user},
            ],
            temperature=0.1,
            max_tokens=2048,
        )
        usage = completion.usage
        in_tok  = getattr(usage, 'prompt_tokens', 0) or 0
        out_tok = getattr(usage, 'completion_tokens', 0) or 0
        self._last_cost = CostRecord(
            model=model,
            input_tokens=in_tok,
            output_tokens=out_tok,
            input_cost_usd=0.0,
            output_cost_usd=0.0,
        )
        return completion.choices[0].message.content

    @staticmethod
    def _resolve_model(client) -> str:
        models = client.models.list()
        if models.data:
            return models.data[0].id
        raise RuntimeError(
            'No models loaded on the server. Load a model in LM Studio / '
            'Ollama first, then retry.'
        )


# ── Pattern writer ────────────────────────────────────────────────────────────

_TABLE_ROW_PREFIXES = ('HEADER:', 'DATA:', 'FOOTER:', 'SPLITTER:', 'SKIP_IF')


class PatternWriter:
    """Parses LLM pipe-delimited text and writes a pattern xlsx."""

    def write(self, llm_text: str, output_path: str) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title  = 'Pattern'
        row_num   = 1

        for line in llm_text.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith('```'):
                continue

            if stripped.upper() in ('START:', 'END:'):
                ws.cell(row=row_num, column=1, value=stripped.upper())
                row_num += 1
                continue

            is_table_row = stripped.startswith('|') or any(
                stripped.upper().startswith(p) for p in _TABLE_ROW_PREFIXES
            )
            if is_table_row:
                # Strip leading '|'; remaining parts map to columns B, C, D, …
                body  = stripped.lstrip('|')
                parts = [p.strip() for p in body.split('|')]
                for rel_i, val in enumerate(parts):
                    if val:
                        ws.cell(row=row_num, column=rel_i + 2, value=val)
                row_num += 1
                continue

            if '|' in stripped:
                parts = [p.strip() for p in stripped.split('|')]
                for col_num, val in enumerate(parts, start=1):
                    if val:
                        ws.cell(row=row_num, column=col_num, value=val)
                row_num += 1
                continue

            # Bare keyword (e.g. 'table:*')
            parts = stripped.split(None, 1)
            ws.cell(row=row_num, column=1, value=parts[0])
            if len(parts) > 1:
                ws.cell(row=row_num, column=2, value=parts[1])
            row_num += 1

        wb.save(output_path)
        from .pattern_colors import colorize_pattern_file
        colorize_pattern_file(output_path)


# ── Validation helpers ────────────────────────────────────────────────────────

def _failed_path(output_path: str) -> str:
    """Derive a _FAILED.txt path alongside the intended output file."""
    p = Path(output_path)
    return str(p.parent / f'{p.stem}_FAILED.txt')


def _write_failed_draft(llm_text: str, error: str, output_path: str) -> None:
    path = _failed_path(output_path)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(f'# Draft validation error\n# {error}\n\n')
        f.write(llm_text)


def _validate_draft(xlsx_path: str) -> None:
    """Parse the draft xlsx with the engine's own PatternParser.

    Raises PatternError or SecurityError if the LLM output is structurally invalid.
    """
    PatternParser().parse(xlsx_path)


# ── Local-backend dependency check (speak clearly to the user) ──────────────────

def _missing_local_deps() -> list[str]:
    """Return the pip names of any local-model deps that aren't importable.

    Uses find_spec so it never imports the heavy packages — just checks presence.
    """
    import importlib.util
    missing = []
    for module_name, pip_name in (
        ('huggingface_hub', 'huggingface_hub'),
        ('llama_cpp',       'llama-cpp-python'),
    ):
        if importlib.util.find_spec(module_name) is None:
            missing.append(pip_name)
    return missing


def _print_local_backend_help(missing: list[str]) -> None:
    """One clear message: what's missing, that draft is optional, and the options."""
    print(
        f"The local 'draft' model isn't installed (missing: {', '.join(missing)}).\n"
        f"\n"
        f"'draft' is an optional feature — 'extract' and 'docs' work without it.\n"
        f"To use the local model (runs fully offline), install it once:\n"
        f"  pip install 'grepxcel[suggest]'\n"
        f"  python3 scripts/install_llm_deps.py    # alternative — autodetects GPU\n"
        f"\n"
        f"Or skip the local model with a cloud backend (no large download):\n"
        f"  grepxcel draft --backend claude ...    # needs ANTHROPIC_API_KEY",
        file=sys.stderr,
    )


# ── Orchestrator ──────────────────────────────────────────────────────────────

class PatternDrafter:
    """Analyse → run LLM backend → validate → write pattern xlsx.

    The inference backend is injectable via the `backend` parameter.
    When omitted the local GGUF model (LlamaCppClient) is used, which
    requires llama-cpp-python and downloads the model on first run.
    Pass any object that satisfies LLMBackend to use a different backend.
    """

    def __init__(
        self,
        input_path: str,
        output_path: str,
        sheet=None,
        max_file_mb: float = 5.0,
        max_uncompressed_mb: float = 50.0,
        verbose: bool = False,
        dry_run: bool = False,
        backend: LLMBackend | None = None,
        allow_unverified: bool = False,
    ):
        self.input_path          = input_path
        self.output_path         = output_path
        self.sheet               = sheet
        self.max_file_mb         = max_file_mb
        self.max_uncompressed_mb = max_uncompressed_mb
        self.verbose             = verbose
        self.dry_run             = dry_run
        self.backend             = backend
        self.allow_unverified    = allow_unverified

    def run(self) -> int:
        """Run the full pipeline. Returns exit code (0 = success, 1 = error)."""
        # Fail fast: if we'll run the LOCAL model (no cloud backend, not a
        # dry-run), confirm its optional deps are installed BEFORE doing any
        # analysis — so the user gets clear guidance up front, not after a wall
        # of output. dry-run and cloud backends don't need these deps.
        if self.backend is None and not self.dry_run:
            missing = _missing_local_deps()
            if missing:
                _print_local_backend_help(missing)
                return 1

        # 1. Analyse Excel structure
        print('Analysing Excel structure...', file=sys.stderr)
        try:
            analysis = ExcelAnalyzer(
                self.input_path, self.sheet,
                self.max_file_mb, self.max_uncompressed_mb,
            ).analyse()
        except SecurityError as exc:
            print(f'Security error: {exc}', file=sys.stderr)
            return 1
        except ValueError as exc:
            # Bad --sheet selection (unknown name / out-of-range index).
            print(f'Error: {exc}', file=sys.stderr)
            return 1

        if self.verbose or self.dry_run:
            print('\n── Excel analysis ──────────────────────────────', file=sys.stderr)
            print(analysis, file=sys.stderr)
            print('────────────────────────────────────────────────\n', file=sys.stderr)

        if self.dry_run:
            print('[dry-run] Model inference skipped.', file=sys.stderr)
            return 0

        # 2. Resolve backend and run inference
        user_prompt = _USER_PROMPT_TEMPLATE.format(analysis=analysis)
        if self.backend is not None:
            print('Running inference...', file=sys.stderr)
            llm_text = self.backend.chat(_SYSTEM_PROMPT, user_prompt)
            # Surface usage / consumption when the backend tracks it.
            cost = self.backend.last_cost() if hasattr(self.backend, 'last_cost') else None
            if cost is not None:
                if cost.rate_remaining_requests is not None:
                    print(
                        f'  Usage: {cost.input_tokens} in + {cost.output_tokens} out tokens '
                        f'(included in GitHub subscription quota)\n'
                        f'  Quota remaining: {cost.rate_remaining_requests}/'
                        f'{cost.rate_limit_requests} requests, '
                        f'{cost.rate_remaining_tokens}/{cost.rate_limit_tokens} tokens',
                        file=sys.stderr,
                    )
                else:
                    print(
                        f'  Usage: {cost.input_tokens} in + {cost.output_tokens} out tokens '
                        f'(cost ${cost.total_cost_usd:.6f})',
                        file=sys.stderr,
                    )
        else:
            # Default: local GGUF model via llama-cpp-python
            model_path = ModelManager(
                allow_unverified=self.allow_unverified,
            ).ensure_ready(verbose=self.verbose)
            print('Running local model inference...', file=sys.stderr)
            llm_text   = LlamaCppClient(str(model_path)).chat(_SYSTEM_PROMPT, user_prompt)

        # 4. Text preview → stdout (pipe-friendly)
        print(llm_text)

        # 5. Write to a temp xlsx, validate, then move to final path
        output_dir = os.path.dirname(os.path.abspath(self.output_path))
        os.makedirs(output_dir, exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=output_dir)
        os.close(fd)
        try:
            PatternWriter().write(llm_text, tmp_path)
            _validate_draft(tmp_path)
        except (PatternError, SecurityError) as exc:
            os.unlink(tmp_path)
            _write_failed_draft(llm_text, str(exc), self.output_path)
            print(
                f'\n[!] Draft validation failed: {exc}\n'
                f'    Raw LLM output saved to: {_failed_path(self.output_path)}',
                file=sys.stderr,
            )
            return 1

        os.replace(tmp_path, self.output_path)
        print(f'\nDraft pattern written to: {self.output_path}', file=sys.stderr)
        print(
            f'Open it in Excel/LibreOffice, refine the regexes, then run:\n'
            f'  grepxcel extract -p {self.output_path} your_data.xlsx',
            file=sys.stderr,
        )
        return 0
