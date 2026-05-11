"""
Template suggestion: analyse a target Excel file and use a locally-running
GGUF model (via llama-cpp-python) to produce a suggested grepxcel pattern file.

No data leaves the process during inference — the model runs entirely in-process.
The only network activity is the one-time model download (and daily update checks)
handled by ModelManager.
"""

import os
import sys

import openpyxl

from .model_manager import MODEL_CHAT_FORMAT, ModelManager
from .security import SecurityError, validate_file
from .utils import infer_cell_type, is_empty


# ── Prompts ───────────────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """\
You are an expert at generating grepxcel pattern files.

A grepxcel pattern file is an Excel workbook that describes how to extract structured data
from another Excel file. Each row uses up to 4 columns (A, B, C, D).
In your output, separate columns with ' | ' (space-pipe-space).

─── SECTIONS (in order) ───────────────────────────────────────────────────────

1. Config rows (optional, before def: rows):
   config: | read.direction | LR          (or TD for top-down column scanning)
   config: | currency.sign  | €

2. Field definitions (required, at least one):
   def: | FieldName | type | regex

   Types:  string  integer  currency  date  datetime
   Regex:  Python re.fullmatch pattern. Use .* to match anything.

   Examples:
     def: | InvoiceNo  | string   | INV-\\d{4,8}
     def: | Amount     | currency | \\d+(\\.\\d{1,2})?
     def: | Qty        | integer  | \\d+
     def: | Label      | string   | .*
     def: | OrderDate  | date     |
   For date/datetime types, leave the regex column empty.

3. Extraction sequence between START: and END: (bare keywords, no pipes):

   For scattered key-value cells:
     cell:1 | FieldName      (extract the next non-empty cell)
     cell:1 | IGNORE         (skip the next non-empty cell)

   For repeating tables:
     table:*                 (bare keyword, no pipe, starts a table block)
       | HEADER:1 | ColA | ColB | ColC
       | DATA:*   | ColA | ColB | ColC
       | FOOTER:1 | ColA | ColB | ColC
     (Table template rows have a blank column A — start the line with ' | ')

   Row type suffixes:  :1  (exactly one)   :*  (one or more)   :N  (exactly N)
   Column keywords:  FieldName  IGNORE  EMPTY

─── OUTPUT RULES ──────────────────────────────────────────────────────────────

- Output ONLY the pattern rows. No explanations, no markdown fences, no comments.
- One row per line, columns separated by ' | '.
- Table template rows must start with ' | ' (blank column A).
- START: and END: are bare keywords with no pipe separator.

─── EXAMPLE OUTPUT ────────────────────────────────────────────────────────────

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

_USER_PROMPT_TEMPLATE = """\
Analyse the Excel structure below and produce a grepxcel pattern file that would extract its key data.

{analysis}

Generate the pattern file now:
"""


# ── Excel analyser ────────────────────────────────────────────────────────────

class ExcelAnalyzer:
    """Reads an xlsx file and returns a plain-text structural description."""

    _MAX_SAMPLE_ROWS = 5
    _MAX_KV_PAIRS    = 30

    def __init__(
        self,
        path: str,
        sheet=None,
        max_file_mb: float = 5.0,
        max_uncompressed_mb: float = 50.0,
    ):
        self.path               = path
        self.sheet              = sheet
        self.max_file_mb        = max_file_mb
        self.max_uncompressed_mb = max_uncompressed_mb

    def analyse(self) -> str:
        validate_file(self.path, self.max_file_mb, self.max_uncompressed_mb)
        wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
        try:
            ws = self._select_sheet(wb)
            return self._analyse_sheet(ws)
        finally:
            wb.close()

    def _select_sheet(self, wb):
        if self.sheet is None:
            return wb.active
        if isinstance(self.sheet, int):
            return wb.worksheets[self.sheet]
        return wb[self.sheet]

    def _analyse_sheet(self, ws) -> str:
        rows      = list(ws.iter_rows(values_only=True))
        col_count = max((len(r) for r in rows), default=0)
        lines     = [
            f"Sheet: '{ws.title}'",
            f"Dimensions: {len(rows)} rows × {col_count} columns",
        ]
        if not rows:
            return "Empty sheet — no data found."

        first_row_values = [v for v in rows[0] if v is not None]
        if len(first_row_values) >= 2:
            lines += self._describe_table_layout(rows)
        else:
            lines += self._describe_kv_layout(rows)

        return '\n'.join(lines)

    def _describe_table_layout(self, rows: list) -> list:
        headers           = list(rows[0])
        non_empty_headers = [(i, h) for i, h in enumerate(headers) if h is not None]
        data_rows         = rows[1:1 + self._MAX_SAMPLE_ROWS]

        lines = [
            '',
            'Layout: TABLE (first row appears to be column headers)',
            f'Columns ({len(non_empty_headers)}):',
        ]
        for col_idx, header in non_empty_headers:
            sample_vals = [r[col_idx] for r in data_rows if col_idx < len(r)]
            col_type    = infer_cell_type(sample_vals)
            samples     = [repr(v) for v in sample_vals if v is not None][:3]
            sample_str  = ', '.join(samples) if samples else '(no data)'
            lines.append(f"  - '{header}' [{col_type}]  samples: {sample_str}")

        data_row_count = len(rows) - 1
        if data_row_count > 0:
            lines.append(
                f'\nData rows: {data_row_count} total (sampled up to {self._MAX_SAMPLE_ROWS})'
            )
        return lines

    def _describe_kv_layout(self, rows: list) -> list:
        lines = [
            '',
            'Layout: KEY-VALUE (scattered cells, not a standard table)',
            'Cell pairs found:',
        ]
        count = 0
        for row in rows:
            if count >= self._MAX_KV_PAIRS:
                break
            for col_idx, val in enumerate(row):
                if val is None or is_empty(val):
                    continue
                next_val = row[col_idx + 1] if col_idx + 1 < len(row) else None
                col_type = infer_cell_type([next_val]) if next_val is not None else 'string'
                if next_val is not None:
                    lines.append(f"  - '{val}': {repr(next_val)} [{col_type}]")
                else:
                    lines.append(f"  - '{val}': (no adjacent value) [string]")
                count += 1
                if count >= self._MAX_KV_PAIRS:
                    break
        return lines


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
                "Error: 'llama-cpp-python' is not installed.\n"
                "Fix:   pip install llama-cpp-python",
                file=sys.stderr,
            )
            sys.exit(1)

        self._llm = Llama(
            model_path  = self.model_path,
            n_ctx       = 4096,
            n_threads   = os.cpu_count() or 4,
            n_gpu_layers= 0,       # CPU-only; set GREPXCEL_GPU_LAYERS to override
            verbose     = False,
            chat_format = MODEL_CHAT_FORMAT,
        )

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


# ── Pattern writer ────────────────────────────────────────────────────────────

_TABLE_ROW_PREFIXES = ('HEADER:', 'DATA:', 'FOOTER:', 'SPLITTER:')


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


# ── Orchestrator ──────────────────────────────────────────────────────────────

class PatternSuggester:
    """Analyse → ensure model → run LLM → write pattern xlsx."""

    def __init__(
        self,
        input_path: str,
        output_path: str,
        sheet=None,
        max_file_mb: float = 5.0,
        max_uncompressed_mb: float = 50.0,
        verbose: bool = False,
    ):
        self.input_path          = input_path
        self.output_path         = output_path
        self.sheet               = sheet
        self.max_file_mb         = max_file_mb
        self.max_uncompressed_mb = max_uncompressed_mb
        self.verbose             = verbose

    def run(self) -> int:
        """Run the full pipeline. Returns exit code (0 = success, 1 = error)."""
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

        if self.verbose:
            print('\n── Excel analysis ──────────────────────────────', file=sys.stderr)
            print(analysis, file=sys.stderr)
            print('────────────────────────────────────────────────\n', file=sys.stderr)

        # 2. Ensure the model is present and up to date
        model_path = ModelManager().ensure_ready(verbose=self.verbose)

        # 3. Run inference (entirely in-process — no network)
        print('Running local model inference...', file=sys.stderr)
        user_prompt = _USER_PROMPT_TEMPLATE.format(analysis=analysis)
        llm_text    = LlamaCppClient(str(model_path)).chat(_SYSTEM_PROMPT, user_prompt)

        # 4. Text preview → stdout (pipe-friendly)
        print(llm_text)

        # 5. Write xlsx
        PatternWriter().write(llm_text, self.output_path)
        print(f'\nPattern file written to: {self.output_path}', file=sys.stderr)
        print(
            f'Open it in Excel/LibreOffice, refine the regexes, then run:\n'
            f'  grepxcel extract -p {self.output_path} your_data.xlsx',
            file=sys.stderr,
        )
        return 0
