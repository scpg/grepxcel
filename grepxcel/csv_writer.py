"""Convert nested extraction output to CSV.

Single-table default: when the pattern has multiple table: blocks, the first
table is exported and a warning is emitted.  Use ``--csv-table N`` to select
a specific table.  Multiple tables cannot be losslessly represented as a flat
CSV — use --format nested (JSON) if you need all tables.

Only table ``data`` rows are exported; per-table ``header``/``footer`` fields
(e.g. subtotals) do not fit a flat row layout and are omitted. Use
``--format nested`` (JSON) or :func:`grepxcel.extract_df` for those.

Every string cell is passed through :func:`grepxcel.utils.neutralize_formula`
so a value extracted from an untrusted file cannot become an executable formula
when the CSV is opened in a spreadsheet app (CWE-1236).
"""

from __future__ import annotations

import csv
import io
from typing import Any

from .models import TableInstruction
from .pattern_parser import PatternParser
from .utils import flatten_nested, neutralize_formula


def count_table_instructions(pattern_file: str) -> int:
    _, _, seq = PatternParser().parse(pattern_file)
    return sum(1 for i in seq if isinstance(i, TableInstruction))


def nested_to_csv(
    result: dict,
    delimiter: str = ',',
    mode: str = 'extended',
    table_idx: int | None = None,
) -> tuple[str, list[str]]:
    """Convert a nested extraction result to a CSV string.

    Returns ``(csv_text, dropped)`` where ``dropped`` is a list of
    human-readable strings describing table fields that were omitted because
    they cannot be represented in flat CSV (header/footer rows).  The caller
    is responsible for surfacing these as warnings.

    - Scalar fields are flattened with dot-notation columns.
    - Table data rows become one CSV row each.
    - When both exist and ``mode='extended'`` (default), scalar values are
      denormalized (repeated) on every row.  ``mode='table'`` exports only
      the table columns, omitting all scalars.
    - ``table_idx`` (0-based) selects which table to export when the result
      has multiple list-valued keys; ``None`` uses the first one found.
    - ``_source`` metadata and the ``_meta`` block are excluded.
    - ``lineterminator='\\r\\n'`` per RFC 4180 for Windows Excel compatibility.
    """
    scalars: dict[str, Any] = {}
    table_rows: list[dict[str, Any]] = []
    dropped: list[str] = []

    if mode == 'extended':
        for key, value in result.items():
            if key == '_meta':
                continue
            if isinstance(value, dict):
                scalars.update(dict(flatten_nested(value, key)))
            elif not isinstance(value, list):
                scalars[key] = value

    table_keys = [k for k, v in result.items() if k != '_meta' and isinstance(v, list)]

    if table_keys:
        idx = 0 if table_idx is None else table_idx
        if idx < 0 or idx >= len(table_keys):
            n = len(table_keys)
            raise ValueError(
                f'--csv-table {idx + 1} is out of range: result has {n} '
                f'table(s) (valid: 1–{n})'
            )
        selected_key = table_keys[idx]
        instances = result[selected_key]
        for inst_idx, instance in enumerate(instances):
            if not isinstance(instance, dict):
                continue
            for data_row in instance.get('data', []):
                table_rows.append({f'{selected_key}.{k}': v for k, v in data_row.items()})
            suffix = f'[{inst_idx}]' if len(instances) > 1 else ''
            if instance.get('header'):
                dropped.append(f'{selected_key}{suffix}.header')
            if instance.get('footer'):
                dropped.append(f'{selected_key}{suffix}.footer')

    if not table_rows and not scalars:
        return '', dropped

    rows = [{**scalars, **row} for row in table_rows] if table_rows else [scalars]

    fieldnames: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for k in row:
            if k not in seen:
                seen.add(k)
                fieldnames.append(k)

    buf = io.StringIO()
    writer = csv.DictWriter(
        buf, fieldnames=fieldnames, lineterminator='\r\n', delimiter=delimiter,
    )
    writer.writeheader()
    for row in rows:
        writer.writerow({k: neutralize_formula(v) for k, v in row.items()})
    return buf.getvalue(), dropped
