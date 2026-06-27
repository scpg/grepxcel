"""Convert nested extraction output to CSV.

Single-table guard: CSV output is only supported when the pattern has at most
one table: block. Multiple tables cannot be losslessly represented as a flat
CSV — use --format nested (JSON) instead.

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


def nested_to_csv(result: dict) -> str:
    """Convert a nested extraction result to a CSV string.

    - Scalar fields are flattened with dot-notation columns.
    - Table data rows become one CSV row each.
    - When both exist, scalar values are denormalized (repeated) on every row.
    - ``_source`` metadata and the ``_meta`` block are excluded.
    """
    scalars: dict[str, Any] = {}
    table_rows: list[dict[str, Any]] = []

    for key, value in result.items():
        if key == '_meta':
            continue
        if isinstance(value, list):
            for instance in value:
                if not isinstance(instance, dict):
                    continue
                for data_row in instance.get('data', []):
                    # Qualify field names with the table key (e.g. txn.date)
                    table_rows.append({f'{key}.{k}': v for k, v in data_row.items()})
        elif isinstance(value, dict):
            scalars.update(dict(flatten_nested(value, key)))
        else:
            scalars[key] = value

    if not table_rows and not scalars:
        return ''

    if table_rows:
        rows = [{**scalars, **row} for row in table_rows]
    else:
        rows = [scalars]

    fieldnames: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for k in row:
            if k not in seen:
                seen.add(k)
                fieldnames.append(k)

    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, lineterminator='\n')
    writer.writeheader()
    for row in rows:
        writer.writerow({k: neutralize_formula(v) for k, v in row.items()})
    return buf.getvalue()
