"""`grepxcel schema` — generate a JSON Schema from a pattern file.

The pattern already defines field names, types, nesting (dot notation), and
whether fields live in scalars or tables. This module translates that into a
standard JSON Schema (draft 2020-12) so consumers can validate extraction
output and understand the structure without reading the pattern.
"""
from __future__ import annotations

import json
import sys
from typing import Any

from .models import CellInstruction, TableInstruction
from .pattern_parser import PatternParser

_SCHEMA_DRAFT = 'https://json-schema.org/draft/2020-12/schema'

_TYPE_MAP: dict[str, dict[str, Any]] = {
    'string':     {'type': ['string', 'null']},
    'text':       {'type': ['string', 'null']},
    'integer':    {'type': ['integer', 'null']},
    'number':     {'type': ['number', 'null']},
    'float':      {'type': ['number', 'null']},
    'decimal':    {'type': ['number', 'null']},
    'currency':   {'type': ['number', 'null']},
    'percentage': {'type': ['number', 'null']},
    'boolean':    {'type': ['boolean', 'null']},
    'bool':       {'type': ['boolean', 'null']},
    'date':       {'type': ['string', 'null'], 'format': 'date-time'},
    'datetime':   {'type': ['string', 'null'], 'format': 'date-time'},
    'timestamp':  {'type': ['string', 'null'], 'format': 'date-time'},
    'time':       {'type': ['string', 'null']},
    'duration':   {'type': ['string', 'null']},
}


def _json_type(field_type: str) -> dict[str, Any]:
    return dict(_TYPE_MAP.get(field_type, {'type': ['string', 'null']}))


def _set_nested(d: dict, dotted_key: str, value: dict) -> None:
    parts = dotted_key.split('.')
    for part in parts[:-1]:
        child = d.setdefault(part, {'type': 'object', 'properties': {}})
        if 'properties' not in child:
            child['properties'] = {}
        d = child['properties']
    d[parts[-1]] = value


def _field_group(name: str) -> str:
    return name.split('.', 1)[0]


def _field_local(name: str) -> str:
    _, _, rest = name.partition('.')
    return rest if rest else name


def _source_schema() -> dict:
    return {
        'type': 'object',
        'properties': {
            'sheet': {'type': 'string'},
            'ref': {'type': 'string'},
        },
    }


def _collect_table_fields(instr: TableInstruction, defs: dict) -> dict:
    """Classify table template fields into header/data/footer buckets."""
    header_fields: list[str] = []
    data_fields: list[str] = []
    footer_fields: list[str] = []

    for trow in instr.rows:
        for col in trow.columns:
            if col.field in ('IGNORE', 'EMPTY'):
                continue
            fd = defs.get(col.field)
            if fd and fd.role == 'lbl':
                continue
            if trow.row_type == 'HEADER':
                if col.field not in header_fields:
                    header_fields.append(col.field)
            elif trow.row_type == 'DATA':
                if col.field not in data_fields:
                    data_fields.append(col.field)
            elif trow.row_type == 'FOOTER':
                if col.field not in footer_fields:
                    footer_fields.append(col.field)
    return {
        'header': header_fields,
        'data': data_fields,
        'footer': footer_fields,
    }


def _fields_to_schema(fields: list[str], defs: dict) -> dict:
    """Build a properties dict from a list of field names (using local names)."""
    props: dict = {}
    for field in fields:
        fd = defs.get(field)
        if not fd:
            continue
        local = _field_local(field)
        _set_nested(props, local, _json_type(fd.type))
    return props


def _table_group_name(data_fields: list[str], defs: dict, table_index: int) -> str:
    """Derive the output key for a table group, mirroring engine._table_group."""
    counts: dict[str, int] = {}
    for field in data_fields:
        fd = defs.get(field)
        if fd and fd.role == 'var' and '.' in field:
            g = _field_group(field)
            counts[g] = counts.get(g, 0) + 1
    if counts:
        return max(counts, key=lambda k: counts[k])
    return f'table_{table_index}'


def generate_schema(pattern_path: str) -> dict:
    """Generate a JSON Schema from a parsed pattern file.

    The schema describes the nested output format that ``grepxcel extract``
    produces, so extraction results can be validated with any JSON Schema
    library.
    """
    config, defs, sequence = PatternParser().parse(pattern_path)

    root: dict = {
        '$schema': _SCHEMA_DRAFT,
        'type': 'object',
        'properties': {},
        'additionalProperties': False,
    }
    props = root['properties']

    table_index = 0
    for instr in sequence:
        if isinstance(instr, CellInstruction):
            if instr.field in ('IGNORE', 'EMPTY'):
                continue
            fd = defs.get(instr.field)
            if not fd or fd.role == 'lbl':
                continue
            _set_nested(props, instr.field, _json_type(fd.type))

        elif isinstance(instr, TableInstruction):
            buckets = _collect_table_fields(instr, defs)
            group = _table_group_name(buckets['data'], defs, table_index)
            table_index += 1

            instance_props: dict = {
                '_source': _source_schema(),
            }

            if buckets['header']:
                header_props = _fields_to_schema(buckets['header'], defs)
                if header_props:
                    instance_props['header'] = {
                        'type': 'object',
                        'properties': header_props,
                    }

            if buckets['data']:
                data_props = _fields_to_schema(buckets['data'], defs)
                if data_props:
                    instance_props['data'] = {
                        'type': 'array',
                        'items': {
                            'type': 'object',
                            'properties': data_props,
                        },
                    }

            if buckets['footer']:
                footer_props = _fields_to_schema(buckets['footer'], defs)
                if footer_props:
                    instance_props['footer'] = {
                        'type': 'object',
                        'properties': footer_props,
                    }

            table_schema = {
                'type': 'array',
                'items': {
                    'type': 'object',
                    'properties': instance_props,
                },
            }
            props[group] = table_schema

    return root


def run_schema(paths: list[str], out=None) -> int:
    """Generate and print JSON Schema for each pattern file. Returns 0 on success."""
    out = out or sys.stdout
    for path in paths:
        try:
            schema = generate_schema(path)
        except FileNotFoundError:
            print(f'Error: file not found: {path}', file=sys.stderr)
            return 1
        except Exception as exc:
            print(f'Error: {path}: {exc}', file=sys.stderr)
            return 1
        json.dump(schema, out, indent=2)
        out.write('\n')
    return 0
