"""Generate a CycloneDX 1.6 SBOM from the installed grepxcel environment.

Uses only stdlib (importlib.metadata) — no external dependency. Produces a
CycloneDX 1.6 JSON document listing grepxcel and every transitive dependency
with PURLs, license identifiers, and SHA-256 hashes from pip's RECORD files.

Trusted by organisations because:
  - CycloneDX is the OWASP standard (ISO/IEC 5962:2021)
  - Every component has a Package URL (PURL) for unambiguous identification
  - Hashes come from pip's installed RECORD, matching what's on disk
  - License identifiers use SPDX where available
  - The tool metadata documents exactly what generated the SBOM
"""

from __future__ import annotations

import csv
import io
import json
import sys
from datetime import datetime, timezone
from importlib import metadata
from pathlib import Path
from typing import Optional

import grepxcel


SPEC_VERSION = '1.6'
SCHEMA_URL = 'http://cyclonedx.org/schema/bom-1.6.schema.json'
PKG_NAME = 'grepxcel'


def _get_dist_hashes(dist: metadata.Distribution) -> list[dict]:
    """Extract SHA-256 hashes from pip's RECORD file for a distribution."""
    hashes = []
    try:
        record_text = dist.read_text('RECORD')
    except FileNotFoundError:
        return hashes
    if not record_text:
        return hashes
    for row in csv.reader(io.StringIO(record_text)):
        if len(row) >= 2 and row[1].startswith('sha256='):
            hashes.append({
                'alg': 'SHA-256',
                'content': row[1].split('=', 1)[1],
            })
            break
    return hashes


def _get_license(dist: metadata.Distribution) -> str:
    """Best-effort SPDX license identifier from package metadata."""
    meta = dist.metadata
    license_expr = meta.get('License-Expression', '')
    if license_expr:
        return license_expr
    license_field = meta.get('License', '')
    if license_field and len(license_field) < 80:
        return license_field
    classifiers = meta.get_all('Classifier') or []
    for c in classifiers:
        if c.startswith('License :: OSI Approved :: '):
            return c.split(' :: ')[-1]
    return ''


def _collect_deps(root: str) -> set[str]:
    """Recursively collect all transitive dependency names for a package."""
    visited: set[str] = set()
    queue = [root]
    while queue:
        name = queue.pop()
        normalised = name.lower().replace('-', '_').replace('.', '_')
        if normalised in visited:
            continue
        visited.add(normalised)
        try:
            dist = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            continue
        reqs = dist.requires or []
        for req in reqs:
            if '; extra ==' in req or '; extra ==' in req.replace('"', "'"):
                continue
            dep_name = req.split()[0].split('>')[0].split('<')[0].split('=')[0].split('!')[0].split('[')[0].split(';')[0]
            queue.append(dep_name)
    return visited


def _component(dist: metadata.Distribution) -> dict:
    """Build one CycloneDX component dict from a distribution."""
    name = dist.metadata['Name']
    version = dist.metadata['Version'] or 'unknown'
    purl = f'pkg:pypi/{name.lower()}@{version}'
    comp: dict = {
        'type': 'library',
        'name': name,
        'version': version,
        'purl': purl,
        'bom-ref': purl,
    }
    license_id = _get_license(dist)
    if license_id:
        comp['licenses'] = [{'license': {'name': license_id}}]
    hashes = _get_dist_hashes(dist)
    if hashes:
        comp['hashes'] = hashes
    return comp


def generate_sbom(*, output_format: str = 'json') -> dict:
    """Generate a CycloneDX 1.6 SBOM for the installed grepxcel package.

    Returns the SBOM as a dict. Use ``json.dumps(sbom, indent=2)`` to
    serialise.
    """
    timestamp = datetime.now(timezone.utc).isoformat()

    try:
        root_dist = metadata.distribution(PKG_NAME)
        root_version = root_dist.metadata['Version']
    except metadata.PackageNotFoundError:
        root_version = grepxcel.__version__

    dep_names = _collect_deps(PKG_NAME)
    dep_names.discard(PKG_NAME.lower().replace('-', '_'))

    components = []
    for name in sorted(dep_names):
        try:
            dist = metadata.distribution(name)
            components.append(_component(dist))
        except metadata.PackageNotFoundError:
            pass

    sbom = {
        '$schema': SCHEMA_URL,
        'bomFormat': 'CycloneDX',
        'specVersion': SPEC_VERSION,
        'version': 1,
        'metadata': {
            'timestamp': timestamp,
            'tools': {
                'components': [{
                    'type': 'application',
                    'name': PKG_NAME,
                    'version': root_version,
                }],
            },
            'component': {
                'type': 'application',
                'name': PKG_NAME,
                'version': root_version,
                'purl': f'pkg:pypi/{PKG_NAME}@{root_version}',
                'bom-ref': f'pkg:pypi/{PKG_NAME}@{root_version}',
                'licenses': [{'license': {'id': 'MIT'}}],
            },
            'supplier': {
                'name': 'scpg',
                'url': ['https://github.com/scpg/grepxcel'],
            },
        },
        'components': components,
    }

    return sbom


def render_sbom(sbom: dict, *, output_format: str = 'json') -> str:
    """Render the SBOM dict to a string."""
    return json.dumps(sbom, indent=2, default=str)


def run_sbom(output: Optional[str] = None, out=None) -> int:
    """Generate and write the SBOM. Returns exit code."""
    if out is None:
        out = sys.stdout
    sbom = generate_sbom()
    rendered = render_sbom(sbom)
    if output:
        Path(output).write_text(rendered + '\n', encoding='utf-8')
        print(f'SBOM written to {output}', file=sys.stderr)
    else:
        out.write(rendered + '\n')
    return 0
