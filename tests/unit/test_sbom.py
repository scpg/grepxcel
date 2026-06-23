"""Tests for the SBOM generator (CycloneDX 1.6)."""

import io
import json

import pytest

from grepxcel.cli import main as cli_main
from grepxcel.sbom import generate_sbom, render_sbom, run_sbom, _collect_deps


class TestGenerateSbom:
    """The generate_sbom function produces a valid CycloneDX document."""

    def test_returns_dict(self):
        sbom = generate_sbom()
        assert isinstance(sbom, dict)

    def test_has_cyclonedx_format(self):
        sbom = generate_sbom()
        assert sbom['bomFormat'] == 'CycloneDX'

    def test_spec_version_1_6(self):
        sbom = generate_sbom()
        assert sbom['specVersion'] == '1.6'

    def test_has_schema_url(self):
        sbom = generate_sbom()
        assert '$schema' in sbom
        assert 'cyclonedx.org' in sbom['$schema']

    def test_has_metadata(self):
        sbom = generate_sbom()
        meta = sbom['metadata']
        assert 'timestamp' in meta
        assert 'tools' in meta
        assert 'component' in meta
        assert 'supplier' in meta

    def test_root_component_is_grepxcel(self):
        sbom = generate_sbom()
        comp = sbom['metadata']['component']
        assert comp['name'] == 'grepxcel'
        assert comp['type'] == 'application'
        assert 'purl' in comp
        assert comp['purl'].startswith('pkg:pypi/grepxcel@')

    def test_root_has_license(self):
        sbom = generate_sbom()
        licenses = sbom['metadata']['component'].get('licenses', [])
        assert len(licenses) >= 1

    def test_supplier_present(self):
        sbom = generate_sbom()
        supplier = sbom['metadata']['supplier']
        assert supplier['name'] == 'scpg'

    def test_has_components(self):
        sbom = generate_sbom()
        assert isinstance(sbom['components'], list)
        assert len(sbom['components']) >= 3

    def test_components_have_purls(self):
        sbom = generate_sbom()
        for comp in sbom['components']:
            assert 'purl' in comp, f"{comp['name']} missing purl"
            assert comp['purl'].startswith('pkg:pypi/')

    def test_components_have_bom_refs(self):
        sbom = generate_sbom()
        refs = set()
        for comp in sbom['components']:
            assert 'bom-ref' in comp
            refs.add(comp['bom-ref'])
        assert len(refs) == len(sbom['components'])

    def test_known_deps_present(self):
        sbom = generate_sbom()
        names = {c['name'].lower() for c in sbom['components']}
        assert 'openpyxl' in names
        assert 'defusedxml' in names

    def test_grepxcel_not_in_components(self):
        """grepxcel itself is the root component, not a dependency."""
        sbom = generate_sbom()
        names = {c['name'].lower() for c in sbom['components']}
        assert 'grepxcel' not in names

    def test_tool_metadata(self):
        sbom = generate_sbom()
        tools = sbom['metadata']['tools']['components']
        assert any(t['name'] == 'grepxcel' for t in tools)


class TestCollectDeps:
    """Dependency collection."""

    def test_collects_grepxcel_deps(self):
        deps = _collect_deps('grepxcel')
        assert 'openpyxl' in deps
        assert 'defusedxml' in deps

    def test_includes_transitive(self):
        deps = _collect_deps('grepxcel')
        assert 'et_xmlfile' in deps


class TestRenderSbom:
    """Serialization."""

    def test_returns_valid_json(self):
        sbom = generate_sbom()
        rendered = render_sbom(sbom)
        parsed = json.loads(rendered)
        assert parsed['bomFormat'] == 'CycloneDX'

    def test_is_indented(self):
        sbom = generate_sbom()
        rendered = render_sbom(sbom)
        assert '\n  ' in rendered


class TestRunSbom:
    """CLI integration."""

    def test_writes_to_stdout(self, capsys):
        import io
        out = io.StringIO()
        code = run_sbom(out=out)
        assert code == 0
        parsed = json.loads(out.getvalue())
        assert parsed['bomFormat'] == 'CycloneDX'

    def test_writes_to_file(self, tmp_path):
        outfile = str(tmp_path / 'sbom.cdx.json')
        code = run_sbom(output=outfile)
        assert code == 0
        parsed = json.loads((tmp_path / 'sbom.cdx.json').read_text())
        assert parsed['specVersion'] == '1.6'
        assert len(parsed['components']) >= 3


class TestCLISubcommand:
    """The 'grepxcel sbom' subcommand."""

    def test_sbom_stdout(self, capsys, tmp_path):
        outfile = str(tmp_path / 'sbom.cdx.json')
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['sbom', '-o', outfile])
        assert exc_info.value.code == 0
        parsed = json.loads((tmp_path / 'sbom.cdx.json').read_text())
        assert parsed['bomFormat'] == 'CycloneDX'

    def test_sbom_help(self, capsys):
        with pytest.raises(SystemExit) as exc_info:
            cli_main(['sbom', '--help'])
        assert exc_info.value.code == 0
        out = capsys.readouterr().out
        assert 'CycloneDX' in out
