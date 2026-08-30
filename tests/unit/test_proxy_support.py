"""Unit tests for corporate-proxy / TLS-trust support (grepxcel.proxy_support).

TLS verification is never disabled; these tests assert the trust *configuration*
is resolved and propagated correctly.
"""
import pytest

import grepxcel.proxy_support as ps

_CA_VARS = ('GREPXCEL_CA_BUNDLE', 'REQUESTS_CA_BUNDLE', 'SSL_CERT_FILE', 'CURL_CA_BUNDLE')
_PROXY_VARS = ('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY', 'NO_PROXY',
               'http_proxy', 'https_proxy', 'all_proxy', 'no_proxy')


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    """Start each test from a known-clean proxy/CA environment."""
    for var in _CA_VARS + _PROXY_VARS:
        monkeypatch.delenv(var, raising=False)
    ps._announced = False
    yield


# ── resolve_ca_bundle ────────────────────────────────────────────────────────

def test_resolve_none_when_unset():
    assert ps.resolve_ca_bundle() is None


def test_resolve_explicit_wins(monkeypatch):
    monkeypatch.setenv('REQUESTS_CA_BUNDLE', '/from/env.pem')
    assert ps.resolve_ca_bundle('/explicit.pem') == '/explicit.pem'


def test_resolve_env_precedence(monkeypatch):
    monkeypatch.setenv('SSL_CERT_FILE', '/ssl.pem')
    monkeypatch.setenv('REQUESTS_CA_BUNDLE', '/req.pem')
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', '/grepxcel.pem')
    assert ps.resolve_ca_bundle() == '/grepxcel.pem'   # our var first


def test_resolve_falls_through_to_ssl_cert_file(monkeypatch):
    monkeypatch.setenv('SSL_CERT_FILE', '/ssl.pem')
    assert ps.resolve_ca_bundle() == '/ssl.pem'


# ── enable_corporate_tls ─────────────────────────────────────────────────────

def test_enable_propagates_ca_to_all_libs(monkeypatch, tmp_path):
    ca = tmp_path / 'corp-ca.pem'
    ca.write_text('-----BEGIN CERTIFICATE-----\n')
    status = ps.enable_corporate_tls(str(ca), announce=False)
    assert status['ca_bundle'] == str(ca)
    assert status['ca_bundle_exists'] is True
    import os
    for var in ('REQUESTS_CA_BUNDLE', 'SSL_CERT_FILE', 'CURL_CA_BUNDLE'):
        assert os.environ[var] == str(ca)   # propagated so requests + httpx see it


def test_enable_reports_missing_ca(monkeypatch):
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', '/does/not/exist.pem')
    status = ps.enable_corporate_tls(announce=False)
    assert status['ca_bundle'] == '/does/not/exist.pem'
    assert status['ca_bundle_exists'] is False


def test_enable_noop_when_nothing_set(monkeypatch):
    status = ps.enable_corporate_tls(announce=False)
    assert status['ca_bundle'] is None
    assert status['proxies'] == {}


def test_enable_announces_once(monkeypatch, tmp_path, capsys):
    ca = tmp_path / 'ca.pem'; ca.write_text('x')
    ps.enable_corporate_tls(str(ca), announce=True)
    first = capsys.readouterr().err
    assert 'EXPERIMENTAL' in first
    ps.enable_corporate_tls(str(ca), announce=True)
    assert capsys.readouterr().err == ''     # second call is silent (once per process)


# ── proxy_env ────────────────────────────────────────────────────────────────

def test_proxy_env_reports_set_vars(monkeypatch):
    monkeypatch.setenv('HTTPS_PROXY', 'http://proxy.corp:8080')
    monkeypatch.setenv('NO_PROXY', 'localhost,.corp')
    env = ps.proxy_env()
    assert env['HTTPS_PROXY'] == 'http://proxy.corp:8080'
    assert env['NO_PROXY'] == 'localhost,.corp'


def test_proxy_env_lowercase_normalized(monkeypatch):
    monkeypatch.setenv('https_proxy', 'http://p:3128')
    assert ps.proxy_env().get('HTTPS_PROXY') == 'http://p:3128'


# ── cert-error helpers ───────────────────────────────────────────────────────

def test_looks_like_cert_error_true():
    exc = Exception('[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed')
    assert ps.looks_like_cert_error(exc) is True


def test_looks_like_cert_error_false():
    assert ps.looks_like_cert_error(ValueError('totally unrelated')) is False


def test_cert_failure_hint_is_actionable():
    hint = ps.cert_failure_hint()
    assert 'GREPXCEL_CA_BUNDLE' in hint
    assert 'truststore' in hint
    assert 'verification stays ON' in hint.lower() or 'never disable' in hint.lower()


# ── make_httpx_client ────────────────────────────────────────────────────────

def test_make_httpx_client_with_ca(monkeypatch):
    # anthropic >= 1.x and openai >= 3.x both require httpx2; make_httpx_client()
    # prefers httpx2.Client and falls back to httpx.Client on older installs.
    certifi = pytest.importorskip('certifi')
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', certifi.where())
    client = ps.make_httpx_client()
    try:
        import httpx2
        assert isinstance(client, httpx2.Client), (
            f'Expected httpx2.Client, got {type(client)}'
        )
    except ImportError:
        import httpx
        assert isinstance(client, httpx.Client), (
            f'Expected httpx.Client (httpx2 unavailable), got {type(client)}'
        )
    client.close()


def test_make_httpx_client_none_without_config(monkeypatch):
    # No CA bundle and (in the test env) no truststore → fall back to SDK default.
    if ps.truststore_available():
        pytest.skip('truststore installed — would return a client, not None')
    assert ps.make_httpx_client() is None
