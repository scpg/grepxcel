"""Unit tests for `grepxcel doctor` (grepxcel.doctor)."""
import io
import ssl
import urllib.error

import pytest

import grepxcel.doctor as doctor
import grepxcel.proxy_support as ps

_CA_VARS = ('GREPXCEL_CA_BUNDLE', 'REQUESTS_CA_BUNDLE', 'SSL_CERT_FILE', 'CURL_CA_BUNDLE')
_PROXY_VARS = ('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY', 'NO_PROXY')


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    for var in _CA_VARS + _PROXY_VARS:
        monkeypatch.delenv(var, raising=False)
    ps._announced = False
    yield


def _statuses(results):
    return {name: status for status, name, _detail in results}


# ── extract checks ───────────────────────────────────────────────────────────

def test_check_extract_all_ok():
    st = _statuses(doctor.check_extract())
    assert st['openpyxl'] == doctor.OK
    assert st['defusedxml'] == doctor.OK
    assert st['regex'] == doctor.OK
    assert any(k.startswith('Python') and v == doctor.OK for k, v in st.items())


def test_run_doctor_extract_ready():
    out = io.StringIO()
    code = doctor.run_doctor('extract', probe=False, out=out)
    assert code == 0
    text = out.getvalue()
    assert 'extract — core' in text
    assert 'network — proxy / TLS' not in text   # offline area: no proxy section


# ── proxy / CA checks ────────────────────────────────────────────────────────

def test_proxy_section_flags_missing_ca(monkeypatch):
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', '/no/such/ca.pem')
    st = _statuses(doctor.check_proxy_tls(probe=False))
    assert st['CA bundle'] == doctor.FAIL


def test_proxy_section_accepts_existing_ca(monkeypatch, tmp_path):
    ca = tmp_path / 'ca.pem'; ca.write_text('x')
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', str(ca))
    st = _statuses(doctor.check_proxy_tls(probe=False))
    assert st['CA bundle'] == doctor.OK


def test_missing_ca_fails_doctor_and_prints_hint(monkeypatch):
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', '/no/such/ca.pem')
    out = io.StringIO()
    code = doctor.run_doctor('all', probe=False, out=out)
    assert code == 1                              # a ✗ → non-zero exit
    assert 'GREPXCEL_CA_BUNDLE' in out.getvalue() # the actionable cert hint


def test_missing_ca_does_not_fail_extract_area(monkeypatch):
    # The bad CA must not affect `doctor extract` (offline area, no proxy check).
    monkeypatch.setenv('GREPXCEL_CA_BUNDLE', '/no/such/ca.pem')
    assert doctor.run_doctor('extract', probe=False, out=io.StringIO()) == 0


# ── TLS probe classification (mock urlopen — no real network) ────────────────

def _patch_urlopen(monkeypatch, fn):
    monkeypatch.setattr(doctor.urllib.request, 'urlopen', fn)


def test_probe_ok_on_success(monkeypatch):
    _patch_urlopen(monkeypatch, lambda *a, **k: object())
    status, _name, _detail = doctor.tls_probe()
    assert status == doctor.OK


def test_probe_ok_on_http_error(monkeypatch):
    def raise_http(*a, **k):
        raise urllib.error.HTTPError('u', 405, 'Method Not Allowed', {}, None)
    _patch_urlopen(monkeypatch, raise_http)
    assert doctor.tls_probe()[0] == doctor.OK     # reached server, TLS verified


def test_probe_fail_on_cert_error(monkeypatch):
    def raise_cert(*a, **k):
        raise urllib.error.URLError(ssl.SSLCertVerificationError('certificate verify failed'))
    _patch_urlopen(monkeypatch, raise_cert)
    assert doctor.tls_probe()[0] == doctor.FAIL


def test_probe_warn_on_unreachable(monkeypatch):
    def raise_timeout(*a, **k):
        raise urllib.error.URLError('timed out')
    _patch_urlopen(monkeypatch, raise_timeout)
    assert doctor.tls_probe()[0] == doctor.WARN


# ── server backend check ────────────────────────────────────────────────────

def test_check_server_ok_when_reachable(monkeypatch):
    import json
    body = json.dumps({'data': [{'id': 'qwen2.5-coder-7b'}]}).encode()
    resp = io.BytesIO(body)
    monkeypatch.setattr(doctor.urllib.request, 'urlopen', lambda *a, **k: resp)
    st = _statuses(doctor.check_server(url='http://localhost:1234/v1'))
    assert st['server'] == doctor.OK


def test_check_server_warn_when_unreachable(monkeypatch):
    def raise_err(*a, **k):
        raise urllib.error.URLError('Connection refused')
    monkeypatch.setattr(doctor.urllib.request, 'urlopen', raise_err)
    st = _statuses(doctor.check_server(url='http://localhost:1234/v1'))
    assert st['server'] == doctor.WARN


def test_check_server_warn_when_no_models(monkeypatch):
    import json
    body = json.dumps({'data': []}).encode()
    resp = io.BytesIO(body)
    monkeypatch.setattr(doctor.urllib.request, 'urlopen', lambda *a, **k: resp)
    st = _statuses(doctor.check_server(url='http://localhost:1234/v1'))
    assert st['server'] == doctor.WARN


def test_check_server_refuses_non_http_scheme(monkeypatch):
    """A file:// (or other non-http) URL must be refused WITHOUT opening it —
    otherwise GREPXCEL_SERVER_URL=file:///etc/passwd is a file-read primitive."""
    def boom(*a, **k):
        raise AssertionError('urlopen must not be called for a non-http(s) URL')
    monkeypatch.setattr(doctor.urllib.request, 'urlopen', boom)
    res = doctor.check_server(url='file:///etc/passwd')
    # FAIL status and a message that explains the refusal
    assert any(sev == doctor.FAIL for sev, _, _ in res)
    assert any('http' in msg.lower() or 'refus' in msg.lower() for _, _, msg in res)
