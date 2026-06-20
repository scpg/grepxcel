"""`grepxcel doctor` — preflight diagnostics.

Reports, with a ✓ / ⚠ / ✗ checklist, whether the environment is ready to run
``extract`` and/or ``draft`` — Python version, core/optional dependencies, API
keys, the local-model cache, and corporate-proxy / TLS trust (with a live
handshake probe). Exits non-zero if the selected area has a hard failure, so it
is scriptable in CI or onboarding.

Design notes:
* Optional packages are detected with ``importlib.util.find_spec`` — never
  imported — so a probe for a missing extra cannot crash.
* The TLS probe uses stdlib ``urllib`` which honors ``HTTPS_PROXY`` and the
  (truststore-injected) default SSL context, so it reflects the real configured
  trust path.
"""
from __future__ import annotations

import importlib.util
import os
import shutil
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request

from . import proxy_support
from .color import colorize_marks, should_color

OK, WARN, FAIL = 'ok', 'warn', 'fail'
_MARK = {OK: '✓', WARN: '⚠', FAIL: '✗'}

# A single check result: (status, name, detail/hint).
Result = tuple


def _have(module: str) -> bool:
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ValueError):
        return False


# ── individual check groups ─────────────────────────────────────────────────

def check_extract() -> list[Result]:
    res: list[Result] = []
    v = sys.version_info
    if v >= (3, 11):
        res.append((OK, f'Python {v.major}.{v.minor}', 'meets the >= 3.11 floor'))
    else:
        res.append((FAIL, f'Python {v.major}.{v.minor}', 'grepxcel requires Python >= 3.11'))
    from .pattern_parser import (CURRENT_PATTERN_VERSION,
                                 MIN_SUPPORTED_PATTERN_VERSION)
    pv = (f'{CURRENT_PATTERN_VERSION}'
          if MIN_SUPPORTED_PATTERN_VERSION == CURRENT_PATTERN_VERSION
          else f'{MIN_SUPPORTED_PATTERN_VERSION}..{CURRENT_PATTERN_VERSION}')
    res.append((OK, 'pattern.version',
                f'understands {pv} (a pattern with no version is read as 1)'))
    for mod, why in (('openpyxl', 'reads .xlsx files'),
                     ('defusedxml', 'XXE protection (fail-closed)'),
                     ('regex', 'ReDoS-bounded matching')):
        if _have(mod):
            res.append((OK, mod, why))
        else:
            res.append((FAIL, mod, f'required ({why}) — reinstall grepxcel'))
    return res


def _model_cache_dir() -> str | None:
    override = os.environ.get('GREPXCEL_MODEL_DIR')
    if override:
        return override
    try:
        from platformdirs import user_cache_dir
        return os.path.join(user_cache_dir('grepxcel'), 'models')
    except Exception:
        return os.path.expanduser('~/.cache/grepxcel/models')


def _find_gguf(cache_dir: str) -> str | None:
    if not cache_dir or not os.path.isdir(cache_dir):
        return None
    for root, _dirs, files in os.walk(cache_dir):
        for f in files:
            if f.lower().endswith('.gguf'):
                return os.path.join(root, f)
    return None


def check_draft_local() -> list[Result]:
    res: list[Result] = []
    for mod, pkg in (('llama_cpp', 'llama-cpp-python'),
                     ('huggingface_hub', 'huggingface_hub'),
                     ('platformdirs', 'platformdirs')):
        if _have(mod):
            res.append((OK, pkg, 'installed'))
        else:
            res.append((WARN, pkg, "needed for local draft — pip install 'grepxcel[suggest]'"))
    cache = _model_cache_dir()
    gguf = _find_gguf(cache) if cache else None
    if gguf:
        res.append((OK, 'local model', f'cached: {gguf}'))
    else:
        res.append((WARN, 'local model', f'not downloaded yet (~5 GB on first run) → {cache}'))
    try:
        target = cache if cache and os.path.isdir(cache) else os.path.expanduser('~')
        free_gb = shutil.disk_usage(target).free / 1e9
        if free_gb >= 6:
            res.append((OK, 'disk space', f'{free_gb:.1f} GB free'))
        else:
            res.append((WARN, 'disk space', f'only {free_gb:.1f} GB free; local model is ~5 GB'))
    except OSError:
        pass
    return res


def check_draft_cloud() -> list[Result]:
    res: list[Result] = []
    for name, mod, env in (('claude', 'anthropic', 'ANTHROPIC_API_KEY'),
                           ('github', 'openai', 'GITHUB_TOKEN')):
        have, key = _have(mod), bool(os.environ.get(env))
        if have and key:
            res.append((OK, f'{name} backend', f'{mod} installed, {env} set'))
        elif have:
            res.append((WARN, f'{name} backend', f'{mod} installed but {env} not set'))
        else:
            res.append((WARN, f'{name} backend',
                        f"not configured — pip install 'grepxcel[draft-cloud]' and set {env}"))
    res.append((WARN, 'gemini backend', 'planned for a future release (disabled)'))
    return res


def check_server(url: str = 'http://localhost:1234/v1') -> list[Result]:
    """Probe an OpenAI-compatible server at *url*/models."""
    import json
    from .security import is_http_url
    res: list[Result] = []
    # Never let urlopen handle file://, ftp://, data: etc. — a non-http(s)
    # GREPXCEL_SERVER_URL would otherwise be a file-read / SSRF primitive.
    if not is_http_url(url):
        return [(FAIL, 'server', f'{url} — refusing to probe a non-http(s) URL')]
    models_url = url.rstrip('/') + '/models'
    try:
        resp = urllib.request.urlopen(
            urllib.request.Request(models_url, method='GET'),
            timeout=3,
        )
        data = json.loads(resp.read())
        model_ids = [m.get('id', '?') for m in data.get('data', [])]
        if model_ids:
            res.append((OK, 'server', f'{url} — loaded: {", ".join(model_ids[:3])}'))
        else:
            res.append((WARN, 'server', f'{url} — reachable but no models loaded'))
    except (urllib.error.URLError, OSError) as exc:
        res.append((WARN, 'server', f'{url} — not reachable: {exc}'))
    except Exception as exc:
        res.append((WARN, 'server', f'{url} — probe error: {exc}'))
    return res


def tls_probe(url: str = 'https://huggingface.co', timeout: float = 6.0) -> Result:
    """Live HTTPS handshake honoring proxy env + configured CA trust.

    Returns OK if TLS verified (even on an HTTP error response — the connection
    and certificate were fine), FAIL on a certificate-trust failure (the corp
    proxy case), WARN if simply unreachable (offline / blocked)."""
    # Only ever probe http(s) — never let urlopen handle file://, ftp://, etc.
    from .security import is_http_url
    if not is_http_url(url):
        return (FAIL, f'TLS handshake {url}', 'refusing to probe a non-http(s) URL')
    proxy_support.enable_corporate_tls(announce=False)
    try:
        # nosec B310 — scheme is validated to http/https just above.
        urllib.request.urlopen(urllib.request.Request(url, method='HEAD'), timeout=timeout)  # nosec B310
        return (OK, f'TLS handshake {url}', 'verified')
    except urllib.error.HTTPError:
        return (OK, f'TLS handshake {url}', 'reached server, certificate verified')
    except urllib.error.URLError as exc:
        reason = getattr(exc, 'reason', exc)
        if isinstance(reason, ssl.SSLCertVerificationError) or proxy_support.looks_like_cert_error(
                reason if isinstance(reason, BaseException) else exc):
            return (FAIL, f'TLS handshake {url}', f'certificate NOT trusted: {reason}')
        return (WARN, f'TLS handshake {url}', f'could not connect: {reason} (offline / blocked?)')
    except Exception as exc:  # pragma: no cover - defensive
        if proxy_support.looks_like_cert_error(exc):
            return (FAIL, f'TLS handshake {url}', f'certificate NOT trusted: {exc}')
        return (WARN, f'TLS handshake {url}', f'probe error: {exc}')


def check_proxy_tls(probe: bool = True) -> list[Result]:
    res: list[Result] = []
    proxies = proxy_support.proxy_env()
    if proxies:
        res.append((OK, 'proxy env', ', '.join(f'{k}={v}' for k, v in proxies.items())))
    else:
        res.append((OK, 'proxy env', 'none set (direct connection)'))
    ca = proxy_support.resolve_ca_bundle()
    if ca and os.path.isfile(ca):
        res.append((OK, 'CA bundle', ca))
    elif ca:
        res.append((FAIL, 'CA bundle', f'set to {ca} but that file does not exist'))
    else:
        res.append((WARN, 'CA bundle',
                    'none set (GREPXCEL_CA_BUNDLE / REQUESTS_CA_BUNDLE / SSL_CERT_FILE)'))
    if proxy_support.truststore_available():
        res.append((OK, 'truststore', 'installed — OS trust store is used'))
    else:
        res.append((WARN, 'truststore',
                    'not installed — `pip install truststore` to trust the OS CA store'))
    if probe:
        res.append(tls_probe())
    return res


def check_env(strict_env: bool = False) -> list[Result]:
    """Show how cloud-credential .env resolution will be decided, so the user
    always knows which source is used before any cloud call."""
    from .cli import _discover_project_env, _config_dir, _config_dir_env
    res: list[Result] = []
    cfg = _config_dir()
    res.append((OK, 'config dir', cfg))
    cfg_env = _config_dir_env()
    if cfg_env:
        res.append((OK, 'config .env', f'{cfg_env} (loaded as fallback)'))
    else:
        res.append((OK, 'config .env', f'none at {os.path.join(cfg, ".env")}'))

    project_env, in_project = _discover_project_env(os.getcwd())
    if project_env and in_project:
        res.append((OK, 'project .env', project_env))
    elif project_env and not in_project:
        if strict_env:
            res.append((FAIL, 'project .env',
                        f'{project_env} — out-of-project, REFUSED (--strict-env)'))
        else:
            res.append((WARN, 'project .env',
                        f'{project_env} — out-of-project (loaded; would be refused '
                        f'with --strict-env)'))
    else:
        res.append((OK, 'project .env', 'none found'))
    res.append((OK, 'strict-env', 'on' if strict_env else 'off'))
    return res


# ── runner ──────────────────────────────────────────────────────────────────

def run_doctor(area: str = 'all', probe: bool = True, out=None,
               strict_env: bool = False) -> int:
    """Run the selected checks, print a checklist, return an exit code
    (0 = ready, 1 = a hard failure in the selected area)."""
    out = out or sys.stderr
    # Activate any configured corporate trust so the probe reflects reality.
    proxy_support.enable_corporate_tls(announce=False)

    sections: list[tuple[str, list[Result]]] = []
    if area in ('extract', 'all'):
        sections.append(('extract — core', check_extract()))
    if area in ('draft', 'all'):
        sections.append(('draft — credentials (.env)', check_env(strict_env)))
        sections.append(('draft — local model', check_draft_local()))
        sections.append(('draft — cloud backends', check_draft_cloud()))
        srv_url = os.environ.get('GREPXCEL_SERVER_URL', 'http://localhost:1234/v1')
        sections.append(('draft — server backend', check_server(url=srv_url)))
        sections.append(('network — proxy / TLS', check_proxy_tls(probe=probe)))

    color = should_color(out)
    print(f'grepxcel doctor — checking: {area}\n' + '─' * 62, file=out)
    any_fail = False
    proxy_fail = False
    for title, checks in sections:
        print(f'\n  {title}', file=out)
        for status, name, detail in checks:
            if status == FAIL:
                any_fail = True
                if 'TLS' in name or 'CA bundle' in name:
                    proxy_fail = True
            print(colorize_marks(
                f'    {_MARK[status]}  {name:<28} {detail}', color), file=out)

    if proxy_fail:
        print('\n' + proxy_support.cert_failure_hint(), file=out)

    print('\n' + '─' * 62, file=out)
    if any_fail:
        print(colorize_marks('  ✗ Not ready — resolve the ✗ items above.', color), file=out)
    else:
        print(colorize_marks('  ✓ Ready. (⚠ items are optional / situational.)', color), file=out)
    return 1 if any_fail else 0
