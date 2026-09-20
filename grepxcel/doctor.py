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

from . import __version__, proxy_support
from .color import MARK_FAIL, MARK_OK, MARK_WARN, colorize_marks, paint, should_color

OK, WARN, FAIL = 'ok', 'warn', 'fail'
_MARK = {OK: MARK_OK, WARN: MARK_WARN, FAIL: MARK_FAIL}

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


def check_config_dir() -> list[Result]:
    """Config directory existence and writability."""
    from .cli import _config_dir
    res: list[Result] = []
    cfg = _config_dir()
    if not os.path.exists(cfg):
        res.append((WARN, 'config dir',
                    f'{cfg} — does not exist yet (created on first use)'))
    elif not os.access(cfg, os.W_OK):
        res.append((FAIL, 'config dir',
                    f'{cfg} — exists but is not writable (check permissions)'))
    else:
        res.append((OK, 'config dir', f'{cfg} (writable)'))
    return res


def check_autocomplete() -> list[Result]:
    """Shell TAB completion via argcomplete."""
    res: list[Result] = []

    # argcomplete is a core dep — always installed; just confirm it
    if _have('argcomplete'):
        res.append((OK, 'argcomplete', 'installed'))
    else:
        res.append((FAIL, 'argcomplete',
                    'missing — reinstall grepxcel (core dependency)'))

    # Detect shell and find the right RC file
    shell_bin = os.environ.get('SHELL', '')
    shell_name = os.path.basename(shell_bin) if shell_bin else ''
    rc_map = {
        'bash': os.path.expanduser('~/.bashrc'),
        'zsh':  os.path.expanduser('~/.zshrc'),
        'fish': os.path.expanduser('~/.config/fish/config.fish'),
    }
    rc_file = rc_map.get(shell_name)

    if shell_name in ('bash', 'zsh'):
        register_cmd = 'eval "$(register-python-argcomplete grepxcel)"'
    elif shell_name == 'fish':
        register_cmd = 'register-python-argcomplete --shell fish grepxcel | source'
    else:
        register_cmd = 'eval "$(register-python-argcomplete grepxcel)"'

    # Check if the registration is present in the RC file
    in_rc = False
    if rc_file and os.path.isfile(rc_file):
        try:
            in_rc = 'register-python-argcomplete grepxcel' in open(rc_file).read()
        except OSError:
            pass

    rc_label = f'persistent ({shell_name})' if shell_name else 'persistent'
    rc_path = rc_file or f'~/{shell_name}rc'
    if in_rc:
        res.append((OK, rc_label, f'registered in {rc_path}'))
    else:
        res.append((WARN, rc_label,
                    f'not in {rc_path} — to activate permanently add:  {register_cmd}'))

    return res


def check_web_wizard(probe: bool = True) -> list[Result]:
    """FastAPI / uvicorn / jinja2 for grepxcel web-wizard."""
    res: list[Result] = []
    all_present = True
    for mod, pkg in (('fastapi', 'fastapi'), ('uvicorn', 'uvicorn'), ('jinja2', 'jinja2')):
        if _have(mod):
            res.append((OK, pkg, 'installed'))
        else:
            all_present = False
            res.append((WARN, pkg, "not installed — pip install 'grepxcel[web]'"))

    if probe and all_present:
        import socket
        port = int(os.environ.get('GREPXCEL_WEB_PORT', 8765))
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.3)
            in_use = s.connect_ex(('127.0.0.1', port)) == 0
            s.close()
        except OSError:
            in_use = False
        if in_use:
            res.append((WARN, f'port {port}',
                        f'already in use — pass --port to web-wizard to override'))
        else:
            res.append((OK, f'port {port}', 'available'))

    return res


def check_mcp() -> list[Result]:
    """MCP server package for grepxcel mcp."""
    res: list[Result] = []
    if _have('mcp'):
        res.append((OK, 'mcp', "installed — start with: grepxcel mcp"))
    else:
        res.append((WARN, 'mcp', "not installed — pip install 'grepxcel[mcp]'"))
    res.append((OK, 'mcp config',
                "get Claude Code config with: grepxcel mcp-config"))
    return res


def check_watcher() -> list[Result]:
    """File-watcher for grepxcel watch."""
    res: list[Result] = []
    if _have('watchdog'):
        res.append((OK, 'watchdog', 'installed'))
    else:
        res.append((WARN, 'watchdog',
                    "not installed — pip install 'grepxcel[watch]'"))
    return res


def check_dataframe() -> list[Result]:
    """DataFrame export via grepxcel extract_df."""
    res: list[Result] = []
    for mod, pkg in (('pandas', 'pandas'), ('polars', 'polars')):
        if _have(mod):
            res.append((OK, pkg, 'installed'))
        else:
            res.append((WARN, pkg,
                        f"not installed — pip install 'grepxcel[{pkg}]'"))
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
    have, key = _have('anthropic'), bool(os.environ.get('ANTHROPIC_API_KEY'))
    if have and key:
        res.append((OK, 'claude backend', 'anthropic installed, ANTHROPIC_API_KEY set'))
    elif have:
        res.append((WARN, 'claude backend', 'anthropic installed but ANTHROPIC_API_KEY not set'))
    else:
        res.append((WARN, 'claude backend',
                    "not configured — pip install 'grepxcel[draft-cloud]' and set ANTHROPIC_API_KEY"))
    res.append((WARN, 'gemini backend', 'planned for a future release (disabled)'))
    res.append((WARN, 'github backend',
                'currently unavailable — GitHub retired the free-tier Models endpoint '
                '(HTTP 410 retirement brownout); implementation preserved for future re-enable'))
    have_openai = _have('openai')
    nv_key = bool(os.environ.get('NVIDIA_API_KEY'))
    if have_openai and nv_key:
        res.append((OK, 'nvidia backend',
                    'openai installed, NVIDIA_API_KEY set — free-tier NVIDIA NIM available'))
    elif have_openai:
        res.append((WARN, 'nvidia backend',
                    'openai installed but NVIDIA_API_KEY not set — '
                    'get a free key at https://build.nvidia.com'))
    else:
        res.append((WARN, 'nvidia backend',
                    "not configured — pip install openai and set NVIDIA_API_KEY "
                    "(free key at https://build.nvidia.com)"))
    return res


def check_server(url: str = 'http://localhost:1234/v1') -> list[Result]:
    """Probe an OpenAI-compatible server at *url*/models."""
    import json
    from .security import is_http_url
    res: list[Result] = []
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
    out = out or sys.stdout
    # Activate any configured corporate trust so the probe reflects reality.
    proxy_support.enable_corporate_tls(announce=False)

    sections: list[tuple[str, str, list[Result]]] = []
    if area in ('extract', 'all'):
        sections.append(('extract — core',
                         'verify that extraction from Excel files works end-to-end',
                         check_extract()))
    if area == 'all':
        sections.append(('config',
                         'grepxcel config directory',
                         check_config_dir()))
        sections.append(('autocomplete',
                         'shell TAB completion via argcomplete',
                         check_autocomplete()))
        sections.append(('web wizard',
                         'browser-based pattern wizard (grepxcel web-wizard)',
                         check_web_wizard(probe=probe)))
        sections.append(('mcp server',
                         'Model Context Protocol server (grepxcel mcp)',
                         check_mcp()))
        sections.append(('watcher',
                         'file-change watcher (grepxcel watch)',
                         check_watcher()))
        sections.append(('dataframe export',
                         'pandas / polars DataFrame output via extract_df()',
                         check_dataframe()))
    if area in ('draft', 'all'):
        sections.append(('draft — credentials (.env)',
                         'API keys and .env config used by cloud pattern drafters',
                         check_env(strict_env)))
        sections.append(('draft — local model',
                         'local AI model to generate a starter pattern from your data file',
                         check_draft_local()))
        sections.append(('draft — cloud backends',
                         'cloud AI services for pattern drafting (Anthropic, Gemini, NVIDIA)',
                         check_draft_cloud()))
        srv_url = os.environ.get('GREPXCEL_SERVER_URL', 'http://localhost:1234/v1')
        sections.append(('draft — server backend',
                         'local OpenAI-compatible API server (LM Studio, Ollama, vLLM, …)',
                         check_server(url=srv_url)))
        sections.append(('network — proxy / TLS',
                         'corporate proxy, TLS certificates, and outbound connectivity',
                         check_proxy_tls(probe=probe)))

    color = should_color(out)
    print(f'grepxcel doctor — checking: {area}   version: {__version__}\n' + '─' * 62, file=out)
    any_fail = False
    proxy_fail = False
    for title, desc, checks in sections:
        subtitle = paint(f'  # {desc}', 'dim', color)
        print(f'\n  {title}{subtitle}', file=out)
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
        msg = paint('Not ready', 'red', color) + f' — resolve the {MARK_FAIL} items above.'
        print(colorize_marks(f'  {MARK_FAIL} {msg}', color), file=out)
    else:
        msg = paint('Ready.', 'green', color) + f' ({MARK_WARN} items are optional / situational.)'
        print(colorize_marks(f'  {MARK_OK} {msg}', color), file=out)
    return 1 if any_fail else 0
