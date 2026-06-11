"""Corporate TLS-inspection proxy support (verification ALWAYS stays on).

grepxcel makes outbound HTTPS in two places: the local-model download
(``huggingface_hub`` → ``requests``) and the cloud ``draft`` backends
(Anthropic / OpenAI / Google SDKs → ``httpx``). Behind a corporate intercept
proxy (e.g. NetSkope, Zscaler), TLS is re-signed by a company CA that Python's
bundled ``certifi`` store does not trust, producing::

    [SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: self-signed
    certificate in certificate chain

This module wires up trust for that CA — never by disabling verification:

* :func:`resolve_ca_bundle`  — find a user-supplied CA bundle (flag / env vars)
* :func:`enable_corporate_tls` — startup hook: propagate an explicit CA bundle to
  the env vars the HTTP libraries read, and (if ``truststore`` is installed) make
  Python trust the OS certificate store, where corporate IT normally installs the
  CA. Fixes both the HF download and the cloud SDKs at once.
* :func:`make_httpx_client` — an ``httpx.Client`` honoring the CA bundle / OS
  store, to hand to a cloud SDK constructor.
* :func:`looks_like_cert_error` / :func:`cert_failure_hint` — turn a cryptic SSL
  failure into actionable guidance.
* :func:`proxy_env` — report active proxy env vars (used by ``grepxcel doctor``).

.. warning::
   This support has NOT been validated against a real TLS-inspection proxy.
   :func:`enable_corporate_tls` prints a one-time caveat when it is in effect.

All third-party imports here are deferred/guarded so a core ``extract``-only
install (no ``httpx`` / ``truststore`` / ``requests``) is never burdened.
"""
from __future__ import annotations

import os
import sys

# Explicit CA-bundle env vars we read, in priority order (our own var first).
_CA_ENV_VARS = ('GREPXCEL_CA_BUNDLE', 'REQUESTS_CA_BUNDLE', 'SSL_CERT_FILE', 'CURL_CA_BUNDLE')
# Env vars the HTTP libraries read; an explicit bundle is mirrored into all of them
# (requests reads REQUESTS_CA_BUNDLE; httpx/stdlib read SSL_CERT_FILE).
_CA_PROPAGATE = ('REQUESTS_CA_BUNDLE', 'SSL_CERT_FILE', 'CURL_CA_BUNDLE')
_PROXY_VARS = ('HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY', 'NO_PROXY',
               'http_proxy', 'https_proxy', 'all_proxy', 'no_proxy')

_announced = False  # so the experimental caveat prints at most once per process


def resolve_ca_bundle(explicit: str | None = None) -> str | None:
    """Return a CA-bundle path from the explicit flag or env vars, else ``None``.

    The path is returned even if it does not exist on disk — callers (and
    ``grepxcel doctor``) report a missing file rather than silently ignoring it.
    """
    if explicit:
        return explicit
    for var in _CA_ENV_VARS:
        val = os.environ.get(var)
        if val:
            return val
    return None


def truststore_available() -> bool:
    try:
        import truststore  # noqa: F401
        return True
    except Exception:
        return False


def _truststore_inject() -> bool:
    """Make Python trust the OS certificate store via ``truststore`` if installed.

    Returns True if injection happened. Safe/no-op when truststore is absent.
    """
    try:
        import truststore
    except Exception:
        return False
    try:
        truststore.inject_into_ssl()
        return True
    except Exception:
        return False


def enable_corporate_tls(ca_bundle: str | None = None, *, announce: bool = True) -> dict:
    """Startup hook for commands that go online (``draft``, ``doctor``).

    * If an explicit CA bundle is resolved, propagate it to REQUESTS_CA_BUNDLE /
      SSL_CERT_FILE / CURL_CA_BUNDLE so ``requests`` (HF download) and ``httpx``
      (cloud SDKs) both pick it up.
    * Attempt ``truststore`` OS-store injection (no-op if not installed).

    Returns a status dict and, when any proxy/CA configuration is in effect,
    prints a one-time EXPERIMENTAL caveat to stderr. TLS verification stays ON.
    """
    global _announced
    resolved = resolve_ca_bundle(ca_bundle)
    if resolved:
        for var in _CA_PROPAGATE:
            os.environ.setdefault(var, resolved)
    truststore_active = _truststore_inject()
    proxies = proxy_env()
    status = {
        'ca_bundle': resolved,
        'ca_bundle_exists': bool(resolved) and os.path.isfile(resolved),
        'truststore_active': truststore_active,
        'proxies': proxies,
    }
    active = bool(resolved) or truststore_active or bool(proxies)
    if announce and active and not _announced:
        _announced = True
        bits = []
        if proxies:
            bits.append('proxy env detected')
        if resolved:
            bits.append(f'CA bundle {resolved!r}')
        if truststore_active:
            bits.append('OS trust store via truststore')
        print(
            '[!] Corporate-proxy / custom-CA TLS support is ACTIVE '
            f'({"; ".join(bits)}).\n'
            '    This support is EXPERIMENTAL and has not been tested against a '
            'real TLS-inspection proxy.\n'
            '    If outbound HTTPS still fails, run `grepxcel doctor` for guidance.',
            file=sys.stderr,
        )
    return status


def make_httpx_client():
    """Return an ``httpx.Client`` honoring the resolved CA bundle or the OS trust
    store, or ``None`` to let the SDK use its default. Verification is ALWAYS on.
    """
    try:
        import httpx
    except Exception:
        return None
    import ssl
    ca = resolve_ca_bundle()
    if ca and os.path.isfile(ca):
        # Build an SSLContext from the bundle (httpx deprecated verify=<str path>).
        ctx = ssl.create_default_context(cafile=ca)
        return httpx.Client(verify=ctx)
    # No explicit bundle: prefer the OS trust store if truststore is present
    # (httpx does NOT read REQUESTS_CA_BUNDLE, so this is how we cover it).
    try:
        import truststore
        ctx = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        return httpx.Client(verify=ctx)
    except Exception:
        return None  # SDK default (certifi) — correct when there is no proxy


def looks_like_cert_error(exc: BaseException) -> bool:
    """Heuristic: does this exception look like a TLS trust failure?"""
    text = f'{type(exc).__name__}: {exc}'.lower()
    needles = (
        'certificate_verify_failed',
        'certificate verify failed',
        'self-signed certificate',
        'self signed certificate',
        'unable to get local issuer',
        'ssl: ',
    )
    return any(n in text for n in needles)


def cert_failure_hint() -> str:
    """Actionable guidance for a TLS verification failure behind a corp proxy."""
    return (
        'TLS certificate verification failed. This usually means a corporate '
        'TLS-inspection proxy (e.g. NetSkope / Zscaler) is re-signing HTTPS with '
        'a company CA that Python does not trust yet.\n'
        '  Fixes (verification stays ON — never disable it):\n'
        '   1. Best: have IT install the corporate root CA in your OS trust '
        'store, then `pip install truststore` so grepxcel uses it automatically.\n'
        '   2. Or point grepxcel at the CA .pem file:\n'
        '        export GREPXCEL_CA_BUNDLE=/path/to/corporate-ca.pem\n'
        '        (REQUESTS_CA_BUNDLE / SSL_CERT_FILE are also honored)\n'
        '   3. If your network requires a proxy, set HTTPS_PROXY (and NO_PROXY).\n'
        '  Then run `grepxcel doctor` to confirm the setup.'
    )


def proxy_env() -> dict:
    """Return the proxy-related env vars that are currently set (de-duplicated,
    upper-case keys preferred)."""
    out: dict[str, str] = {}
    for var in _PROXY_VARS:
        val = os.environ.get(var)
        if val:
            out.setdefault(var.upper(), val)
    return out
