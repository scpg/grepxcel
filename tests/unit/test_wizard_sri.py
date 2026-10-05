"""Every external script in the wizard must carry a verifiable SRI hash.

The wizard loads three libraries from cdnjs. Nothing scans them: there is no
package.json, so neither Dependabot nor Snyk sees those versions, and they are
pinned and bumped by hand. Subresource Integrity is what stops a compromised or
MITM'd CDN serving different JavaScript into a tool people point at their own
spreadsheets.

SRI introduces its own failure mode, which these tests are really about: bump the
version and forget the hash, and the browser silently refuses to run the script.
The page still loads, the library is just gone. The offline tests below catch the
shape of that mistake; the network-marked one catches the substance.
"""
import base64
import hashlib
import re
import urllib.request
from pathlib import Path

import pytest

_TEMPLATE = Path(__file__).resolve().parents[2] / 'grepxcel' / 'templates' / 'wizard.html'
_HTML = _TEMPLATE.read_text(encoding='utf-8')

#: Any <script> pointing at an absolute http(s) URL, with its full attribute set.
_EXTERNAL_SCRIPT = re.compile(
    r'<script\b(?P<attrs>[^>]*\bsrc="(?P<src>https?://[^"]+)"[^>]*)>',
    re.IGNORECASE | re.DOTALL,
)


def _external_scripts():
    return [(m.group('src'), m.group('attrs')) for m in _EXTERNAL_SCRIPT.finditer(_HTML)]


def test_there_are_external_scripts_to_check():
    """Guards against the regex silently matching nothing after a refactor."""
    assert _external_scripts(), 'no external <script> found — has the markup changed?'


@pytest.mark.parametrize('src,attrs', _external_scripts(),
                         ids=[s.rsplit('/', 2)[-2] for s, _ in _external_scripts()])
class TestEveryExternalScriptIsPinned:
    def test_has_an_integrity_attribute(self, src, attrs):
        assert 'integrity=' in attrs, (
            f'{src} has no integrity attribute — a compromised CDN could serve '
            f'anything into the wizard'
        )

    def test_uses_a_supported_strong_hash(self, src, attrs):
        value = re.search(r'integrity="([^"]+)"', attrs).group(1)
        assert value.startswith(('sha384-', 'sha512-')), (
            f'{src}: {value[:12]}… — use sha384 or sha512, not sha256'
        )

    def test_has_crossorigin_anonymous(self, src, attrs):
        """Without it the response is opaque and the browser cannot verify the
        hash — the integrity attribute is then silently decorative."""
        assert 'crossorigin="anonymous"' in attrs, f'{src} lacks crossorigin'

    def test_the_url_is_version_pinned(self, src, attrs):
        """A floating URL (…/latest/…) would change under a fixed hash and break
        the page on the vendor's schedule rather than ours."""
        assert '/latest/' not in src, f'{src} is not version-pinned'
        assert re.search(r'/\d+\.\d+', src), f'{src} has no version in the path'


@pytest.mark.network
@pytest.mark.parametrize('src,attrs', _external_scripts(),
                         ids=[s.rsplit('/', 2)[-2] for s, _ in _external_scripts()])
def test_the_hash_matches_what_the_cdn_serves(src, attrs):
    """The substance: download the file and confirm the committed hash is right.

    Marked `network` so it is opt-in — the default suite must not depend on
    cdnjs being reachable. Run with: pytest -m network
    """
    declared = re.search(r'integrity="([^"]+)"', attrs).group(1)
    algo, _, expected_b64 = declared.partition('-')

    request = urllib.request.Request(src, headers={'User-Agent': 'grepxcel-sri-test'})
    with urllib.request.urlopen(request, timeout=60) as response:  # nosec B310 - https, from our own template
        body = response.read()

    digest = {'sha384': hashlib.sha384, 'sha512': hashlib.sha512}[algo](body).digest()
    actual = base64.b64encode(digest).decode()

    assert actual == expected_b64, (
        f'{src}\n'
        f'  committed: {algo}-{expected_b64}\n'
        f'  served   : {algo}-{actual}\n'
        f'The pinned hash does not match what the CDN serves. Either the version '
        f'was bumped without recomputing the hash (the browser is now silently '
        f'refusing to run this library), or the file changed under a fixed '
        f'version — which would be a supply-chain event worth investigating.'
    )
