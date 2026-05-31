#!/usr/bin/env python3
"""
install_cuda.py — Install the NVIDIA CUDA Toolkit on Linux.

This script downloads the official CUDA runfile installer from NVIDIA,
verifies it (file size + SHA256), installs the toolkit (compiler + libraries
only — NOT the driver), and optionally updates your shell profile.

It runs with system Python 3 — no virtual environment required.

Usage:
    python3 scripts/install_cuda.py [OPTIONS]

    --cuda-version VERSION     CUDA Toolkit version to install (default: 12.6.0)
    --download-dir  PATH       Where to save the runfile (default: /tmp)
    --install-prefix PATH      Parent directory for the install (default: /usr/local)
    --bashrc-file   FILE       Shell profile to update (default: ~/.bashrc)
    --skip-bashrc              Do not modify the shell profile
    --keep-runfile             Do not delete the runfile after installation
    --yes / -y                 Auto-confirm all prompts (non-interactive / CI mode)
    --dry-run                  Show every action that would be taken, make no changes
    --help / -h                Show this message and exit

Examples:
    # Interactive (recommended for first-time setup):
    python3 scripts/install_cuda.py

    # Non-interactive CI / headless install:
    python3 scripts/install_cuda.py --yes

    # Different CUDA version, custom download dir:
    python3 scripts/install_cuda.py --cuda-version 12.5.0 --download-dir ~/downloads

    # Only show what would happen — no changes:
    python3 scripts/install_cuda.py --dry-run
"""

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import urllib.request
from pathlib import Path


# ── Known CUDA releases ────────────────────────────────────────────────────────
# Maps version → (min_driver_in_url, exact_bytes_from_cdn, sha256_or_None)
#
# exact_bytes: verified via HEAD request against developer.download.nvidia.com
# sha256: compute locally after download with `sha256sum <file>`, then add here.
#         None means "not yet hardcoded — will be computed and shown post-download".
#
# To add a new release:
#   1. Run: python3 tmp.local/fetch_cuda_checksums.py   (gets exact byte count)
#   2. Download the file, run: sha256sum <runfile>       (gets SHA256)
#   3. Add the entry below.

KNOWN_RELEASES = {
    '12.6.1': ('560.35.03', 4345714567, None),
    '12.6.0': ('560.28.03', 4333105923, None),
    '12.5.0': ('555.42.02', 4294677299, None),
    '12.4.0': ('550.54.14', 4454353277, None),
}
DEFAULT_CUDA_VERSION = '12.6.1'


# ── Styling helpers ────────────────────────────────────────────────────────────

BOLD   = '\033[1m'
GREEN  = '\033[32m'
YELLOW = '\033[33m'
RED    = '\033[31m'
CYAN   = '\033[36m'
RESET  = '\033[0m'

def _h(text):      print(f'\n{BOLD}{CYAN}{"─"*60}{RESET}\n{BOLD}{text}{RESET}')
def _ok(text):     print(f'  {GREEN}✓{RESET}  {text}')
def _warn(text):   print(f'  {YELLOW}⚠{RESET}  {text}')
def _err(text):    print(f'  {RED}✗{RESET}  {text}', file=sys.stderr)
def _info(text):   print(f'     {text}')
def _step(n, t, text): print(f'\n{BOLD}[{n}/{t}]{RESET} {text}')
def _gb(n):        return f'{n / 1e9:.2f} GB ({n:,} bytes)'


# ── Argument parser ────────────────────────────────────────────────────────────

def _build_parser():
    p = argparse.ArgumentParser(
        prog='install_cuda.py',
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument('--cuda-version', default=DEFAULT_CUDA_VERSION, metavar='VERSION',
                   help=f'CUDA Toolkit version (default: {DEFAULT_CUDA_VERSION}). '
                        f'Known: {", ".join(KNOWN_RELEASES)}')
    p.add_argument('--download-dir', default='/tmp', metavar='PATH',
                   help='Directory to save the runfile (default: /tmp)')
    p.add_argument('--install-prefix', default='/usr/local', metavar='PATH',
                   help='Parent directory for the CUDA install (default: /usr/local)')
    p.add_argument('--bashrc-file', default=str(Path.home() / '.bashrc'), metavar='FILE',
                   help='Shell profile to update (default: ~/.bashrc)')
    p.add_argument('--skip-bashrc', action='store_true',
                   help='Do not modify the shell profile')
    p.add_argument('--keep-runfile', action='store_true',
                   help='Do not delete the runfile after installation')
    p.add_argument('-y', '--yes', action='store_true',
                   help='Auto-confirm all prompts (non-interactive / CI mode)')
    p.add_argument('--dry-run', action='store_true',
                   help='Print every action that would be taken; make no changes')
    return p


# ── Confirmation ──────────────────────────────────────────────────────────────

def _confirm(question, args):
    if args.yes or args.dry_run:
        print(f'  {YELLOW}[auto-confirm]{RESET} {question}')
        return True
    answer = input(f'\n  {BOLD}{question}{RESET} [y/N] ').strip().lower()
    return answer in ('y', 'yes')


# ── File integrity ────────────────────────────────────────────────────────────

def _server_size(url):
    """Return Content-Length from server via HEAD request, or None on failure."""
    try:
        req = urllib.request.Request(url, method='HEAD')
        req.add_header('User-Agent', 'Mozilla/5.0')
        with urllib.request.urlopen(req, timeout=15) as r:
            val = r.headers.get('Content-Length')
            return int(val) if val else None
    except Exception as e:
        return None


def _sha256(path, chunk=1 << 20):
    """Compute SHA256 of a file, printing progress every 200 MB."""
    h = hashlib.sha256()
    total = path.stat().st_size
    done  = 0
    with open(path, 'rb') as f:
        while True:
            block = f.read(chunk)
            if not block:
                break
            h.update(block)
            done += len(block)
            pct = done * 100 // total
            print(f'  Computing SHA256 ... {pct:3d}%\r', end='', flush=True)
    print()
    return h.hexdigest()


def _check_local_file(path, expected_bytes, known_sha256, args):
    """
    Inspect an already-downloaded file. Returns one of:
      'missing'        — file does not exist
      'valid'          — size + hash (if known) both match
      'size_mismatch'  — file exists but wrong size
      'hash_mismatch'  — size OK but SHA256 differs from known value
      'unverified'     — size matches, no known hash to check against
    """
    if not path.exists():
        return 'missing'

    local_size = path.stat().st_size
    _info(f'Local file  : {path}')
    _info(f'Local size  : {_gb(local_size)}')
    _info(f'Expected    : {_gb(expected_bytes)}')

    if local_size != expected_bytes:
        _warn(f'Size mismatch: local {local_size:,} ≠ expected {expected_bytes:,}')
        return 'size_mismatch'

    _ok(f'File size matches ({_gb(local_size)})')

    if known_sha256:
        _info(f'Verifying SHA256 (known: {known_sha256[:16]}...)')
        if args.dry_run:
            _info('[dry-run] SHA256 verification skipped')
            return 'valid'
        actual = _sha256(path)
        if actual == known_sha256:
            _ok(f'SHA256 verified: {actual}')
            return 'valid'
        else:
            _err(f'SHA256 MISMATCH')
            _info(f'  Expected : {known_sha256}')
            _info(f'  Got      : {actual}')
            return 'hash_mismatch'

    _warn('No hardcoded SHA256 for this version — size check only.')
    _info('(Add the hash to KNOWN_RELEASES in scripts/install_cuda.py after verifying.)')
    return 'unverified'


def _verify_after_download(path, expected_bytes, known_sha256, args):
    """Full integrity check immediately after downloading. Exits on failure."""
    _info(f'Verifying downloaded file...')

    local_size = path.stat().st_size
    _info(f'Downloaded size : {_gb(local_size)}')
    _info(f'Expected size   : {_gb(expected_bytes)}')

    if local_size != expected_bytes:
        _err(f'Size mismatch after download. File may be corrupt or truncated.')
        _info('Try re-running the script to re-download.')
        sys.exit(1)
    _ok('File size matches.')

    if not args.dry_run:
        actual_sha256 = _sha256(path)
        _info(f'SHA256: {actual_sha256}')
        if known_sha256:
            if actual_sha256 == known_sha256:
                _ok('SHA256 verified against known good value.')
            else:
                _err('SHA256 mismatch — file is corrupt or tampered.')
                _info(f'  Expected : {known_sha256}')
                _info(f'  Got      : {actual_sha256}')
                _info(f'Deleting corrupt file: {path}')
                path.unlink(missing_ok=True)
                sys.exit(1)
        else:
            _warn('No hardcoded SHA256 for this version.')
            _info(f'To add it, copy the line below into KNOWN_RELEASES in scripts/install_cuda.py:')
            _info(f"  '{args.cuda_version}': ('{KNOWN_RELEASES[args.cuda_version][0]}', "
                  f"{expected_bytes}, '{actual_sha256}'),")


# ── Pre-flight checks ─────────────────────────────────────────────────────────

def _preflight(args):
    _h('Pre-flight checks')
    findings = {}

    if sys.platform != 'linux':
        _err(f'Linux only (detected: {sys.platform}).'); sys.exit(1)
    _ok('Platform: Linux')

    # NVIDIA driver
    try:
        r = subprocess.run(
            ['nvidia-smi', '--query-gpu=name,driver_version,compute_cap', '--format=csv,noheader'],
            capture_output=True, text=True, timeout=5,
        )
        if r.returncode == 0 and r.stdout.strip():
            findings['gpus'] = r.stdout.strip().splitlines()
            for g in findings['gpus']:
                _ok(f'NVIDIA GPU: {g.strip()}')
        else:
            _err('nvidia-smi ran but found no GPU.'); sys.exit(1)
    except FileNotFoundError:
        _err('nvidia-smi not found — install the NVIDIA display driver first.'); sys.exit(1)

    # nvcc
    if shutil.which('nvcc'):
        r = subprocess.run(['nvcc', '--version'], capture_output=True, text=True)
        _warn(f'nvcc already present at {shutil.which("nvcc")} — will be overwritten.')
        findings['nvcc_exists'] = True
    else:
        _ok('nvcc: not yet installed (will be installed by this script)')
        findings['nvcc_exists'] = False

    # Existing installs
    prefix   = Path(args.install_prefix)
    existing = sorted(prefix.glob('cuda-*'))
    if existing:
        _warn('Existing CUDA installs found:')
        for e in existing: _info(str(e))
    else:
        _ok(f'No existing CUDA installs under {prefix}')
    findings['existing_cuda'] = existing

    # Build tools
    if shutil.which('gcc') and shutil.which('g++'):
        r = subprocess.run(['gcc', '--version'], capture_output=True, text=True)
        _ok(f'gcc/g++: {r.stdout.splitlines()[0]}')
    else:
        _err('gcc/g++ not found.  Fix: sudo apt install build-essential'); sys.exit(1)

    if shutil.which('wget'):
        _ok('wget: available')
    else:
        _err('wget not found.  Fix: sudo apt install wget'); sys.exit(1)

    # libxml2.so.2 — required by the CUDA runfile installer.
    # Ubuntu 24.04+ ships libxml2-16 (libxml2.so.16) instead of libxml2 (libxml2.so.2).
    # The CUDA installer was built against the old .so.2 name — we fix this with a symlink.
    r = subprocess.run(['ldconfig', '-p'], capture_output=True, text=True)
    if 'libxml2.so.2' in r.stdout:
        _ok('libxml2.so.2: found (required by CUDA installer)')
    else:
        # Find whatever libxml2.so.* is available
        import glob
        candidates = glob.glob('/usr/lib/x86_64-linux-gnu/libxml2.so.*') + \
                     glob.glob('/usr/lib/libxml2.so.*')
        candidates = [c for c in candidates if not c.endswith('.so')]
        real_lib   = next((c for c in sorted(candidates, reverse=True)
                           if not c.endswith('.so')), None)

        if real_lib:
            _warn(f'libxml2.so.2 not found, but {real_lib} is present.')
            _info('Ubuntu 24.04+ renamed the library. A compatibility symlink fixes this.')
            symlink_path = Path(real_lib).parent / 'libxml2.so.2'
            _info(f'Will create: sudo ln -s {real_lib} {symlink_path}')
            answer = input(
                f'\n  {BOLD}Create compatibility symlink '
                f'(sudo ln -s {Path(real_lib).name} {symlink_path})?{RESET} [y/N] '
            ).strip().lower() \
                if not (getattr(args, 'yes', False) or getattr(args, 'dry_run', False)) \
                else 'y'
            if answer in ('y', 'yes'):
                if getattr(args, 'dry_run', False):
                    _info(f'[dry-run] would run: sudo ln -s {real_lib} {symlink_path}')
                else:
                    r2 = subprocess.run(
                        ['sudo', 'ln', '-sf', real_lib, str(symlink_path)]
                    )
                    if r2.returncode == 0:
                        _ok(f'Symlink created: {symlink_path} → {real_lib}')
                    else:
                        _err('Failed to create symlink. Try manually:')
                        _info(f'  sudo ln -sf {real_lib} {symlink_path}')
                        sys.exit(1)
            else:
                _err('libxml2.so.2 is required by the CUDA installer.')
                sys.exit(1)
        else:
            _warn('libxml2 not found at all — need to install it.')
            # Try package names in order (libxml2-16 for Ubuntu 24.04+, libxml2 for older)
            pkg = None
            for candidate_pkg in ('libxml2-16', 'libxml2'):
                r2 = subprocess.run(
                    ['apt-cache', 'show', candidate_pkg],
                    capture_output=True,
                )
                if r2.returncode == 0:
                    pkg = candidate_pkg
                    break
            if pkg is None:
                _err('Cannot find a libxml2 package. Install it manually and re-run.')
                sys.exit(1)
            _info(f'Will install: sudo apt install -y {pkg}')
            answer = input(
                f'\n  {BOLD}Install {pkg} now via apt?{RESET} [y/N] '
            ).strip().lower() \
                if not (getattr(args, 'yes', False) or getattr(args, 'dry_run', False)) \
                else 'y'
            if answer in ('y', 'yes'):
                if getattr(args, 'dry_run', False):
                    _info(f'[dry-run] would run: sudo apt install -y {pkg}')
                else:
                    r2 = subprocess.run(['sudo', 'apt', 'install', '-y', pkg])
                    if r2.returncode != 0:
                        _err(f'Failed to install {pkg}.')
                        sys.exit(1)
                    _ok(f'{pkg} installed.')
            else:
                _err('libxml2 is required. Install it manually and re-run.')
                sys.exit(1)

    # sudo
    r = subprocess.run(['sudo', '-n', 'true'], capture_output=True)
    if r.returncode == 0:
        _ok('sudo: available (cached credentials)')
        findings['sudo_cached'] = True
    else:
        _warn('sudo: will prompt for password during installation')
        findings['sudo_cached'] = False

    # Disk space
    dl_dir  = Path(args.download_dir)
    in_dir  = Path(args.install_prefix)
    dl_dir.mkdir(parents=True, exist_ok=True)
    free_dl  = os.statvfs(dl_dir)
    free_ins = os.statvfs(in_dir)
    gb_dl    = free_dl.f_bavail  * free_dl.f_frsize  // (1024 ** 3)
    gb_ins   = free_ins.f_bavail * free_ins.f_frsize // (1024 ** 3)
    _ok(f'Free disk in {dl_dir}  : {gb_dl} GB')
    _ok(f'Free disk in {in_dir}  : {gb_ins} GB')
    if gb_dl < 5:
        _err(f'Need ≥5 GB free in {dl_dir} for the runfile.'); sys.exit(1)
    if gb_ins < 5:
        _err(f'Need ≥5 GB free in {in_dir} for the toolkit.'); sys.exit(1)

    # Shell profile
    if not args.skip_bashrc:
        bashrc = Path(args.bashrc_file)
        if bashrc.exists():
            content = bashrc.read_text()
            if 'cuda' in content.lower():
                _warn(f'{bashrc}: already contains CUDA references (will append).')
                findings['bashrc_has_cuda'] = True
            else:
                _ok(f'Shell profile: {bashrc} — no existing CUDA entries')
                findings['bashrc_has_cuda'] = False
        else:
            _warn(f'Shell profile {bashrc} does not exist — will be created.')
            findings['bashrc_has_cuda'] = False

    return findings


# ── Plan ──────────────────────────────────────────────────────────────────────

def _print_plan(args, runfile_url, runfile_path, cuda_dir,
                driver_ver, expected_bytes, known_sha256):
    _h('Installation plan')
    print()
    _info(f'CUDA version    : {args.cuda_version}')
    _info(f'Download URL    : {runfile_url}')
    _info(f'Download to     : {runfile_path}')
    _info(f'Expected size   : {_gb(expected_bytes)}')
    _info(f'SHA256 check    : {"yes — " + known_sha256[:24] + "..." if known_sha256 else "size-only (hash not yet hardcoded)"}')
    _info(f'Install into    : {cuda_dir}')
    _info(f'Min driver ver  : {driver_ver}  (yours is newer — OK)')
    print()
    print(f'  {BOLD}Changes that will be made to this machine:{RESET}')
    print()
    i = 1
    changes = []
    changes.append(f'{i}. DOWNLOAD  {_gb(expected_bytes)} → {runfile_path}'); i += 1
    changes.append(f'{i}. VERIFY    file size ({expected_bytes:,} bytes)'
                   + (f' + SHA256' if known_sha256 else ' (no hardcoded hash)')); i += 1
    changes.append(f'{i}. RUN       sudo sh {runfile_path} --toolkit --silent --override'); i += 1
    changes.append(f'{i}. CREATES   {cuda_dir}/  (~4 GB compiler + headers + libraries)'); i += 1
    if not args.skip_bashrc:
        changes.append(f'{i}. APPEND    export PATH={cuda_dir}/bin:$PATH  →  {args.bashrc_file}'); i += 1
        changes.append(f'{i}. APPEND    export LD_LIBRARY_PATH={cuda_dir}/lib64:$LD_LIBRARY_PATH  →  {args.bashrc_file}'); i += 1
    if not args.keep_runfile:
        changes.append(f'{i}. DELETE    {runfile_path}  (after successful install)')
    for c in changes:
        _info(c)
    print()
    if args.dry_run:
        print(f'  {YELLOW}DRY-RUN mode — no changes will be made.{RESET}')


# ── Shell runner ──────────────────────────────────────────────────────────────

def _run(cmd, args, check=True):
    if args.dry_run:
        print(f'  {YELLOW}[dry-run]{RESET} would run: {" ".join(str(c) for c in cmd)}')
        return subprocess.CompletedProcess(cmd, 0)
    print(f'  $ {" ".join(str(c) for c in cmd)}', flush=True)
    result = subprocess.run(cmd)
    if check and result.returncode != 0:
        _err(f'Command failed with exit code {result.returncode}')
        sys.exit(result.returncode)
    return result


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = _build_parser().parse_args()

    print()
    print(f'{BOLD}{"═"*60}{RESET}')
    print(f'{BOLD}  grepxcel — CUDA Toolkit Installer{RESET}')
    print(f'{BOLD}{"═"*60}{RESET}')
    print()
    print('  Installs the NVIDIA CUDA Toolkit (toolkit only).')
    print('  Your existing NVIDIA display driver is NOT modified.')
    if args.dry_run:
        print(f'\n  {YELLOW}DRY-RUN — nothing will be changed.{RESET}')
    if args.yes:
        print(f'\n  {YELLOW}Non-interactive mode (--yes).{RESET}')

    # ── Resolve release ────────────────────────────────────────────────────────
    ver = args.cuda_version
    if ver not in KNOWN_RELEASES:
        _err(f'Unknown CUDA version: {ver}')
        _info(f'Known: {", ".join(KNOWN_RELEASES)}')
        sys.exit(1)

    driver_ver, expected_bytes, known_sha256 = KNOWN_RELEASES[ver]
    ver_short    = '.'.join(ver.split('.')[:2])
    runfile_name = f'cuda_{ver}_{driver_ver}_linux.run'
    runfile_url  = (
        f'https://developer.download.nvidia.com/compute/cuda/'
        f'{ver}/local_installers/{runfile_name}'
    )
    runfile_path = Path(args.download_dir) / runfile_name
    cuda_dir     = Path(args.install_prefix) / f'cuda-{ver_short}'

    total_steps  = 4 if args.skip_bashrc else 5

    # ── Pre-flight ─────────────────────────────────────────────────────────────
    _preflight(args)

    # ── Check server-side file size ────────────────────────────────────────────
    _h('Checking download source')
    _info(f'Contacting NVIDIA CDN...')
    server_bytes = _server_size(runfile_url)
    if server_bytes is None:
        _warn('Could not reach NVIDIA CDN (HEAD request failed).')
        _info('Proceeding with hardcoded size from KNOWN_RELEASES.')
        server_bytes = expected_bytes
    else:
        _ok(f'Server reports file size: {_gb(server_bytes)}')
        if server_bytes != expected_bytes:
            _warn(f'Server size ({server_bytes:,}) differs from hardcoded value '
                  f'({expected_bytes:,}).')
            _info('The file on NVIDIA\'s CDN may have been updated.')
            _info('Using server-reported size for verification.')
            expected_bytes = server_bytes

    # ── Plan ──────────────────────────────────────────────────────────────────
    _print_plan(args, runfile_url, runfile_path, cuda_dir,
                driver_ver, expected_bytes, known_sha256)

    if not _confirm('Proceed with the installation plan above?', args):
        print('\n  Aborted. No changes were made.')
        sys.exit(0)

    # ── Step 1: Download (or reuse) ────────────────────────────────────────────
    _step(1, total_steps, f'Download CUDA {ver} runfile ({_gb(expected_bytes)})')

    skip_download = False
    if runfile_path.exists():
        _info('File already exists locally — checking integrity...')
        status = _check_local_file(runfile_path, expected_bytes, known_sha256, args)
        if status in ('valid', 'unverified'):
            _ok('Existing file passes integrity check.')
            if _confirm('Skip re-download and use the existing file?', args):
                skip_download = True
            else:
                _info('Will re-download.')
        else:
            _warn(f'Existing file failed integrity check ({status}) — will re-download.')

    if not skip_download:
        _info(f'Source : {runfile_url}')
        _info(f'Dest   : {runfile_path}')
        if not _confirm(f'Download {_gb(expected_bytes)} from developer.download.nvidia.com?', args):
            print('\n  Download skipped. Cannot continue without the runfile.')
            sys.exit(1)
        _run(['wget', '--progress=bar:force', '-O', str(runfile_path), runfile_url], args)

        # Post-download integrity check
        _step(1, total_steps, 'Verify downloaded file')
        if not args.dry_run:
            _verify_after_download(runfile_path, expected_bytes, known_sha256, args)
        else:
            _info('[dry-run] Integrity verification skipped.')
    else:
        _ok('Using existing file — download skipped.')

    # ── Step 2: Install ────────────────────────────────────────────────────────
    _step(2, total_steps, 'Install CUDA Toolkit (requires sudo)')
    _info('Installer flags:')
    _info('  --toolkit   installs compiler + libraries only, NOT the driver')
    _info('  --silent    non-interactive')
    _info('  --override  allows install even if driver source differs')
    _info(f'Destination : {cuda_dir}')

    if not _confirm(
        f'Run the CUDA installer (sudo) — installs ~4 GB into {cuda_dir}?', args
    ):
        print('\n  Installation skipped.')
        sys.exit(0)

    _run(['sudo', 'sh', str(runfile_path), '--toolkit', '--silent', '--override'], args)
    _ok(f'CUDA Toolkit installed at {cuda_dir}')

    # Verify nvcc
    if not args.dry_run:
        _step(2, total_steps, 'Verify installation')
        nvcc_path = cuda_dir / 'bin' / 'nvcc'
        if nvcc_path.exists():
            r = subprocess.run([str(nvcc_path), '--version'], capture_output=True, text=True)
            _ok(f'nvcc: {r.stdout.splitlines()[0] if r.stdout else "found"}')
        else:
            _warn(f'nvcc not found at {nvcc_path} — installation may have failed.')

    # ── Step 3: Shell profile ──────────────────────────────────────────────────
    if not args.skip_bashrc:
        _step(3, total_steps, f'Update shell profile: {args.bashrc_file}')
        path_line  = f'export PATH={cuda_dir}/bin:$PATH'
        ldlib_line = f'export LD_LIBRARY_PATH={cuda_dir}/lib64:$LD_LIBRARY_PATH'
        marker     = f'# Added by grepxcel install_cuda.py (CUDA {ver})'
        block      = f'\n{marker}\n{path_line}\n{ldlib_line}\n'

        _info(f'Will append to {args.bashrc_file}:')
        _info(f'  {marker}')
        _info(f'  {path_line}')
        _info(f'  {ldlib_line}')

        if not _confirm(f'Append the two export lines to {args.bashrc_file}?', args):
            _warn('Shell profile not updated. Add these lines manually:')
            _info(path_line)
            _info(ldlib_line)
        else:
            if not args.dry_run:
                with open(args.bashrc_file, 'a') as f:
                    f.write(block)
            _ok(f'{args.bashrc_file} updated.')
            _warn(f'Run: source {args.bashrc_file}  (or open a new terminal)')

    # ── Step 4: Cleanup ────────────────────────────────────────────────────────
    _step(total_steps, total_steps, 'Cleanup')
    if not args.keep_runfile:
        _info(f'Runfile: {runfile_path}  ({_gb(runfile_path.stat().st_size) if runfile_path.exists() else "n/a"})')
        if not _confirm(f'Delete the runfile to free space?', args):
            _warn('Runfile kept — delete it manually when no longer needed.')
        else:
            if not args.dry_run:
                runfile_path.unlink(missing_ok=True)
            _ok('Runfile deleted.')
    else:
        _ok(f'Runfile kept at {runfile_path}  (--keep-runfile).')

    # ── Done ──────────────────────────────────────────────────────────────────
    print()
    print(f'{BOLD}{"═"*60}{RESET}')
    if args.dry_run:
        print(f'{BOLD}  Dry run complete. No changes were made.{RESET}')
    else:
        print(f'{BOLD}{GREEN}  CUDA {ver} Toolkit installed successfully.{RESET}')
        print()
        print(f'  Next steps:')
        print(f'    1. Reload shell   :  source {args.bashrc_file}')
        print(f'    2. Verify nvcc    :  nvcc --version')
        print(f'    3. Enable GPU     :  python3 scripts/install_llm_deps.py')
    print(f'{BOLD}{"═"*60}{RESET}')
    print()


if __name__ == '__main__':
    main()
