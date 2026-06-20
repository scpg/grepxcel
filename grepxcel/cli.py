"""
Command-line interface for grepxcel.

Usage:
    grepxcel extract -p pattern.xlsx data.xlsx
    grepxcel extract -p pattern.xlsx data1.xlsx data2.xlsx data3.xlsx
    grepxcel extract -p pattern.xlsx data.xlsx -v --output results/
    grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log
    grepxcel draft data.xlsx -o draft_pattern.xlsx
    grepxcel draft data.xlsx -v
"""

import argparse
import datetime
import json
import os
import sys

from .color import colorize_marks, should_color
from .engine import Engine
from .logger import Logger, VerbosityLevel


# ── JSON serialisation ────────────────────────────────────────────────────────

def _json_default(obj):
    """Serialise types that json.dump does not handle natively.

    Excel cells surface as several datetime flavours: dates/datetimes and
    time-of-day all have ``.isoformat()``; durations ([h]:mm cells) come through
    as ``timedelta``, which has no isoformat, so render it as ``str`` (e.g.
    ``"8:30:00"``).
    """
    if isinstance(obj, (datetime.date, datetime.datetime, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, datetime.timedelta):
        return str(obj)
    raise TypeError(f'Type {type(obj).__name__} is not JSON serialisable')


# ── Shared argument helpers ───────────────────────────────────────────────────

def _add_security_args(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        '--max-size',
        type=float, default=5, metavar='MB',
        help='Maximum compressed file size in MB (default: 5)',
    )
    p.add_argument(
        '--max-uncompressed',
        type=float, default=50, metavar='MB',
        help='Maximum uncompressed ZIP content in MB (default: 50)',
    )


def _add_sheet_arg(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        '--sheet',
        metavar='NAME_OR_INDEX',
        help='Sheet to use: name (e.g. Sheet2) or 0-based index (default: active sheet)',
    )


def _resolve_sheet(args) -> str | None:
    """Return the --sheet value unchanged (string or None).

    Resolution happens downstream: a string is matched as a sheet NAME first and
    falls back to a 0-based index only when no sheet has that name. This lets a
    numerically-named sheet (e.g. '2025') be selected by name, while '--sheet 0'
    still works as an index. Converting to int here would wrongly turn the name
    '2025' into index 2025.
    """
    return getattr(args, 'sheet', None)


# ── Argument parser ───────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='grepxcel',
        description='Extract structured data from Excel files using a pattern.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
commands:
  extract            Extract data from Excel files using a pattern file
  validate-pattern   Check a pattern file is valid to use (no extraction)
  draft              Use a local LLM to draft a starter pattern file
  docs               Write a pattern-format reference xlsx (pattern-reference.xlsx)
  lint               Inspect an Excel file for potential extraction issues
  schema             Generate a JSON Schema from a pattern file
  generate-skill     Write an AI-agent skill doc (Claude / AGENTS.md)
  doctor             Check the environment is ready (deps, keys, model, proxy/TLS)

Run 'grepxcel <command> --help' for per-command options.
        """,
    )
    from . import __version__
    p.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}',
    )
    sub = p.add_subparsers(dest='command', metavar='COMMAND')
    sub.required = True

    _add_extract_subparser(sub)
    _add_validate_subparser(sub)
    _add_draft_subparser(sub)
    _add_docs_subparser(sub)
    _add_lint_subparser(sub)
    _add_schema_subparser(sub)
    _add_skill_subparser(sub)
    _add_doctor_subparser(sub)
    return p


def _add_validate_subparser(sub) -> None:
    p = sub.add_parser(
        'validate-pattern',
        help='Check a pattern file (.xlsx or .csv) is valid to use',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Parses the pattern with the same rules extraction uses (structure, types,
multiplicities, regex safety, comments) and also flags an empty extraction
sequence and references to undefined fields. Exits non-zero if any file is
invalid.

examples:
  grepxcel validate-pattern pattern.xlsx
  grepxcel validate-pattern pattern.csv -v        # + parsed fields & steps
  grepxcel validate-pattern a.xlsx b.csv          # validate several
        """,
    )
    p.add_argument('files', nargs='+', metavar='FILE',
                   help='Pattern file(s) to validate (.xlsx or .csv)')
    p.add_argument('-v', '--verbose', action='store_true',
                   help='Print the parsed config, fields, and extraction sequence')


def _add_lint_subparser(sub) -> None:
    p = sub.add_parser(
        'lint',
        help='Inspect an Excel file for potential extraction issues',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Checks an Excel data file before extraction: format, encryption/IRM,
sheet dimensions (declared vs real extent), merged cells, formula cells,
and known corporate-environment issues. Reports ✓/⚠/✗/ℹ.

examples:
  grepxcel lint data.xlsx
  grepxcel lint jan.xlsx feb.xlsx        # lint several files
        """,
    )
    p.add_argument('files', nargs='+', metavar='FILE',
                   help='Excel file(s) to inspect (.xlsx)')


def _add_schema_subparser(sub) -> None:
    p = sub.add_parser(
        'schema',
        help='Generate a JSON Schema from a pattern file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Generates a JSON Schema (draft 2020-12) that describes the extraction
output for the given pattern. Use it to validate extracted JSON with
any standard JSON Schema validator.

examples:
  grepxcel schema pattern.xlsx
  grepxcel schema pattern.xlsx -o schema.json
  grepxcel schema pattern.csv
        """,
    )
    p.add_argument('files', nargs='+', metavar='PATTERN',
                   help='Pattern file(s) to generate schema for (.xlsx or .csv)')
    p.add_argument('-o', '--output', metavar='FILE',
                   help='Write schema to file (default: stdout)')


def _add_skill_subparser(sub) -> None:
    p = sub.add_parser(
        'generate-skill',
        help='Write an AI-agent skill doc (Claude / AGENTS.md)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Writes a markdown 'skill' doc that teaches an AI agent how and when to use
grepxcel. The command reference is introspected from the live CLI; the prose is
curated. v1 targets: claude (SKILL.md) and agents-md (AGENTS.md).

examples:
  grepxcel generate-skill                       # Claude SKILL.md to stdout
  grepxcel generate-skill -o SKILL.md
  grepxcel generate-skill --target agents-md -o AGENTS.md
        """,
    )
    p.add_argument('--target', choices=['claude', 'agents-md'], default='claude',
                   help='Skill format to emit (default: claude)')
    p.add_argument('-o', '--output', metavar='FILE',
                   help='Write the skill doc to FILE (default: stdout)')


def _add_doctor_subparser(sub) -> None:
    p = sub.add_parser(
        'doctor',
        help='Check the environment is ready (deps, API keys, model, proxy/TLS)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  grepxcel doctor              # check everything (extract + draft + proxy/TLS)
  grepxcel doctor extract      # only what 'extract' needs (offline)
  grepxcel doctor draft        # only what 'draft' needs (local + cloud + proxy)

Exits non-zero if the selected area has a blocking (✗) problem.
        """,
    )
    p.add_argument(
        'area',
        nargs='?',
        choices=['extract', 'draft', 'all'],
        default='all',
        help="Which area to check: extract, draft, or all (default: all)",
    )
    p.add_argument(
        '--no-probe',
        action='store_true',
        help='Skip the live TLS handshake probe (offline / faster)',
    )
    p.add_argument(
        '--strict-env',
        action='store_true',
        help='Refuse a .env from outside the current project (also '
             'GREPXCEL_STRICT_ENV); the per-user config-dir .env stays allowed',
    )


def _add_extract_subparser(sub) -> None:
    p = sub.add_parser(
        'extract',
        help='Extract data from Excel files using a pattern file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
verbosity:
  (none)  warnings + summary on stderr, extracted JSON on stdout
  -v      + per-field trace: 'field ← B1 = value ✓/✗' for cells and table data
  -vv     + every anchor probe and rejection reason
  -d      same as -vv (debug mode)

examples:
  grepxcel extract -p pattern.xlsx report.xlsx
  grepxcel extract -p pattern.xlsx jan.xlsx feb.xlsx mar.xlsx
  grepxcel extract -p pattern.xlsx data.xlsx -v
  grepxcel extract -p pattern.xlsx data.xlsx --output results/
  grepxcel extract -p pattern.xlsx data.xlsx --log logs/run.log --output results/
        """,
    )
    p.add_argument(
        '-p', '--pattern',
        required=True,
        metavar='FILE',
        help='Pattern Excel file that describes the layout to extract',
    )
    p.add_argument(
        'files',
        nargs='+',
        metavar='FILE',
        help='One or more data Excel files to process',
    )
    p.add_argument(
        '-v', '--verbose',
        action='count', default=0,
        help='Increase verbosity (-v: step-by-step, -vv: anchor probes)',
    )
    p.add_argument(
        '-d', '--debug',
        action='store_true',
        help='Debug output (anchor probes and rejections); equivalent to -vv',
    )
    p.add_argument(
        '-o', '--output',
        metavar='DIR',
        help='Write extracted data as JSON into DIR  (<datafile-stem>.json per file)',
    )
    p.add_argument(
        '-l', '--log',
        metavar='FILE',
        help='Append a log to FILE (text by default; NDJSON with --log-format json)',
    )
    p.add_argument(
        '--log-format',
        choices=['text', 'json'], default='text',
        help='Log file format: text (human, shows values) or json (NDJSON for '
             'SIEM/cloud — never contains extracted cell values). Default: text',
    )
    p.add_argument(
        '--max-cell-len',
        type=int, default=1000, metavar='CHARS',
        help='Maximum cell character length passed to regex matching (default: 1000)',
    )
    p.add_argument(
        '--max-rows',
        type=int, default=2048, metavar='N',
        help='Maximum rows in the data sheet (default: 2048; Excel max is 1048576, '
             'untested above the default)',
    )
    p.add_argument(
        '--max-columns',
        type=int, default=1024, metavar='N',
        help='Maximum columns in the data sheet (default: 1024; Excel max is 16384, '
             'untested above the default)',
    )
    p.add_argument(
        '--format',
        choices=['nested', 'legacy'], default='nested',
        help='Output format: nested (default) or legacy ({"cells":{}, "tables":[]})',
    )
    p.add_argument(
        '--meta',
        action='store_true',
        help='Add a _meta block to the JSON output (run_id, stats, issues) '
             'for pipeline auto-verification',
    )
    sheet_group = p.add_mutually_exclusive_group()
    sheet_group.add_argument(
        '--all-sheets',
        action='store_true',
        help='Process every sheet in the workbook; output is a dict keyed by sheet name',
    )
    sheet_group.add_argument(
        '--sheet',
        metavar='NAME_OR_INDEX',
        help='Sheet to use: name (e.g. Sheet2) or 0-based index (default: active sheet)',
    )
    _add_security_args(p)


def _add_docs_subparser(sub) -> None:
    p = sub.add_parser(
        'docs',
        help='Write a self-documenting pattern-format reference xlsx',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  grepxcel docs
  grepxcel docs -o reference/pattern-reference.xlsx
        """,
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        default='pattern-reference.xlsx',
        help='Output path for the reference file (default: pattern-reference.xlsx)',
    )


def _add_draft_subparser(sub) -> None:
    _draft_args(sub.add_parser(
        'draft',
        help='Draft a starter pattern file for an Excel file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=r"""
'draft' is an OPTIONAL feature. 'extract' and 'docs' work out of the box; the
local 'draft' model needs an optional install:
    pip install 'grepxcel[suggest]'      (local model, runs offline)
    pip install 'grepxcel[draft-cloud]'  (cloud backends, e.g. --backend claude)
If a dependency is missing, grepxcel tells you exactly what to install.

backends:
  local   (default) Run a local GGUF model (Gemma-4-E4B) via llama-cpp-python.
          Needs:  pip install 'grepxcel[suggest]'
          Model is downloaded automatically on first run (~5 GB).
          No data leaves your machine during inference.
  github  Send the Excel structure description to GitHub Models (free with a
          GitHub subscription, quota-limited). Highest draft quality in our eval.
          Needs GITHUB_TOKEN (Models: read) and pip install 'grepxcel[draft-cloud]'.
          Choose a model with --github-model (e.g. openai/gpt-4.1, openai/gpt-4o).
  claude  Send the Excel structure description to the Claude API.
          Needs ANTHROPIC_API_KEY and pip install 'grepxcel[draft-cloud]'.
          The raw file is NOT transmitted — only column types, sample
          values, and labels are sent.
  server  Send the Excel structure description to any OpenAI-compatible API
          server (LM Studio, Ollama, vLLM, text-generation-inference, etc.).
          Needs pip install openai. Default URL: http://localhost:1234/v1
          (override with --server-url or GREPXCEL_SERVER_URL). Model is
          auto-discovered unless --server-model is set. Data stays local
          unless you point --server-url at a remote host.
  gemini  Planned for a future release — not yet available.

Keys are read from a .env file (current dir or any parent) if present.
The cloud backends print the per-call token usage; github also prints the
remaining quota, claude prints the per-call dollar cost.

model cache (local backend):
  Stored once in the per-user cache (platform-appropriate, via platformdirs):
    Linux    ~/.cache/grepxcel/models/   (honors $XDG_CACHE_HOME)
    macOS    ~/Library/Caches/grepxcel/models/
    Windows  %LOCALAPPDATA%\grepxcel\Cache\models\
  Override the location with GREPXCEL_MODEL_DIR (takes precedence).
  grepxcel only downloads when the file is missing — to avoid the HuggingFace
  download you can place the GGUF there yourself (exact filename), from any
  source. Faster/alternative downloads:
    HF_TOKEN=hf_...                 remove the anonymous rate limit (fastest fix)
    HF_ENDPOINT=https://hf-mirror.com   use a mirror
  Integrity: a downloaded model is checksummed (sha256) and re-verified on every
  run; a checksum mismatch (tampering/corruption) aborts. The check is cheap by
  default — it re-hashes only when the file's size/mtime changed. For high
  assurance, GREPXCEL_VERIFY_MODEL=full forces a full re-hash every run. A
  manually-placed file has no recorded checksum and can't be verified — grepxcel
  warns and proceeds on trust. Use --allow-unverified-model (or
  GREPXCEL_ALLOW_UNVERIFIED_MODEL=1) to silence the warning / override a mismatch.

examples:
  grepxcel draft report.xlsx
  grepxcel draft report.xlsx -o my_pattern.xlsx
  grepxcel draft report.xlsx --backend claude
  grepxcel draft report.xlsx --dry-run
        """,
    ))


def _draft_args(p: argparse.ArgumentParser) -> None:
    p.add_argument(
        'file',
        metavar='FILE',
        help='Excel file to analyse',
    )
    p.add_argument(
        '-o', '--output',
        metavar='FILE',
        default='draft_pattern.xlsx',
        help='Write draft pattern to FILE (default: draft_pattern.xlsx)',
    )
    p.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Print the Excel analysis sent to the LLM and model update status',
    )
    p.add_argument(
        '--dry-run',
        action='store_true',
        help='Print the Excel analysis that would be sent to the model, then exit without running inference',
    )
    p.add_argument(
        '--backend',
        choices=['local', 'claude', 'gemini', 'github', 'server'],
        default='local',
        help='Inference backend: local (default, GGUF model), claude (requires '
             'ANTHROPIC_API_KEY), github (GitHub Models, requires GITHUB_TOKEN '
             "with 'Models: read'; use --github-model to pick a model), "
             'server (any OpenAI-compatible server, e.g. LM Studio / Ollama / '
             'vLLM; use --server-url). gemini is planned for a future release.',
    )
    p.add_argument(
        '--github-model',
        default='openai/gpt-4o-mini',
        help="GitHub Models model id when --backend github, e.g. 'openai/gpt-4o', "
             "'meta/llama-3.3-70b-instruct' (default: openai/gpt-4o-mini).",
    )
    p.add_argument(
        '--server-url',
        default=os.environ.get('GREPXCEL_SERVER_URL', 'http://localhost:1234/v1'),
        help='Base URL for --backend server (default: http://localhost:1234/v1). '
             'Also settable via GREPXCEL_SERVER_URL.',
    )
    p.add_argument(
        '--server-model',
        default=os.environ.get('GREPXCEL_SERVER_MODEL'),
        help='Model id for --backend server; if omitted, auto-discovers the '
             'first loaded model via /v1/models. Also via GREPXCEL_SERVER_MODEL.',
    )
    p.add_argument(
        '--server-api-key',
        default=os.environ.get('GREPXCEL_SERVER_API_KEY', 'not-needed'),
        help='API key for --backend server (most local servers ignore this). '
             'Also via GREPXCEL_SERVER_API_KEY.',
    )
    p.add_argument(
        '--allow-unverified-model',
        action='store_true',
        help='Proceed even if the cached local model fails its integrity check '
             '(checksum mismatch, or a manually-placed file with no recorded '
             'checksum). Also settable via GREPXCEL_ALLOW_UNVERIFIED_MODEL=1.',
    )
    p.add_argument(
        '--ca-bundle',
        metavar='FILE',
        default=None,
        help='Path to a corporate CA bundle (.pem) to trust behind a '
             'TLS-inspection proxy (NetSkope/Zscaler). Also via GREPXCEL_CA_BUNDLE '
             '/ REQUESTS_CA_BUNDLE / SSL_CERT_FILE. TLS verification stays on. '
             '(Experimental — not tested against a real intercept proxy.)',
    )
    p.add_argument(
        '--strict-env',
        action='store_true',
        help='Refuse a .env from outside the current project (also '
             'GREPXCEL_STRICT_ENV); the per-user config-dir .env stays allowed.',
    )
    _add_security_args(p)
    _add_sheet_arg(p)


# ── extract helpers ───────────────────────────────────────────────────────────

def _resolve_level(args) -> VerbosityLevel:
    if args.debug or args.verbose >= 2:
        return VerbosityLevel.DEBUG
    if args.verbose == 1:
        return VerbosityLevel.VERBOSE
    return VerbosityLevel.NORMAL


def _process_file(pattern: str, data_file: str, args, stem: str = None) -> bool:
    """
    Run the engine on one data file.
    Returns True if the file had no errors or warnings, False otherwise.
    """
    level = _resolve_level(args)

    if len(args.files) > 1:
        print(f'\n{"─" * 62}', file=sys.stderr)
        print(f'  File: {data_file}', file=sys.stderr)
        print(f'{"─" * 62}', file=sys.stderr)

    logger = Logger(
        level=level,
        log_file=args.log,
        log_format=getattr(args, 'log_format', 'text'),
        source=data_file,
    )
    output_format = getattr(args, 'format', 'nested')
    all_sheets = getattr(args, 'all_sheets', False)

    try:
        engine = Engine()
        if all_sheets:
            result = engine.process_all(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                max_rows=args.max_rows,
                max_cols=args.max_columns,
                output_format=output_format,
            )
        else:
            result = engine.process(
                pattern, data_file, logger=logger,
                max_file_mb=args.max_size,
                max_uncompressed_mb=args.max_uncompressed,
                max_cell_len=args.max_cell_len,
                max_rows=args.max_rows,
                max_cols=args.max_columns,
                sheet=_resolve_sheet(args),
                output_format=output_format,
            )
    except Exception as exc:
        print(colorize_marks(
            f'\n  ✗  Unexpected error processing {data_file}: {exc}',
            should_color(sys.stderr)), file=sys.stderr)
        return False
    finally:
        logger.close()

    if getattr(args, 'meta', False):
        result['_meta'] = logger.build_meta()

    if args.output:
        os.makedirs(args.output, exist_ok=True)
        out_path = os.path.join(args.output, f'{stem}.json')
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, default=_json_default)
        print(f'\n  JSON written to: {out_path}', file=sys.stderr)
    else:
        json.dump(result, sys.stdout, indent=2, default=_json_default)
        sys.stdout.write('\n')

    return not (logger.has_errors() or logger.has_warnings())


def _output_stem(data_file: str, all_files: list[str]) -> str:
    """
    Build a collision-safe output filename stem.
    If all input files have unique basenames, use the basename alone.
    If there are duplicates, prefix with the parent directory name.
    """
    basename = os.path.splitext(os.path.basename(data_file))[0]
    all_basenames = [os.path.splitext(os.path.basename(f))[0] for f in all_files]
    if all_basenames.count(basename) > 1:
        parent = os.path.basename(os.path.dirname(os.path.abspath(data_file)))
        return f'{parent}_{basename}'
    return basename


# ── docs handler ─────────────────────────────────────────────────────────────

def _run_docs(args) -> int:
    from .docs_generator import DocsGenerator
    DocsGenerator().write(args.output)
    print(f'Pattern reference written to: {args.output}', file=sys.stderr)
    return 0


# ── lint handler ─────────────────────────────────────────────────────────────

def _run_lint(args) -> int:
    from .lint import run_lint
    return run_lint(args.files)


# ── validate-pattern handler ─────────────────────────────────────────────────

def _run_validate(args) -> int:
    from .pattern_check import run_validate
    return run_validate(args.files, verbose=getattr(args, 'verbose', False))


# ── schema handler ──────────────────────────────────────────────────────────

def _run_skill(args) -> int:
    from .skill import run_skill
    return run_skill(args.target, getattr(args, 'output', None))


def _run_schema(args) -> int:
    from .schema import run_schema
    if not args.output:
        return run_schema(args.files)
    out = open(args.output, 'w', encoding='utf-8')
    try:
        rc = run_schema(args.files, out=out)
    finally:
        out.close()
    if rc == 0:
        print(f'Schema written to: {args.output}', file=sys.stderr)
    return rc


# ── doctor handler ───────────────────────────────────────────────────────────

def _run_doctor(args) -> int:
    from .doctor import run_doctor
    _load_dotenv(strict=getattr(args, 'strict_env', False))  # reflect .env/keys
    return run_doctor(area=getattr(args, 'area', 'all'),
                      probe=not getattr(args, 'no_probe', False),
                      strict_env=_strict_env_enabled(getattr(args, 'strict_env', False)))


# ── draft handler ─────────────────────────────────────────────────────────────

def _is_project_root(directory: str) -> bool:
    """A directory that marks a project boundary (.git or pyproject.toml)."""
    return (os.path.isdir(os.path.join(directory, '.git'))
            or os.path.isfile(os.path.join(directory, 'pyproject.toml')))


def _discover_project_env(start: str):
    """Walk cwd upward for a .env, stopping at the project root.

    Returns (path_or_None, in_project). in_project is False only when the .env
    was found in an ancestor that is NOT part of a recognised project (a loose /
    global .env) — the case --strict-env refuses.
    """
    directory = start
    while True:
        candidate = os.path.join(directory, '.env')
        has_marker = _is_project_root(directory)
        if os.path.isfile(candidate):
            return candidate, (directory == start or has_marker)
        if has_marker:
            return None, True   # project root, no .env — never escape above it
        parent = os.path.dirname(directory)
        if parent == directory:
            return None, False  # filesystem root, no project
        directory = parent


def _config_dir() -> str:
    """Per-user config dir for grepxcel (XDG-correct via platformdirs)."""
    try:
        from platformdirs import user_config_dir
        return user_config_dir('grepxcel')
    except Exception:
        return os.path.expanduser('~/.config/grepxcel')


def _config_dir_env() -> "str | None":
    """The sanctioned global .env at <config-dir>/.env, if it exists."""
    p = os.path.join(_config_dir(), '.env')
    return p if os.path.isfile(p) else None


def _read_env_file(path: str) -> None:
    """Load KEY=VALUE lines into the environment. Real env vars always win
    (setdefault); the file is only read (no execution); quotes are stripped."""
    try:
        with open(path, encoding='utf-8') as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, val = line.split('=', 1)
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key:
                    os.environ.setdefault(key, val)
    except OSError:
        pass


def _strict_env_enabled(explicit: bool = False) -> bool:
    """--strict-env flag OR GREPXCEL_STRICT_ENV truthy."""
    if explicit:
        return True
    return os.environ.get('GREPXCEL_STRICT_ENV', '').strip().lower() in (
        '1', 'true', 'yes', 'on', 'y')


def _load_dotenv(strict: bool = False) -> None:
    """Load .env credentials for cloud backends. Precedence (highest wins):
    real environment variables > project .env (cwd→project root) >
    ~/.config/grepxcel/.env (sanctioned global).

    A .env found in an unrelated ancestor is announced on stderr; with
    --strict-env / GREPXCEL_STRICT_ENV it is refused instead (the config-dir
    .env is the sanctioned exception and is always allowed).
    """
    strict = _strict_env_enabled(strict)
    start = os.getcwd()
    project_env, in_project = _discover_project_env(start)

    if project_env and not in_project:
        if strict:
            print(f'grepxcel: refusing out-of-project .env {project_env} '
                  f'(--strict-env / GREPXCEL_STRICT_ENV). Put credentials in a '
                  f'project .env or {os.path.join(_config_dir(), ".env")}.',
                  file=sys.stderr)
            sys.exit(2)
        print(f'grepxcel: loaded environment from {project_env} '
              f'(ancestor — not part of your project)', file=sys.stderr)
    elif project_env and (os.path.realpath(os.path.dirname(project_env))
                          != os.path.realpath(start)):
        print(f'grepxcel: loaded environment from {project_env}', file=sys.stderr)

    # Project .env first (its keys win via setdefault), then the config-dir .env.
    if project_env:
        _read_env_file(project_env)
    config_env = _config_dir_env()
    if config_env:
        _read_env_file(config_env)


def _run_draft(args) -> int:
    _load_dotenv(strict=getattr(args, 'strict_env', False))  # find cloud keys
    # Wire corporate-proxy / custom-CA TLS trust before any network call
    # (model download or cloud backend). Verification stays on.
    from .proxy_support import enable_corporate_tls
    enable_corporate_tls(getattr(args, 'ca_bundle', None))
    from .drafter import ClaudeBackend, GitHubModelsBackend, OpenAICompatBackend, PatternDrafter
    sheet    = _resolve_sheet(args)
    selected = getattr(args, 'backend', 'local')
    backend  = None
    if selected == 'gemini':
        # Implemented but disabled — planned for a future release.
        print(
            '[!] The Gemini backend is planned for a future release and is not yet '
            'available.\n    Use --backend local or --backend claude.',
            file=sys.stderr,
        )
        return 1
    if selected == 'claude':
        print("[!] Excel structure description will be sent to Anthropic's API.",
              file=sys.stderr)
        backend = ClaudeBackend()
    elif selected == 'github':
        gh_model = getattr(args, 'github_model', 'openai/gpt-4o-mini')
        print(f"[!] Excel structure description will be sent to GitHub Models ({gh_model}).",
              file=sys.stderr)
        backend = GitHubModelsBackend(model=gh_model)
    elif selected == 'server':
        srv_url = getattr(args, 'server_url', 'http://localhost:1234/v1')
        srv_model = getattr(args, 'server_model', None)
        srv_key = getattr(args, 'server_api_key', 'not-needed')
        backend = OpenAICompatBackend(
            base_url=srv_url, model=srv_model, api_key=srv_key,
        )
    drafter = PatternDrafter(
        input_path=args.file,
        output_path=args.output,
        sheet=sheet,
        max_file_mb=args.max_size,
        max_uncompressed_mb=args.max_uncompressed,
        verbose=args.verbose,
        dry_run=args.dry_run,
        backend=backend,
        allow_unverified=getattr(args, 'allow_unverified_model', False),
    )
    return drafter.run()


# ── Entry point ───────────────────────────────────────────────────────────────

def main(argv=None):
    # 'suggest' is a transparent backward-compat alias for 'draft'
    if argv is None:
        argv = sys.argv[1:]
    if argv and argv[0] == 'suggest':
        argv = ['draft'] + argv[1:]

    args = _build_parser().parse_args(argv)

    if args.command == 'draft':
        sys.exit(_run_draft(args))

    if args.command == 'docs':
        sys.exit(_run_docs(args))

    if args.command == 'doctor':
        sys.exit(_run_doctor(args))

    if args.command == 'lint':
        sys.exit(_run_lint(args))

    if args.command == 'validate-pattern':
        sys.exit(_run_validate(args))

    if args.command == 'schema':
        sys.exit(_run_schema(args))

    if args.command == 'generate-skill':
        sys.exit(_run_skill(args))

    # extract
    all_ok = True
    for data_file in args.files:
        ok = _process_file(args.pattern, data_file, args,
                           stem=_output_stem(data_file, args.files))
        all_ok = all_ok and ok

    sys.exit(0 if all_ok else 1)
