# PyInstaller spec for grepxcel — base CLI binary (no optional extras).
#
# Bundles: core extraction engine, pattern parser, all base commands.
# Excluded: fastapi/uvicorn/jinja2 (web), textual (wizard TUI),
#           llama-cpp/huggingface_hub (draft), pandas/polars, mcp.
#
# Usage (from repo root):
#   pip install ".[build]"
#   pyinstaller build/grepxcel.spec
#
# Output: dist/grepxcel  (or dist/grepxcel.exe on Windows)

block_cipher = None

datas = [
    # Jinja2 HTML template — used by the web-wizard command. Bundled so
    # `grepxcel web-wizard` produces a helpful "install grepxcel[web]" error
    # rather than a FileNotFoundError when run from the binary.
    ('grepxcel/templates/wizard.html', 'grepxcel/templates'),
    # Bundled examples — needed by `grepxcel generate-examples`.
    ('grepxcel/examples', 'grepxcel/examples'),
]

hiddenimports = [
    # openpyxl uses lazy imports for its cell writer — PyInstaller misses it.
    'openpyxl',
    'openpyxl.cell._writer',
    'openpyxl.styles.stylesheet',
    # defusedxml, structlog, regex are direct dependencies; list them
    # explicitly so PyInstaller's static analysis does not prune them.
    'defusedxml',
    'structlog',
    'regex',
    # importlib.metadata is used by sbom.py and __version__ detection.
    'importlib.metadata',
    'importlib.metadata._meta',
    'pkg_resources',
    'pkg_resources.extern',
]

a = Analysis(
    ['grepxcel/__main__.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Exclude all optional extras — keeps the binary lean.
    excludes=[
        'fastapi', 'uvicorn', 'jinja2', 'starlette', 'anyio', 'httpx',
        'textual', 'rich',
        'llama_cpp', 'huggingface_hub', 'platformdirs',
        'pandas', 'polars',
        'mcp',
        'anthropic', 'openai', 'google.genai',
        'pytest', 'hypothesis', 'jsonschema',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='grepxcel',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX compression — reduces binary size ~40-60%. CI installs upx-ucl.
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,   # macOS: do not rewrite sys.argv from Apple events
    target_arch=None,       # use the runner's native arch
    codesign_identity=None,
    entitlements_file=None,
)
