"""MCP server exposing grepxcel tools to AI agents."""

import io
import json
import os
import sys
import tempfile

try:
    from mcp.server.fastmcp import FastMCP
except ImportError:
    FastMCP = None


def _require_mcp():
    if FastMCP is None:
        print(
            "MCP dependencies not installed.\n"
            "Install with:  pip install 'grepxcel[mcp]'",
            file=sys.stderr,
        )
        raise SystemExit(1)


def _json_default(obj):
    import datetime
    if isinstance(obj, (datetime.date, datetime.datetime, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, datetime.timedelta):
        return str(obj)
    raise TypeError(f'Type {type(obj).__name__} is not JSON serialisable')


def create_server() -> "FastMCP":
    _require_mcp()

    mcp = FastMCP(
        "grepxcel",
        instructions=(
            "grepxcel extracts structured JSON from Excel files using a "
            "pattern file. Use 'extract' for data extraction, 'validate_pattern' "
            "to check patterns, 'lint' to inspect data files, 'schema' to get "
            "the JSON Schema of extraction output, 'docs' to generate the "
            "pattern format reference, 'doctor' to diagnose environment issues, "
            "and 'generate_examples' to create ready-to-run example files."
        ),
    )

    @mcp.tool()
    def extract(
        pattern: str,
        data: str,
        sheet: str | None = None,
        all_sheets: bool = False,
    ) -> str:
        """Extract structured JSON from an Excel file using a pattern file.

        Args:
            pattern: Path to the pattern file (.xlsx or .csv).
            data: Path to the data Excel file (.xlsx).
            sheet: Sheet name or 0-based index (default: active sheet).
            all_sheets: Process every sheet (mutually exclusive with sheet).

        Returns:
            Extracted data as a JSON string.
        """
        import grepxcel as gx
        result = gx.extract(pattern, data, sheet=sheet, all_sheets=all_sheets)
        return json.dumps(result, indent=2, default=_json_default)

    @mcp.tool()
    def validate_pattern(pattern: str, verbose: bool = False) -> str:
        """Validate a pattern file without extracting data.

        Args:
            pattern: Path to the pattern file (.xlsx or .csv).
            verbose: Show parsed config, fields, and extraction sequence.

        Returns:
            Validation result text.
        """
        from .pattern_check import run_validate
        buf = io.StringIO()
        rc = run_validate([pattern], verbose=verbose, out=buf)
        output = buf.getvalue()
        if rc != 0 and not output.strip():
            output = f"Validation failed (exit code {rc})"
        return output

    @mcp.tool()
    def lint(file: str) -> str:
        """Inspect an Excel file for potential extraction issues.

        Checks format, encryption, sheet dimensions, merged cells,
        formula cells, and known corporate-environment issues.

        Args:
            file: Path to the Excel file (.xlsx).

        Returns:
            Lint report text.
        """
        from .lint import run_lint
        buf = io.StringIO()
        run_lint([file], out=buf)
        return buf.getvalue()

    @mcp.tool()
    def schema(pattern: str) -> str:
        """Generate a JSON Schema (draft 2020-12) from a pattern file.

        The schema describes the structure of the extraction output for
        the given pattern. Use it to validate extracted JSON.

        Args:
            pattern: Path to the pattern file (.xlsx or .csv).

        Returns:
            JSON Schema as a JSON string.
        """
        from .schema import run_schema
        buf = io.StringIO()
        run_schema([pattern], out=buf)
        return buf.getvalue()

    @mcp.tool()
    def docs(output_dir: str | None = None) -> str:
        """Generate a colour-coded pattern-reference.xlsx explaining the pattern format.

        Args:
            output_dir: Directory to write the file to. Uses a temp directory
                       if not specified.

        Returns:
            Path to the generated reference file.
        """
        from .docs_generator import DocsGenerator
        if output_dir is None:
            output_dir = tempfile.mkdtemp(prefix='grepxcel-docs-')
        path = os.path.join(output_dir, 'pattern-reference.xlsx')
        DocsGenerator().write(path)
        return f"Pattern reference written to: {path}"

    @mcp.tool()
    def doctor(area: str = "all") -> str:
        """Check that the environment is ready for grepxcel.

        Runs preflight checks for dependencies, API keys, model cache,
        and proxy/TLS configuration. Use this to diagnose issues.

        Args:
            area: What to check — 'extract', 'draft', or 'all' (default).

        Returns:
            Diagnostic report text.
        """
        from .doctor import run_doctor
        buf = io.StringIO()
        run_doctor(area=area, probe=False, out=buf)
        return buf.getvalue()

    @mcp.tool()
    def generate_examples(output_dir: str = "grepxcel-examples") -> str:
        """Create ready-to-run example files for learning grepxcel.

        Copies 4 bundled examples (pattern + data xlsx + README) into
        a local directory. Each example demonstrates a different level
        of pattern complexity.

        Args:
            output_dir: Directory to create (default: ./grepxcel-examples/).

        Returns:
            Summary of created examples.
        """
        from .examples_generator import generate_examples as _gen
        _gen(output_dir)
        return f"Examples created in: {output_dir}"

    return mcp


def run_server():
    """Start the MCP server on stdio transport."""
    server = create_server()
    server.run(transport="stdio")
