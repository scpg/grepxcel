"""Macro content is refused by what a file CONTAINS, not by what it is called.

`.xlsm` and `.xlsb` were refused by extension only. Both they and `.xlsx` are ZIP
archives beginning with `PK`, so renaming a macro-enabled workbook to `.xlsx`
passed the extension allow-list and the magic-byte check and was accepted — while
the byte-identical original was refused. The protection read stronger than it was.

Worth stating what this does and does not do, because the tests assert exactly
that boundary: grepxcel never executes macro content (openpyxl has no VBA engine,
and formulas are read from cached values rather than evaluated), so this is not
what prevents code running here. It prevents grepxcel silently accepting and
passing on a macro-bearing file that something downstream may open in Excel, and
it makes the refusal depend on the bytes rather than the filename. It is not
malware detection and must not be described as such.
"""
import shutil
import zipfile

import pytest

import openpyxl

from grepxcel.security import (
    SecurityError,
    find_executable_parts,
    validate_file,
)


@pytest.fixture
def clean_workbook(tmp_path):
    """An ordinary macro-free .xlsx."""
    path = tmp_path / 'clean.xlsx'
    wb = openpyxl.Workbook()
    wb.active['A1'] = 'Invoice No'
    wb.save(path)
    return path


def _with_part(source, dest, part_name, payload=b'\xd0\xcf\x11\xe0FAKE'):
    """Copy *source* and add *part_name* to the archive."""
    shutil.copy(source, dest)
    with zipfile.ZipFile(dest, 'a') as zf:
        zf.writestr(part_name, payload)
    return dest


# ── the bypass this exists to close ───────────────────────────────────────────

def test_a_clean_workbook_is_still_accepted(clean_workbook):
    """The guard must not reject ordinary files."""
    validate_file(str(clean_workbook), 5, 50)


def test_vba_project_renamed_to_xlsx_is_refused(clean_workbook, tmp_path):
    """The original bypass: macro bytes under an .xlsx name."""
    disguised = _with_part(clean_workbook, tmp_path / 'looks_clean.xlsx',
                           'xl/vbaProject.bin')
    with pytest.raises(SecurityError) as excinfo:
        validate_file(str(disguised), 5, 50)
    assert 'xl/vbaProject.bin' in str(excinfo.value), (
        'the message must name the part found, so the user can verify the claim'
    )


def test_excel_4_macro_sheet_is_refused(clean_workbook, tmp_path):
    """XLM macro sheets are a separate mechanism from VBA and a known vector."""
    xlm = _with_part(clean_workbook, tmp_path / 'xlm.xlsx',
                     'xl/macrosheets/sheet1.xml', b'<xml/>')
    with pytest.raises(SecurityError) as excinfo:
        validate_file(str(xlm), 5, 50)
    assert 'macrosheets' in str(excinfo.value)


def test_detection_is_case_insensitive(clean_workbook, tmp_path):
    """OOXML part names are fixed, so odd case means a hand-built archive —
    exactly the file least deserving of the benefit of the doubt."""
    odd = _with_part(clean_workbook, tmp_path / 'odd_case.xlsx',
                     'xl/VBAProject.BIN')
    with pytest.raises(SecurityError):
        validate_file(str(odd), 5, 50)


def test_the_extension_check_still_fires_first_for_xlsm(clean_workbook, tmp_path):
    """A correctly-named .xlsm keeps its own clearer message rather than the
    generic content one."""
    xlsm = tmp_path / 'macro.xlsm'
    shutil.copy(clean_workbook, xlsm)
    with pytest.raises(SecurityError) as excinfo:
        validate_file(str(xlsm), 5, 50)
    assert '.xlsm' in str(excinfo.value)


# ── the message must not overclaim ────────────────────────────────────────────

def test_the_refusal_does_not_claim_the_file_is_malicious(clean_workbook, tmp_path):
    """A refusal to process is not a verdict on the file. Claiming otherwise
    would be both wrong and alarming — grepxcel does no malware analysis."""
    disguised = _with_part(clean_workbook, tmp_path / 'd.xlsx', 'xl/vbaProject.bin')
    with pytest.raises(SecurityError) as excinfo:
        validate_file(str(disguised), 5, 50)
    message = str(excinfo.value)
    assert 'not a finding that the file is malicious' in message
    assert 'does not scan for malware' in message


def test_the_refusal_says_how_to_proceed(clean_workbook, tmp_path):
    disguised = _with_part(clean_workbook, tmp_path / 'd.xlsx', 'xl/vbaProject.bin')
    with pytest.raises(SecurityError) as excinfo:
        validate_file(str(disguised), 5, 50)
    assert 'Save As' in str(excinfo.value)


# ── find_executable_parts ─────────────────────────────────────────────────────

def test_find_returns_empty_for_a_clean_file(clean_workbook):
    assert find_executable_parts(str(clean_workbook)) == []


def test_find_reports_the_part_and_a_description(clean_workbook, tmp_path):
    disguised = _with_part(clean_workbook, tmp_path / 'd.xlsx', 'xl/vbaProject.bin')
    found = find_executable_parts(str(disguised))
    assert len(found) == 1
    name, description = found[0]
    assert name == 'xl/vbaProject.bin'
    assert 'VBA' in description


def test_find_reports_every_distinct_mechanism(clean_workbook, tmp_path):
    both = _with_part(clean_workbook, tmp_path / 'both.xlsx', 'xl/vbaProject.bin')
    with zipfile.ZipFile(both, 'a') as zf:
        zf.writestr('xl/macrosheets/sheet1.xml', b'<xml/>')
    assert len(find_executable_parts(str(both))) == 2, (
        'a file using both mechanisms must report both, not stop at the first'
    )


def test_find_is_quiet_on_an_unreadable_archive(tmp_path):
    """It returns [] rather than raising, because `lint` calls it to *report*.
    Refusing a corrupt archive belongs to the ZIP check, which runs first."""
    junk = tmp_path / 'junk.xlsx'
    junk.write_bytes(b'PK\x03\x04not actually a zip')
    assert find_executable_parts(str(junk)) == []


# ── lint reports instead of refusing ──────────────────────────────────────────

def test_lint_reports_macro_content_as_a_failure(clean_workbook, tmp_path):
    """lint must still open and describe the file — a user whose extraction was
    refused comes here to find out why."""
    from grepxcel.lint import lint_file

    disguised = _with_part(clean_workbook, tmp_path / 'd.xlsx', 'xl/vbaProject.bin')
    labels = [label for _sev, label, _msg in lint_file(str(disguised))]
    assert 'macro content' in labels


def test_lint_says_nothing_about_macros_for_a_clean_file(clean_workbook):
    from grepxcel.lint import lint_file

    labels = [label for _sev, label, _msg in lint_file(str(clean_workbook))]
    assert 'macro content' not in labels


def test_lint_and_security_share_one_part_list():
    """Two copies of the list would drift, and the drift would be silent."""
    import inspect

    from grepxcel import lint as lint_mod

    source = inspect.getsource(lint_mod._check_macro_content)
    assert '_EXECUTABLE_PARTS' in source, (
        'lint must import the part list from security rather than restate it'
    )
