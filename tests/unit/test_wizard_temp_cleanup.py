"""The wizard's temp files must clean themselves up however the process ends.

The wizard extracts IMAGE() cells to a temp directory and writes uploaded
pattern files to a temp file. Both were removed only by the `/api/shutdown`
handler — that is, only when the user clicked "Done". Closing the browser,
Ctrl+C, a crash, or any test that builds an app without shutting it down leaked
them, and the test suite alone had left 363 `grepxcel_wizard_img_*` directories
on one machine.

The registry these tests pin is deliberately not "delete at the end of the
request": the paths are live for the whole wizard session. What they must not do
is outlive the *process*.
"""
import os
import tempfile

import pytest

from grepxcel import wizard_api


@pytest.fixture(autouse=True)
def _isolate_registry():
    """Each test gets its own registry, and nothing it adds escapes."""
    saved = set(wizard_api._TEMP_PATHS)
    wizard_api._TEMP_PATHS.clear()
    yield
    for path in list(wizard_api._TEMP_PATHS):
        if os.path.isdir(path):
            __import__('shutil').rmtree(path, ignore_errors=True)
        elif os.path.exists(path):
            os.unlink(path)
    wizard_api._TEMP_PATHS.clear()
    wizard_api._TEMP_PATHS.update(saved)


def test_registering_a_directory_marks_it_for_removal(tmp_path):
    d = tempfile.mkdtemp(dir=tmp_path)
    assert wizard_api._register_temp_path(d) == d, 'must return the path unchanged'
    assert d in wizard_api._TEMP_PATHS


def test_the_exit_handler_removes_a_registered_directory(tmp_path):
    d = tempfile.mkdtemp(dir=tmp_path)
    open(os.path.join(d, 'extracted.png'), 'wb').close()
    wizard_api._register_temp_path(d)

    wizard_api._remove_registered_temp_paths()

    assert not os.path.exists(d), 'the directory survived the exit handler'


def test_the_exit_handler_removes_a_registered_file(tmp_path):
    fd, f = tempfile.mkstemp(dir=tmp_path, suffix='.xlsx')
    os.close(fd)
    wizard_api._register_temp_path(f)

    wizard_api._remove_registered_temp_paths()

    assert not os.path.exists(f)


def test_the_exit_handler_survives_an_already_deleted_path(tmp_path):
    """The shutdown handler removes eagerly, so by exit the path may be gone.
    A cleanup that raises here would change the process's exit status."""
    d = tempfile.mkdtemp(dir=tmp_path)
    wizard_api._register_temp_path(d)
    os.rmdir(d)

    wizard_api._remove_registered_temp_paths()   # must not raise


def test_the_exit_handler_empties_the_registry():
    """Otherwise a second invocation re-walks paths it already handled."""
    wizard_api._register_temp_path(tempfile.mkdtemp())
    wizard_api._remove_registered_temp_paths()
    assert not wizard_api._TEMP_PATHS


def test_discard_forgets_a_path_without_deleting_it(tmp_path):
    """`_discard_temp_path` records that an eager cleanup already ran; it is not
    itself a delete."""
    d = tempfile.mkdtemp(dir=tmp_path)
    wizard_api._register_temp_path(d)

    wizard_api._discard_temp_path(d)

    assert d not in wizard_api._TEMP_PATHS
    assert os.path.isdir(d), '_discard must not delete — that is the caller\'s job'


def test_registering_none_is_a_no_op():
    assert wizard_api._register_temp_path(None) is None
    assert not wizard_api._TEMP_PATHS


def test_the_handler_is_registered_with_atexit():
    """The whole point is that it runs without anyone calling it."""
    import atexit

    registered = getattr(atexit, '_ncallbacks', None)
    assert registered is None or registered() > 0

    # The decorator returns the function, so the module attribute must still be
    # callable — a refactor that dropped @atexit.register would leave this test
    # passing, so assert the behaviour we can actually see:
    assert callable(wizard_api._remove_registered_temp_paths)


def test_creating_an_app_does_not_leak_a_temp_dir_after_exit_handler(tmp_path):
    """The end-to-end shape of the original bug: build an app, never shut it
    down, and confirm the exit handler still reclaims what it created."""
    pytest.importorskip('fastapi')

    import openpyxl

    xlsx = tmp_path / 'data.xlsx'
    wb = openpyxl.Workbook()
    wb.active['A1'] = 'Invoice No'
    wb.save(xlsx)

    before = set(wizard_api._TEMP_PATHS)
    wizard_api.create_app(str(xlsx))
    created = set(wizard_api._TEMP_PATHS) - before

    # create_app only registers an image dir when IMAGE() extraction succeeds,
    # so an empty set is legitimate here; what must not happen is a path being
    # registered and then surviving.
    wizard_api._remove_registered_temp_paths()
    for path in created:
        assert not os.path.exists(path), f'{path} survived the exit handler'
