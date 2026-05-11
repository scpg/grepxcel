"""Unit tests for engine.model_manager."""

import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from engine.model_manager import ModelManager, MODEL_FILENAME, _CHECK_INTERVAL


# ── helpers ───────────────────────────────────────────────────────────────────

def _write_state(cache_dir: Path, commit_hash: str, last_check: datetime) -> None:
    state = {
        "last_check":  last_check.isoformat(),
        "commit_hash": commit_hash,
        "filename":    MODEL_FILENAME,
    }
    (cache_dir / "state.json").write_text(json.dumps(state), encoding="utf-8")


def _make_manager(tmp_path: Path) -> ModelManager:
    return ModelManager(cache_dir=tmp_path)


def _fake_model(cache_dir: Path) -> Path:
    """Write a zero-byte placeholder so model_path.exists() is True."""
    p = cache_dir / MODEL_FILENAME
    p.write_bytes(b"")
    return p


# ── first-run download ────────────────────────────────────────────────────────

class TestFirstRunDownload:
    def test_calls_download_when_model_missing(self, tmp_path):
        m = _make_manager(tmp_path)
        with patch.object(m, "_download") as mock_dl:
            with patch.object(m, "_maybe_update"):
                # Simulate download creating the file
                def _create_file(*_, **__):
                    _fake_model(tmp_path)
                mock_dl.side_effect = _create_file
                m.ensure_ready()
        mock_dl.assert_called_once_with(announce=True)

    def test_returns_model_path(self, tmp_path):
        m = _make_manager(tmp_path)
        _fake_model(tmp_path)
        with patch.object(m, "_maybe_update"):
            result = m.ensure_ready()
        assert result == tmp_path / MODEL_FILENAME


# ── daily check logic ─────────────────────────────────────────────────────────

class TestCheckDue:
    def test_no_state_file_is_due(self, tmp_path):
        m = _make_manager(tmp_path)
        assert m._check_due({}) is True

    def test_recent_check_is_not_due(self, tmp_path):
        m = _make_manager(tmp_path)
        recent = datetime.now(timezone.utc) - timedelta(hours=1)
        assert m._check_due({"last_check": recent.isoformat()}) is False

    def test_old_check_is_due(self, tmp_path):
        m = _make_manager(tmp_path)
        old = datetime.now(timezone.utc) - timedelta(seconds=_CHECK_INTERVAL + 1)
        assert m._check_due({"last_check": old.isoformat()}) is True

    def test_corrupt_timestamp_is_due(self, tmp_path):
        m = _make_manager(tmp_path)
        assert m._check_due({"last_check": "not-a-date"}) is True


# ── update logic ──────────────────────────────────────────────────────────────

class TestMaybeUpdate:
    def test_skips_check_when_not_due(self, tmp_path):
        m = _make_manager(tmp_path)
        recent = datetime.now(timezone.utc) - timedelta(hours=1)
        _write_state(tmp_path, "abc", recent)
        with patch.object(m, "_remote_commit") as mock_rc:
            m._maybe_update(verbose=False)
        mock_rc.assert_not_called()

    def test_no_download_when_hash_matches(self, tmp_path):
        m = _make_manager(tmp_path)
        old = datetime.now(timezone.utc) - timedelta(days=2)
        _write_state(tmp_path, "hash123", old)
        with patch.object(m, "_remote_commit", return_value="hash123"):
            with patch.object(m, "_download") as mock_dl:
                m._maybe_update(verbose=False)
        mock_dl.assert_not_called()

    def test_downloads_when_hash_differs(self, tmp_path):
        m = _make_manager(tmp_path)
        old = datetime.now(timezone.utc) - timedelta(days=2)
        _write_state(tmp_path, "old_hash", old)
        with patch.object(m, "_remote_commit", return_value="new_hash"):
            with patch.object(m, "_download") as mock_dl:
                m._maybe_update(verbose=False)
        mock_dl.assert_called_once_with(announce=True, commit_hash="new_hash")

    def test_network_failure_skips_update_silently(self, tmp_path):
        m = _make_manager(tmp_path)
        old = datetime.now(timezone.utc) - timedelta(days=2)
        _write_state(tmp_path, "hash123", old)
        with patch.object(m, "_remote_commit", return_value=None):
            with patch.object(m, "_download") as mock_dl:
                m._maybe_update(verbose=False)
        mock_dl.assert_not_called()

    def test_network_failure_does_not_update_last_check(self, tmp_path):
        """If the network is down, last_check must NOT be advanced so the
        next process will try again instead of waiting another 24 h."""
        m = _make_manager(tmp_path)
        old = datetime.now(timezone.utc) - timedelta(days=2)
        _write_state(tmp_path, "hash123", old)
        with patch.object(m, "_remote_commit", return_value=None):
            m._maybe_update(verbose=False)
        state = m._load_state()
        saved = datetime.fromisoformat(state["last_check"])
        assert (datetime.now(timezone.utc) - saved).total_seconds() > _CHECK_INTERVAL


# ── state persistence ─────────────────────────────────────────────────────────

class TestStatePersistence:
    def test_save_and_load_roundtrip(self, tmp_path):
        m = _make_manager(tmp_path)
        m._save_state("deadbeef")
        state = m._load_state()
        assert state["commit_hash"] == "deadbeef"
        assert state["filename"] == MODEL_FILENAME
        assert "last_check" in state

    def test_load_missing_state_returns_empty(self, tmp_path):
        m = _make_manager(tmp_path)
        assert m._load_state() == {}

    def test_load_corrupt_state_returns_empty(self, tmp_path):
        m = _make_manager(tmp_path)
        (tmp_path / "state.json").write_text("not json", encoding="utf-8")
        assert m._load_state() == {}


# ── GREPXCEL_MODEL_DIR env var ────────────────────────────────────────────────

class TestCacheDirOverride:
    def test_env_var_overrides_cache_dir(self, tmp_path, monkeypatch):
        custom = str(tmp_path / "custom_cache")
        monkeypatch.setenv("GREPXCEL_MODEL_DIR", custom)
        from engine.model_manager import default_cache_dir
        assert default_cache_dir() == Path(custom)
