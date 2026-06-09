"""Unit tests for grepxcel.model_manager."""

import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import grepxcel.model_manager as mm
from grepxcel.model_manager import (
    ModelManager,
    MODEL_FILENAME,
    MODEL_REVISION,
    _CHECK_INTERVAL,
    _autoupdate_enabled,
)


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
        from grepxcel.model_manager import default_cache_dir
        assert default_cache_dir() == Path(custom)


# ── revision pinning ──────────────────────────────────────────────────────────

class TestRevisionPinning:
    def test_model_revision_is_immutable_sha(self):
        # A pin must be a 40-char hex commit SHA, never a branch name.
        assert len(MODEL_REVISION) == 40
        assert all(c in "0123456789abcdef" for c in MODEL_REVISION.lower())

    @staticmethod
    def _fake_hf_module(captured: dict):
        """A stand-in huggingface_hub module so these tests run without the
        optional 'suggest' extra installed (it is absent from core CI)."""
        import types

        def fake_download(**kwargs):
            captured.update(kwargs)
            p = Path(kwargs["local_dir"]) / MODEL_FILENAME
            p.write_bytes(b"")
            return str(p)

        mod = types.ModuleType("huggingface_hub")
        mod.hf_hub_download = fake_download
        return mod

    def test_download_pins_revision_and_records_it(self, tmp_path, monkeypatch):
        m = _make_manager(tmp_path)
        captured: dict = {}
        monkeypatch.setitem(sys.modules, "huggingface_hub", self._fake_hf_module(captured))

        m._download(announce=False)

        assert captured["revision"] == MODEL_REVISION
        assert m._load_state()["commit_hash"] == MODEL_REVISION

    def test_download_uses_explicit_commit_when_given(self, tmp_path, monkeypatch):
        m = _make_manager(tmp_path)
        captured: dict = {}
        monkeypatch.setitem(sys.modules, "huggingface_hub", self._fake_hf_module(captured))

        m._download(announce=False, commit_hash="0" * 40)

        assert captured["revision"] == "0" * 40
        assert m._load_state()["commit_hash"] == "0" * 40


# ── auto-update is opt-in ──────────────────────────────────────────────────────

class TestAutoUpdateOptIn:
    def test_disabled_by_default(self, monkeypatch):
        monkeypatch.delenv("GREPXCEL_MODEL_AUTOUPDATE", raising=False)
        assert _autoupdate_enabled() is False

    @pytest.mark.parametrize("val", ["1", "true", "TRUE", "yes", "on"])
    def test_enabled_for_truthy_values(self, monkeypatch, val):
        monkeypatch.setenv("GREPXCEL_MODEL_AUTOUPDATE", val)
        assert _autoupdate_enabled() is True

    @pytest.mark.parametrize("val", ["0", "false", "no", "off", ""])
    def test_disabled_for_falsy_values(self, monkeypatch, val):
        monkeypatch.setenv("GREPXCEL_MODEL_AUTOUPDATE", val)
        assert _autoupdate_enabled() is False

    def test_ensure_ready_skips_update_when_disabled(self, tmp_path, monkeypatch):
        monkeypatch.delenv("GREPXCEL_MODEL_AUTOUPDATE", raising=False)
        m = _make_manager(tmp_path)
        _fake_model(tmp_path)
        with patch.object(m, "_maybe_update") as mock_update:
            m.ensure_ready()
        mock_update.assert_not_called()

    def test_ensure_ready_runs_update_when_enabled(self, tmp_path, monkeypatch):
        monkeypatch.setenv("GREPXCEL_MODEL_AUTOUPDATE", "1")
        m = _make_manager(tmp_path)
        _fake_model(tmp_path)
        with patch.object(m, "_maybe_update") as mock_update:
            m.ensure_ready()
        mock_update.assert_called_once()


# ── cached-model integrity ──────────────────────────────────────────────────────

class TestCachedModelIntegrity:
    def _model_with_fingerprint(self, tmp_path, content=b"weights-v1"):
        """A cached model whose sha256+size+mtime are recorded in state.json."""
        m = _make_manager(tmp_path)
        p = tmp_path / MODEL_FILENAME
        p.write_bytes(content)
        m._save_state(MODEL_REVISION, integrity=m._fingerprint())
        return m, p

    def test_fast_path_trusts_unchanged_file_without_hashing(self, tmp_path, monkeypatch):
        m, _ = self._model_with_fingerprint(tmp_path)
        # Unchanged size+mtime must NOT trigger a re-hash of the multi-GB file.
        monkeypatch.setattr(
            mm, "_sha256_file",
            lambda *_: pytest.fail("fast path should not re-hash an unchanged file"),
        )
        m._verify_integrity(verbose=False)  # no exception, no hashing

    def test_unchanged_file_passes_ensure_ready(self, tmp_path):
        m, p = self._model_with_fingerprint(tmp_path)
        with patch.object(m, "_maybe_update"):
            assert m.ensure_ready() == p

    def test_tampered_file_aborts(self, tmp_path):
        m, p = self._model_with_fingerprint(tmp_path)
        p.write_bytes(b"malicious-content-of-a-different-length")  # size+mtime+hash all change
        with pytest.raises(SystemExit) as exc:
            m._verify_integrity(verbose=False)
        assert exc.value.code == 1

    def test_tampered_file_allowed_with_flag(self, tmp_path):
        m, p = self._model_with_fingerprint(tmp_path)
        m._allow_unverified = True
        p.write_bytes(b"different-content")
        m._verify_integrity(verbose=False)  # warns, does NOT exit

    def test_missing_fingerprint_warns_and_proceeds(self, tmp_path):
        m = _make_manager(tmp_path)
        _fake_model(tmp_path)  # file exists, but no recorded integrity
        m._verify_integrity(verbose=False)  # Case C — no exit

    def test_benign_metadata_change_rehashes_and_refreshes(self, tmp_path):
        m, p = self._model_with_fingerprint(tmp_path)
        old_mtime = m._load_state()["integrity"]["mtime_ns"]
        # Same content, new mtime (e.g. a copy/restore): re-hash matches → refresh.
        bumped = p.stat().st_mtime_ns + 5_000_000_000
        os.utime(p, ns=(bumped, bumped))
        m._verify_integrity(verbose=False)  # no exit
        assert m._load_state()["integrity"]["mtime_ns"] != old_mtime

    def test_full_mode_rehashes_even_when_unchanged(self, tmp_path, monkeypatch):
        m, _ = self._model_with_fingerprint(tmp_path)
        monkeypatch.setenv("GREPXCEL_VERIFY_MODEL", "full")
        calls = {"n": 0}
        real = mm._sha256_file
        monkeypatch.setattr(mm, "_sha256_file",
                            lambda p: (calls.__setitem__("n", calls["n"] + 1), real(p))[1])
        m._verify_integrity(verbose=False)
        assert calls["n"] == 1  # full mode hashes despite unchanged size+mtime

    def test_full_mode_detects_same_size_tamper(self, tmp_path, monkeypatch):
        # The gap the fast path can't see: same byte length + reset mtime.
        m, p = self._model_with_fingerprint(tmp_path, content=b"weights-v1")
        recorded = m._load_state()["integrity"]
        p.write_bytes(b"weights-vX")  # identical length, different content
        os.utime(p, ns=(recorded["mtime_ns"], recorded["mtime_ns"]))  # forge mtime
        # Fast path would TRUST this; full mode catches it.
        monkeypatch.setenv("GREPXCEL_VERIFY_MODEL", "full")
        with pytest.raises(SystemExit) as exc:
            m._verify_integrity(verbose=False)
        assert exc.value.code == 1

    def test_download_records_integrity_fingerprint(self, tmp_path, monkeypatch):
        m = _make_manager(tmp_path)
        captured: dict = {}
        monkeypatch.setitem(
            sys.modules, "huggingface_hub",
            TestRevisionPinning._fake_hf_module(captured),
        )
        m._download(announce=False)
        integ = m._load_state()["integrity"]
        assert set(integ) == {"sha256", "size", "mtime_ns"}
        assert len(integ["sha256"]) == 64


class TestKnownModelHashRegistry:
    """The code-pinned KNOWN_MODEL_HASHES registry is a portable trust anchor:
    a cached file matching a known-good hash is trusted on any machine, even
    with no recorded fingerprint or a stale one, and the local record is healed."""

    import hashlib as _hashlib

    @staticmethod
    def _sha(content: bytes) -> str:
        import hashlib
        return hashlib.sha256(content).hexdigest()

    def _pin_known(self, monkeypatch, content: bytes) -> None:
        """Point the registry at the sha256 of `content` for the model filename."""
        monkeypatch.setattr(
            mm, "KNOWN_MODEL_HASHES",
            {MODEL_FILENAME: self._sha(content)},
        )

    def test_known_hash_trusts_file_with_no_fingerprint(self, tmp_path, monkeypatch):
        content = b"authentic-weights"
        self._pin_known(monkeypatch, content)
        m = _make_manager(tmp_path)
        (tmp_path / MODEL_FILENAME).write_bytes(content)  # file present, NO state.json
        m._verify_integrity(verbose=False)  # must not exit
        # trust is healed into a recorded fingerprint for the fast path next time
        assert m._load_state()["integrity"]["sha256"] == self._sha(content)

    def test_known_hash_heals_stale_state(self, tmp_path, monkeypatch):
        content = b"authentic-weights"
        self._pin_known(monkeypatch, content)
        m = _make_manager(tmp_path)
        p = tmp_path / MODEL_FILENAME
        p.write_bytes(content)
        # Record a WRONG fingerprint (as a synced/stale state.json would have),
        # with a different size so the fast path is skipped and a re-hash occurs.
        m._save_state(MODEL_REVISION, integrity={
            "sha256": "0" * 64, "size": 999999, "mtime_ns": 1,
        })
        m._verify_integrity(verbose=False)  # must not exit — file matches known good
        assert m._load_state()["integrity"]["sha256"] == self._sha(content)  # healed

    def test_unknown_file_with_no_fingerprint_still_proceeds(self, tmp_path, monkeypatch):
        # File does not match the known-good hash and has no fingerprint:
        # warn-and-proceed (does not exit), preserving the manual-file path.
        self._pin_known(monkeypatch, b"the-real-weights")
        m = _make_manager(tmp_path)
        (tmp_path / MODEL_FILENAME).write_bytes(b"some-other-file")
        m._verify_integrity(verbose=False)  # no exit

    def test_known_mismatch_with_stale_state_still_aborts(self, tmp_path, monkeypatch):
        # A recorded fingerprint that mismatches AND a file that is not a known
        # good build is genuine tamper evidence → fatal.
        self._pin_known(monkeypatch, b"the-real-weights")
        m = _make_manager(tmp_path)
        p = tmp_path / MODEL_FILENAME
        p.write_bytes(b"tampered")
        m._save_state(MODEL_REVISION, integrity={
            "sha256": "a" * 64, "size": 4, "mtime_ns": 1,
        })
        with pytest.raises(SystemExit) as exc:
            m._verify_integrity(verbose=False)
        assert exc.value.code == 1


class TestAllowUnverifiedFlag:
    def test_env_var_sets_allow(self, tmp_path, monkeypatch):
        monkeypatch.setenv("GREPXCEL_ALLOW_UNVERIFIED_MODEL", "1")
        assert ModelManager(cache_dir=tmp_path)._allow_unverified is True

    def test_constructor_flag_sets_allow(self, tmp_path, monkeypatch):
        monkeypatch.delenv("GREPXCEL_ALLOW_UNVERIFIED_MODEL", raising=False)
        assert ModelManager(cache_dir=tmp_path, allow_unverified=True)._allow_unverified is True

    def test_default_is_false(self, tmp_path, monkeypatch):
        monkeypatch.delenv("GREPXCEL_ALLOW_UNVERIFIED_MODEL", raising=False)
        assert ModelManager(cache_dir=tmp_path)._allow_unverified is False
