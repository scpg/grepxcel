"""
Model manager: ensures the pinned GGUF model is present locally and downloads
it atomically.

Supply-chain stance (deliberate):
  - The model is pinned to an immutable commit revision (MODEL_REVISION), not a
    moving branch. hf_hub_download verifies the file hash against the Hub for
    that exact revision, so what you get is reproducible and tamper-evident.
  - Auto-update is OFF by default. A tool that silently replaces model weights
    from a third-party repo every day is a supply-chain risk. To opt in to
    tracking upstream main, set GREPXCEL_MODEL_AUTOUPDATE=1.
  - The cached file is checksummed at download time and re-verified cheaply on
    every reuse (size+mtime fast path; re-hash only on change). A mismatch means
    the weights on disk are not what we fetched, so it aborts unless the user
    opts out with GREPXCEL_ALLOW_UNVERIFIED_MODEL / --allow-unverified-model.

All network activity is limited to:
  - one file download of the pinned revision on first run
  - (only when auto-update is opted in) one lightweight commit-check per day,
    plus a download if a newer upstream commit is found

Inference never touches the network — it runs entirely in-process via llama-cpp-python.
"""

import hashlib
import json
import os
import sys
import tempfile
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

# ── Pinned model ──────────────────────────────────────────────────────────────
# To upgrade the bundled model, change these constants only. MODEL_REVISION must
# be an immutable commit SHA from https://huggingface.co/<MODEL_REPO_ID>/commits
# — never a branch name like "main".

# Gemma-4-E4B-it — the best LOCAL model in our execution-based draft eval
# (~51% value-recall across all fixtures vs ~42% for the previous default
# Qwen2.5-Coder-7B). Cloud models and the free GitHub Models backend are still
# stronger (81-86%); for the highest quality at no dollar cost use
# `grepxcel draft --backend github`. Gemma 4 is a large generational jump over
# Gemma 2 on this task. ~5 GB Q4_K_M, comfortable on an 8 GB GPU.
MODEL_REPO_ID    = "unsloth/gemma-4-E4B-it-GGUF"
MODEL_FILENAME   = "gemma-4-E4B-it-Q4_K_M.gguf"
MODEL_REVISION   = "653803f092503c04a65164346f3208a36e707693"  # pinned commit
MODEL_CHAT_FORMAT = None  # auto-detect from GGUF metadata (Gemma embeds its template)

_SIZE_HINT        = "~5.0 GB"
_CHECK_INTERVAL   = 86_400          # seconds — 24 h
_AUTOUPDATE_ENV   = "GREPXCEL_MODEL_AUTOUPDATE"

# Cached-model integrity. hf_hub_download verifies the hash at download time, but
# the cached multi-GB file is then trusted on every later run. We record its
# sha256 + size + mtime at download time and re-check cheaply on reuse (see
# ModelManager._verify_integrity). Set this to opt out of a hard failure when the
# cached file no longer matches what we downloaded.
_ALLOW_UNVERIFIED_ENV = "GREPXCEL_ALLOW_UNVERIFIED_MODEL"
_HASH_CHUNK           = 1024 * 1024  # 1 MiB read buffer for hashing

# Verification mode. Default "fast": trust the size+mtime fast path and only
# re-hash when those change. Set GREPXCEL_VERIFY_MODEL=full to re-hash the whole
# file on EVERY run — slower (reads ~4.7 GB each time) but closes the gap where a
# local attacker overwrites the file with same-size content and resets its mtime.
_VERIFY_MODE_ENV      = "GREPXCEL_VERIFY_MODEL"

# ── Known-good model hashes (code-pinned trust anchor) ─────────────────────────
# sha256 == the HuggingFace git-LFS oid of the pinned build (the same value the
# Hub verifies at download time). Because this lives in the code it is
# version-controlled, so trust travels between machines and survives a lost or
# stale local state.json: a cached file whose sha256 matches the entry here is
# trusted on ANY machine, with no recorded fingerprint required, and a stale
# state.json is healed automatically. Models not listed here fall back to the
# per-machine state.json fingerprint (the original behaviour) — that's expected
# for ad-hoc / experimental models that aren't the shipped default.
#
# When bumping the pinned model (MODEL_REVISION / MODEL_FILENAME), update the
# matching entry here. Fetch the value from the Hub without downloading:
#   HfApi().get_paths_info(MODEL_REPO_ID, [MODEL_FILENAME], revision=MODEL_REVISION)[0].lfs.sha256
KNOWN_MODEL_HASHES: dict[str, str] = {
    # Current default.
    "gemma-4-E4B-it-Q4_K_M.gguf":
        "519b9793ed6ce0ff530f1b7c96e848e08e49e7af4d57bb97f76215963a54146d",
    # Previous default — kept so an already-cached copy stays trusted.
    "Qwen2.5-Coder-7B-Instruct-Q4_K_M.gguf":
        "1664fccab734674a50763490a8c6931b70e3f2f8ec10031b54806d30e5f956b6",
}


def _known_hashes_for(filename: str) -> set[str]:
    """Acceptable sha256(s) for a cached model file from the code-pinned registry."""
    h = KNOWN_MODEL_HASHES.get(filename)
    return {h} if h else set()

# Download progress watchdog (multi-GB download → never leave the user blind).
# huggingface_hub shows a live tqdm bar in a TTY; the watchdog adds (a) periodic
# heartbeats when stderr is NOT a TTY (logs/CI, where the bar can't redraw) and
# (b) an actionable stall warning if no bytes arrive for a while — the exact case
# where an anonymous, rate-limited HF download hangs silently.
_DL_POLL_SECONDS      = 2.0
_DL_HEARTBEAT_SECONDS = 15.0
_DL_STALL_SECONDS     = 30.0


def _largest_file_bytes(root: str) -> int:
    """Largest single file size (bytes) under root — tracks the in-flight download.

    hf_hub writes the partial file as <root>/.cache/huggingface/download/*.incomplete,
    so the biggest file under the temp dir is the download's current size.
    """
    best = 0
    for dirpath, _dirs, names in os.walk(root):
        for name in names:
            try:
                best = max(best, os.path.getsize(os.path.join(dirpath, name)))
            except OSError:
                pass
    return best


def _autoupdate_enabled() -> bool:
    """Auto-update is opt-in. Returns True only when explicitly enabled."""
    return os.environ.get(_AUTOUPDATE_ENV, "").strip().lower() in (
        "1", "true", "yes", "on",
    )


def _allow_unverified_env() -> bool:
    """Returns True when GREPXCEL_ALLOW_UNVERIFIED_MODEL opts out of the check."""
    return os.environ.get(_ALLOW_UNVERIFIED_ENV, "").strip().lower() in (
        "1", "true", "yes", "on",
    )


def _full_verify_enabled() -> bool:
    """True when GREPXCEL_VERIFY_MODEL=full forces a re-hash on every run."""
    return os.environ.get(_VERIFY_MODE_ENV, "").strip().lower() == "full"


def _sha256_file(path: Path) -> str:
    """Streaming sha256 of a file (reads in 1 MiB chunks — never loads 4.7 GB)."""
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(_HASH_CHUNK), b""):
            h.update(chunk)
    return h.hexdigest()


# ── Cache location ────────────────────────────────────────────────────────────

def default_cache_dir() -> Path:
    """
    Returns the directory where the model file and state are stored.

    Resolution order:
      1. GREPXCEL_MODEL_DIR env var — explicit override (Docker volumes, a shared
         model dir, web-service deployments).
      2. The platform-appropriate per-user cache dir via `platformdirs`
         (the de-facto standard, used by pip/poetry):
           - Linux:   $XDG_CACHE_HOME/grepxcel/models  (default ~/.cache/grepxcel/models)
           - macOS:   ~/Library/Caches/grepxcel/models
           - Windows: %LOCALAPPDATA%\\grepxcel\\Cache\\models
      3. Fallback (platformdirs not installed): ~/.cache/grepxcel/models.

    On Linux the platformdirs path matches the historical default, so existing
    caches keep working unchanged.
    """
    base = os.environ.get("GREPXCEL_MODEL_DIR")
    if base:
        return Path(base)
    try:
        from platformdirs import user_cache_dir
        return Path(user_cache_dir("grepxcel")) / "models"
    except ImportError:
        # platformdirs ships in the [suggest] extra; if a caller reaches here
        # without it, fall back to the XDG-style default that works on Linux.
        return Path.home() / ".cache" / "grepxcel" / "models"


# ── Model manager ─────────────────────────────────────────────────────────────

class ModelManager:
    """
    Manages the pinned GGUF model file lifecycle:
      - downloads on first run
      - checks once per day for a new upstream version
      - atomically replaces the old file when an update is found
      - falls back gracefully when the network is unavailable
    """

    _STATE_FILE = "state.json"

    def __init__(self, cache_dir: Path | None = None, allow_unverified: bool = False):
        self.cache_dir  = Path(cache_dir) if cache_dir else default_cache_dir()
        self.model_path = self.cache_dir / MODEL_FILENAME
        self._state_path = self.cache_dir / self._STATE_FILE
        # CLI flag OR env var opts out of the cached-model integrity check.
        self._allow_unverified = bool(allow_unverified) or _allow_unverified_env()

    def ensure_ready(self, verbose: bool = False) -> Path:
        """
        Guarantee the model file is present (downloading the pinned revision on
        first run). On reuse the cached file's integrity is verified cheaply
        before it is handed to the inference engine. Update checks happen only
        when auto-update is opted in via the GREPXCEL_MODEL_AUTOUPDATE env var.
        Returns the path to the model file — always valid on return.
        """
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        if not self.model_path.exists():
            self._download(announce=True)
        else:
            # Reusing a cached file we did not just download+verify: confirm it
            # still matches what we recorded, before trusting it for inference.
            self._verify_integrity(verbose)
            if _autoupdate_enabled():
                self._maybe_update(verbose)

        return self.model_path

    # ── state helpers ─────────────────────────────────────────────────────────

    def _load_state(self) -> dict:
        try:
            return json.loads(self._state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_state(self, commit_hash: str, integrity: dict | None = None) -> None:
        # Merge into existing state so refreshing last_check (e.g. the no-update
        # path) never drops the recorded integrity fingerprint.
        state = self._load_state()
        state.update({
            "last_check":  datetime.now(timezone.utc).isoformat(),
            "commit_hash": commit_hash,
            "filename":    MODEL_FILENAME,
        })
        if integrity is not None:
            state["integrity"] = integrity
        self._state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")

    def _fingerprint(self) -> dict:
        """sha256 + size + mtime_ns of the current model file (hashes the file)."""
        st = self.model_path.stat()
        return {
            "sha256":   _sha256_file(self.model_path),
            "size":     st.st_size,
            "mtime_ns": st.st_mtime_ns,
        }

    def _check_due(self, state: dict) -> bool:
        raw = state.get("last_check")
        if not raw:
            return True
        try:
            elapsed = (datetime.now(timezone.utc) - datetime.fromisoformat(raw)).total_seconds()
            return elapsed > _CHECK_INTERVAL
        except Exception:
            return True

    # ── cached-model integrity ──────────────────────────────────────────────────

    def _verify_integrity(self, verbose: bool) -> None:
        """
        Verify the cached model still matches what we downloaded, cheaply.

        The integrity guarantee is the recorded **sha256** — not size/mtime.
        size+mtime is only a cache-invalidation hint (cheaply forgeable: an
        attacker can reset mtime and match a byte length), used to decide whether
        we may *skip* re-reading 4.7 GB. When they're unchanged we trust the file
        WITHOUT re-hashing (the fast path); when they differ — or when the user
        sets GREPXCEL_VERIFY_MODEL=full — we recompute the full sha256 and compare.
        A hash mismatch is fatal (the weights fed to inference are not what we
        fetched) unless the user opted out via --allow-unverified-model /
        GREPXCEL_ALLOW_UNVERIFIED_MODEL. A file with no recorded fingerprint
        (placed manually, or state.json lost) can't be verified — we warn loudly
        and proceed, since the user may have supplied it deliberately.
        """
        state = self._load_state()
        integ = state.get("integrity") or {}
        stored_hash = integ.get("sha256")
        known = _known_hashes_for(MODEL_FILENAME)

        # Case C — no recorded fingerprint (manual file / lost or stale state).
        if not stored_hash:
            # Code-pinned trust anchor: if we know the good hash for this model,
            # verify against it directly — no local fingerprint needed, and it
            # works identically on a fresh machine (the portable case). A match
            # is silently trusted and recorded; a non-match warns but still
            # proceeds (a file with no fingerprint may have been placed
            # deliberately), unless silenced with --allow-unverified-model.
            if known:
                actual = _sha256_file(self.model_path)
                if actual in known:
                    self._save_state(state.get("commit_hash", MODEL_REVISION),
                                     integrity=self._fingerprint())
                    if verbose:
                        print("Model integrity: matches a known-good release.",
                              file=sys.stderr)
                    return
                if not self._allow_unverified:
                    print(
                        f"  [!] Cached model at {self.model_path} has no recorded "
                        f"checksum and does not match a known-good release "
                        f"(got {actual[:12]}…).\n"
                        f"      Proceeding on trust. Delete it to force a verified "
                        f"re-download, or pass --allow-unverified-model to silence "
                        f"this.",
                        file=sys.stderr,
                    )
                return
            if not self._allow_unverified:
                print(
                    f"  [!] Cannot verify the cached model at {self.model_path}: no "
                    f"recorded checksum (placed manually, or download state lost).\n"
                    f"      Proceeding on trust. Delete it to force a verified "
                    f"re-download, or pass --allow-unverified-model to silence this.",
                    file=sys.stderr,
                )
            return

        st = self.model_path.stat()
        unchanged = (
            st.st_size == integ.get("size") and st.st_mtime_ns == integ.get("mtime_ns")
        )

        # Case A — fast path. size+mtime unchanged AND no full-verify demand:
        # trust without hashing 4.7 GB. (size+mtime is a hint, never the proof.)
        # NOTE: on shared/multi-user or server deployments where the cache dir is
        # writable by others, set GREPXCEL_VERIFY_MODEL=full so the sha256 is
        # re-checked every run — size+mtime can be forged by a local attacker.
        if unchanged and not _full_verify_enabled():
            if verbose:
                print("Model integrity: unchanged since download.", file=sys.stderr)
            return

        # Re-hash: either something moved, or GREPXCEL_VERIFY_MODEL=full was set.
        if verbose:
            why = "full verification requested" if unchanged else "file changed on disk"
            print(f"Verifying model checksum ({why})...", file=sys.stderr)
        actual = _sha256_file(self.model_path)

        # Case B(i) — content matches what we downloaded; refresh the stat hint
        # if metadata moved (a touch/copy), then proceed.
        if actual == stored_hash:
            if not unchanged:
                self._save_state(state.get("commit_hash", MODEL_REVISION),
                                 integrity=self._fingerprint())
            return

        # Case B(i-known) — the recorded fingerprint is stale/wrong, but the file
        # matches a code-pinned known-good release. Trust it and heal the local
        # record. (This is the cross-machine case: a state.json copied from
        # another device can record a different build than the authentic file.)
        if known and actual in known:
            self._save_state(state.get("commit_hash", MODEL_REVISION),
                             integrity=self._fingerprint())
            if verbose:
                print("Model integrity: matches a known-good release "
                      "(healed a stale local record).", file=sys.stderr)
            return

        # Case B(ii) — content differs from what we downloaded.
        msg = (
            f"cached model at {self.model_path} does not match the checksum "
            f"recorded at download time (expected {stored_hash[:12]}…, got "
            f"{actual[:12]}…). The file was modified, corrupted, or replaced."
        )
        if self._allow_unverified:
            print(f"  [!] WARNING: {msg}\n      Continuing because "
                  f"--allow-unverified-model was set.", file=sys.stderr)
            return
        print(
            f"Error: {msg}\n"
            f"Refusing to run a tampered/corrupted model. To fix, delete it and "
            f"re-download:\n    rm {self.model_path}\n"
            f"Or, if you trust this file, re-run with --allow-unverified-model.",
            file=sys.stderr,
        )
        sys.exit(1)

    # ── remote metadata (lightweight — no model download) ─────────────────────

    def _remote_commit(self) -> str | None:
        """
        Fetch the latest commit hash for the model repo from HuggingFace.
        Returns None on any network or API error — callers treat this as
        'network unavailable, keep using cached model'.
        """
        try:
            from huggingface_hub import HfApi
            return HfApi().model_info(MODEL_REPO_ID).sha
        except Exception:
            return None

    # ── update logic ──────────────────────────────────────────────────────────

    def _maybe_update(self, verbose: bool) -> None:
        state = self._load_state()
        if not self._check_due(state):
            return

        if verbose:
            print("Checking for model updates...", file=sys.stderr)

        remote = self._remote_commit()

        if remote is None:
            if verbose:
                print(
                    "Could not reach HuggingFace — using cached model.",
                    file=sys.stderr,
                )
            return

        if remote != state.get("commit_hash"):
            print("New model version available — updating...", file=sys.stderr)
            self._download(announce=True, commit_hash=remote)
        else:
            self._save_state(remote)
            if verbose:
                print("Model is up to date.", file=sys.stderr)

    # ── download ──────────────────────────────────────────────────────────────

    def _download(self, announce: bool = True, commit_hash: str | None = None) -> None:
        """
        Download the model at a specific revision to a temporary file, then
        atomically replace the cached copy.  The old model remains available
        until the moment of the atomic rename, so a running web service is
        never interrupted.

        revision = commit_hash when given (auto-update path), otherwise the
        pinned MODEL_REVISION. hf_hub_download verifies the file integrity hash
        against the Hub for that revision.
        """
        _require_huggingface_hub()

        revision = commit_hash or MODEL_REVISION

        if announce:
            print(
                f"Downloading {MODEL_FILENAME} ({_SIZE_HINT}) at revision {revision[:12]}...",
                file=sys.stderr,
            )
            print(
                "First run only. Live progress is shown below. Tip: set HF_TOKEN "
                "(https://huggingface.co/settings/tokens) for faster, rate-limit-free "
                "downloads.",
                file=sys.stderr,
            )

        from huggingface_hub import hf_hub_download

        # Download into a throw-away temp dir inside cache_dir so that
        # os.replace() is guaranteed to be on the same filesystem (atomic).
        with tempfile.TemporaryDirectory(dir=self.cache_dir, prefix="_dl_") as tmp_dir:
            # Watchdog: heartbeats (non-TTY) + stall warnings, so a hung
            # anonymous download is never silent. Only when we announce — the
            # mocked test path uses announce=False and starts no thread.
            stop = threading.Event()
            watcher: threading.Thread | None = None
            if announce:
                watcher = threading.Thread(
                    target=self._watch_download, args=(tmp_dir, stop), daemon=True,
                )
                watcher.start()
            try:
                downloaded = hf_hub_download(
                    repo_id=MODEL_REPO_ID,
                    filename=MODEL_FILENAME,
                    revision=revision,
                    local_dir=tmp_dir,
                )
            finally:
                stop.set()
                if watcher is not None:
                    watcher.join(timeout=2)
            os.replace(downloaded, str(self.model_path))

        # Record an integrity fingerprint of exactly what we just downloaded
        # (hf_hub_download already verified the hash against the Hub). Later runs
        # compare against this to detect post-download tampering/corruption.
        self._save_state(revision, integrity=self._fingerprint())

        if announce:
            print(f"Model ready: {self.model_path}", file=sys.stderr)

    # ── download progress watchdog ──────────────────────────────────────────────

    def _watch_download(self, tmp_dir: str, stop: threading.Event) -> None:
        """Watch the in-flight download; emit heartbeats / stall warnings.

        Runs in a daemon thread until `stop` is set. huggingface_hub already
        draws a live bar in a TTY, so heartbeats are emitted only when stderr is
        NOT a TTY (keeps an interactive bar clean). A stall warning fires in
        either case — that's the actionable signal when an anonymous, throttled
        HF download hangs with no bytes moving.
        """
        is_tty       = sys.stderr.isatty()
        now          = time.monotonic()
        last_size    = 0
        last_change  = now
        last_beat    = now
        last_stall   = 0.0

        while not stop.wait(_DL_POLL_SECONDS):
            size = _largest_file_bytes(tmp_dir)
            now  = time.monotonic()

            if size > last_size:
                last_size, last_change = size, now

            if size <= 0:
                continue

            mb = size / (1024 * 1024)

            if not is_tty and (now - last_beat) >= _DL_HEARTBEAT_SECONDS:
                print(f"  ... {mb:,.0f} MB downloaded", file=sys.stderr)
                last_beat = now

            stalled_for = now - last_change
            if stalled_for >= _DL_STALL_SECONDS and (now - last_stall) >= _DL_STALL_SECONDS:
                print(
                    f"  [!] No download progress for {int(stalled_for)}s "
                    f"(stuck at {mb:,.0f} MB). The HuggingFace anonymous rate limit "
                    f"may be throttling — set HF_TOKEN "
                    f"(https://huggingface.co/settings/tokens) and re-run.",
                    file=sys.stderr,
                )
                last_stall = now


# ── dependency guard ──────────────────────────────────────────────────────────

def _require_huggingface_hub() -> None:
    try:
        import huggingface_hub  # noqa: F401
    except ImportError:
        print(
            "Error: the local draft backend needs the 'suggest' extra "
            "(huggingface_hub + llama-cpp-python), which isn't installed.\n"
            "Fix:   pip install 'grepxcel[suggest]'\n"
            "       (or: python3 scripts/install_llm_deps.py  — autodetects GPU)\n"
            "Or skip the local model with a cloud backend: "
            "grepxcel draft --backend claude ...",
            file=sys.stderr,
        )
        sys.exit(1)
