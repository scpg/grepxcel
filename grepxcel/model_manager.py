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

All network activity is limited to:
  - one file download of the pinned revision on first run
  - (only when auto-update is opted in) one lightweight commit-check per day,
    plus a download if a newer upstream commit is found

Inference never touches the network — it runs entirely in-process via llama-cpp-python.
"""

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

# Qwen2.5-Coder-7B-Instruct — the best LOCAL model in our execution-based draft
# eval (claude-opus is better but cloud-only). It still hits a capability ceiling
# on the most complex tables, but clearly beats the previous default (Phi-3.5).
MODEL_REPO_ID    = "bartowski/Qwen2.5-Coder-7B-Instruct-GGUF"
MODEL_FILENAME   = "Qwen2.5-Coder-7B-Instruct-Q4_K_M.gguf"
MODEL_REVISION   = "1f629da0c8bed16b9e50cee91c70693650e66c35"  # pinned commit
MODEL_CHAT_FORMAT = None  # auto-detect from GGUF metadata (Qwen embeds a ChatML template)

_SIZE_HINT        = "~4.7 GB"
_CHECK_INTERVAL   = 86_400          # seconds — 24 h
_AUTOUPDATE_ENV   = "GREPXCEL_MODEL_AUTOUPDATE"

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

    def __init__(self, cache_dir: Path | None = None):
        self.cache_dir  = Path(cache_dir) if cache_dir else default_cache_dir()
        self.model_path = self.cache_dir / MODEL_FILENAME
        self._state_path = self.cache_dir / self._STATE_FILE

    def ensure_ready(self, verbose: bool = False) -> Path:
        """
        Guarantee the model file is present (downloading the pinned revision on
        first run). Update checks happen only when auto-update is opted in via
        the GREPXCEL_MODEL_AUTOUPDATE env var.
        Returns the path to the model file — always valid on return.
        """
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        if not self.model_path.exists():
            self._download(announce=True)
        elif _autoupdate_enabled():
            self._maybe_update(verbose)

        return self.model_path

    # ── state helpers ─────────────────────────────────────────────────────────

    def _load_state(self) -> dict:
        try:
            return json.loads(self._state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def _save_state(self, commit_hash: str) -> None:
        state = {
            "last_check":  datetime.now(timezone.utc).isoformat(),
            "commit_hash": commit_hash,
            "filename":    MODEL_FILENAME,
        }
        self._state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")

    def _check_due(self, state: dict) -> bool:
        raw = state.get("last_check")
        if not raw:
            return True
        try:
            elapsed = (datetime.now(timezone.utc) - datetime.fromisoformat(raw)).total_seconds()
            return elapsed > _CHECK_INTERVAL
        except Exception:
            return True

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

        self._save_state(revision)

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
