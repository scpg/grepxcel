"""
Model manager: ensures the pinned GGUF model is present locally,
checks HuggingFace for a newer version once per day, and downloads atomically.

All network activity is limited to:
  - one lightweight metadata call per day (commit-hash check)
  - one file download when a new version is detected

Inference never touches the network — it runs entirely in-process via llama-cpp-python.
"""

import json
import os
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

# ── Pinned model ──────────────────────────────────────────────────────────────
# To upgrade the bundled model, change these three constants only.

MODEL_REPO_ID    = "bartowski/Phi-3.5-mini-instruct-GGUF"
MODEL_FILENAME   = "Phi-3.5-mini-instruct-Q4_K_M.gguf"
MODEL_CHAT_FORMAT = "phi-3"

_SIZE_HINT        = "~2.4 GB"
_CHECK_INTERVAL   = 86_400          # seconds — 24 h


# ── Cache location ────────────────────────────────────────────────────────────

def default_cache_dir() -> Path:
    """
    Returns the directory where the model file and state are stored.
    Override with the GREPXCEL_MODEL_DIR environment variable
    (useful for Docker / web-service deployments with a mounted volume).
    """
    base = os.environ.get("GREPXCEL_MODEL_DIR")
    return Path(base) if base else Path.home() / ".cache" / "grepxcel" / "models"


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
        Guarantee the model file is present and checked for updates.
        Returns the path to the model file — always valid on return.
        """
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        if not self.model_path.exists():
            self._download(announce=True)
        else:
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
        Download the model to a temporary file, then atomically replace
        the cached copy.  The old model remains available until the moment
        of the atomic rename, so a running web service is never interrupted.
        """
        _require_huggingface_hub()

        if announce:
            print(
                f"Downloading {MODEL_FILENAME} ({_SIZE_HINT})...",
                file=sys.stderr,
            )
            print(
                "This only happens on first run or when a new version is released.",
                file=sys.stderr,
            )

        from huggingface_hub import hf_hub_download

        # Download into a throw-away temp dir inside cache_dir so that
        # os.replace() is guaranteed to be on the same filesystem (atomic).
        with tempfile.TemporaryDirectory(dir=self.cache_dir, prefix="_dl_") as tmp_dir:
            downloaded = hf_hub_download(
                repo_id=MODEL_REPO_ID,
                filename=MODEL_FILENAME,
                local_dir=tmp_dir,
            )
            os.replace(downloaded, str(self.model_path))

        if commit_hash is None:
            commit_hash = self._remote_commit() or ""
        self._save_state(commit_hash)

        if announce:
            print(f"Model ready: {self.model_path}", file=sys.stderr)


# ── dependency guard ──────────────────────────────────────────────────────────

def _require_huggingface_hub() -> None:
    try:
        import huggingface_hub  # noqa: F401
    except ImportError:
        print(
            "Error: 'huggingface_hub' is not installed.\n"
            "Fix:   pip install huggingface_hub",
            file=sys.stderr,
        )
        sys.exit(1)
