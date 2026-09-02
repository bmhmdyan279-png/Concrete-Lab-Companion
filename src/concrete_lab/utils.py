"""Small shared utilities: hashing, timestamps, build identifiers."""

from __future__ import annotations

import hashlib
import uuid
from datetime import datetime, timezone
from pathlib import Path

#: Chunk size for streaming file hashes (keeps memory flat for big files).
READ_CHUNK_SIZE: int = 8192


def compute_sha256(filepath: Path) -> str:
    """Compute the SHA-256 hex digest of a file, streamed in chunks.

    Args:
        filepath: Existing file to hash.

    Returns:
        Lowercase hex digest.

    Raises:
        FileNotFoundError: If the file does not exist.
    """
    digest = hashlib.sha256()
    with open(filepath, "rb") as handle:
        while chunk := handle.read(READ_CHUNK_SIZE):
            digest.update(chunk)
    return digest.hexdigest()


def utc_now_iso() -> str:
    """Return the current UTC time as an ISO-8601 string (seconds)."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def new_build_id() -> str:
    """Return a fresh unique build identifier (short UUID)."""
    return uuid.uuid4().hex[:12]
