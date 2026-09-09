#!/usr/bin/env python3
"""Publish the release identity the landing page displays.

The GitHub Pages site used to carry a **hard-coded** SHA-256 digest.  Every
rebuild changed the workbook but not the page, so a user who downloaded a
new artifact and compared hashes got a false "file has been tampered with"
alarm — the exact opposite of what the section is for.

This script is the fix: it reads the manifest a real build produced and
writes a compact ``release.json`` next to the landing page.  The page
fetches that file at load time, so version, artifact name, checksum and QA
outcome are always the ones of the build that was actually published.

Usage::

    python build.py --output output --no-protect
    python scripts/make_release_manifest.py --output output --dest landing/release.json

    # CI guard: fail when the published page disagrees with a fresh build
    python scripts/make_release_manifest.py --output output --dest landing/release.json --check

Exit codes: ``0`` success (or ``--check`` found no drift), ``1`` anything
else.  Every message is written to stderr so it shows up in CI logs.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT / "src") not in sys.path:
    sys.path.insert(0, str(REPO_ROOT / "src"))

from concrete_lab import __version__  # noqa: E402
from concrete_lab.cli import ARTIFACT_PATTERN  # noqa: E402
from concrete_lab.utils import utc_now_iso  # noqa: E402

#: Length of a SHA-256 hex digest.
SHA256_HEX_LENGTH: int = 64


class PublishError(ValueError):
    """Raised when a release manifest cannot be published."""


def _find_manifest(output_dir: Path, version: str) -> Path:
    """Locate the build manifest for ``version`` inside ``output_dir``."""
    expected = output_dir / (Path(ARTIFACT_PATTERN.format(version=version)).stem + ".json")
    if expected.is_file():
        return expected
    candidates = sorted(output_dir.glob("*.json"))
    if not candidates:
        raise PublishError(
            f"no manifest JSON in {output_dir} — run `python build.py --output {output_dir}` first"
        )
    if len(candidates) > 1:
        names = ", ".join(path.name for path in candidates)
        raise PublishError(f"ambiguous manifest in {output_dir}: {names}")
    return candidates[0]


def build_release_info(output_dir: Path, version: str = __version__) -> Dict[str, Any]:
    """Derive the landing-page payload from a real build manifest.

    Args:
        output_dir: Directory produced by ``python build.py --output …``.
        version: Package version the artifact must belong to.

    Returns:
        A JSON-serialisable mapping with the fields the page renders.

    Raises:
        PublishError: If the manifest is missing, malformed, belongs to a
            different version, or carries an implausible checksum.
    """
    output_dir = Path(output_dir)
    manifest_path = _find_manifest(output_dir, version)

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        raise PublishError(f"{manifest_path} is not readable JSON: {exc}") from exc

    manifest_version = str(manifest.get("application_version", ""))
    if manifest_version != version:
        raise PublishError(
            f"{manifest_path.name} was built by version {manifest_version!r} but the "
            f"package is {version!r} — rebuild before publishing"
        )

    sha = str(manifest.get("sha256", ""))
    if len(sha) != SHA256_HEX_LENGTH or any(char not in "0123456789abcdef" for char in sha):
        raise PublishError(f"{manifest_path.name} carries an implausible sha256: {sha!r}")

    # The sidecar checksum must agree with the manifest; a disagreement means
    # the workbook was touched after the manifest was written.
    sidecar = output_dir / (str(manifest.get("filename", "")) + ".sha256")
    if sidecar.is_file():
        fields = sidecar.read_text(encoding="utf-8").split()
        recorded = fields[0] if fields else ""
        if recorded != sha:
            raise PublishError(
                f"{sidecar.name} says {recorded!r} but the manifest says {sha!r} — "
                f"the artifact changed after it was hashed"
            )

    qa = manifest.get("qa") or {}
    tests = manifest.get("tests") or {}
    return {
        "version": manifest_version,
        "filename": manifest.get("filename"),
        "sha256": sha,
        "build_time": manifest.get("build_time"),
        "data_source": manifest.get("data_source"),
        "qa_status": qa.get("status"),
        "qa": {
            "passed": qa.get("passed", 0),
            "failed": qa.get("failed", 0),
            "warnings": qa.get("warnings", 0),
        },
        "tests": {
            "total": tests.get("total", 0),
            "implemented": tests.get("implemented", 0),
            "pending": tests.get("pending", 0),
        },
        "protection_enabled": (manifest.get("protection") or {}).get("enabled", False),
        "source_manifest": manifest_path.name,
        "generated_at": utc_now_iso(),
        "generator": "scripts/make_release_manifest.py",
    }


def write_release_info(info: Dict[str, Any], dest: Path) -> None:
    """Write the payload as pretty, LF-terminated JSON."""
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(info, indent=2, ensure_ascii=False, sort_keys=True) + "\n",
                    encoding="utf-8")


def check_release_info(info: Dict[str, Any], dest: Path) -> Optional[str]:
    """Return a drift description, or ``None`` when the page is up to date.

    The workbook embeds its build timestamp, so **every** rebuild produces a
    different checksum.  Comparing checksums here would make the gate red on
    every local run — and a gate that always fails gets ignored, which is the
    exact failure mode this script exists to prevent.  Drift is therefore
    judged on the fields that identify *what* is being published and are
    stable across rebuilds: version, artifact name, data source and
    implementation progress.

    Correctness of the checksum itself is guaranteed structurally instead: CI
    builds once and hands the same bytes to the Release and to Pages (see
    ``.github/workflows/build.yml``), so the published digest describes the
    uploaded artifact by construction rather than by comparison.
    """
    dest = Path(dest)
    if not dest.is_file():
        return f"{dest} does not exist — run without --check to create it"
    try:
        published = json.loads(dest.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        return f"{dest} is not readable JSON: {exc}"

    drift = []
    for key in ("version", "filename", "data_source"):
        if published.get(key) != info.get(key):
            drift.append(f"{key}: published {published.get(key)!r} != built {info.get(key)!r}")
    if published.get("tests") != info.get("tests"):
        drift.append(f"tests: published {published.get('tests')!r} != built {info.get('tests')!r}")
    return "; ".join(drift) if drift else None


def main(argv: Optional[list] = None) -> int:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="make_release_manifest",
        description="Publish the release identity (version, artifact, SHA-256, QA) "
                    "consumed by the landing page.",
    )
    parser.add_argument("--output", default="output",
                        help="build output directory (default: ./output)")
    parser.add_argument("--dest", default="landing/release.json",
                        help="where to write the landing payload (default: landing/release.json)")
    parser.add_argument("--check", action="store_true",
                        help="do not write; exit 1 when --dest disagrees with the build")
    args = parser.parse_args(argv)

    try:
        info = build_release_info(Path(args.output))
    except PublishError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    dest = Path(args.dest)
    if args.check:
        drift = check_release_info(info, dest)
        if drift:
            print(f"ERROR: landing page is out of date — {drift}", file=sys.stderr)
            return 1
        print(f"OK: {dest} matches the build of {info['version']}", file=sys.stderr)
        return 0

    write_release_info(info, dest)
    print(f"Wrote {dest} (version {info['version']}, sha256 {info['sha256'][:12]}…)", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
