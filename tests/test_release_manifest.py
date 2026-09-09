"""Tests for ``scripts/make_release_manifest.py`` (the landing-page publisher).

The script exists because a hand-typed SHA-256 on the project website is a
maintenance trap: the artifact changes on every rebuild, the page does not,
and an honest user then sees a false tamper warning.  These tests cover the
identity rules that make the published page trustworthy.
"""

import importlib.util
import json
from pathlib import Path
from typing import Any, Dict, Optional

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "scripts" / "make_release_manifest.py"


def load_script():
    """Import the script as a module (it lives outside the package)."""
    spec = importlib.util.spec_from_file_location("make_release_manifest", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def publisher():
    """The publisher module, imported once per test session."""
    return load_script()


def write_artifacts(output_dir: Path, *, version: Optional[str] = None,
                    sha: str = "ab" * 32, sidecar: bool = True,
                    filename: str = "Concrete_Lab_Companion_vX.xlsx") -> Path:
    """Create a minimal but realistic build output directory."""
    from concrete_lab import __version__

    version = version or __version__
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest: Dict[str, Any] = {
        "application_version": version,
        "filename": filename,
        "sha256": sha,
        "build_time": "2026-09-08T00:00:00Z",
        "data_source": "demo",
        "qa": {"status": "warn", "passed": 23, "failed": 0, "warnings": 14},
        "tests": {"total": 20, "implemented": 7, "pending": 13},
        "protection": {"enabled": False},
    }
    manifest_path = output_dir / (Path(filename).stem + ".json")
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    if sidecar:
        (output_dir / (filename + ".sha256")).write_text(f"{sha}  {filename}\n", encoding="utf-8")
    return manifest_path


class TestPayloadGeneration:
    def test_payload_carries_the_release_identity(self, publisher, tmp_path: Path) -> None:
        from concrete_lab import __version__

        write_artifacts(tmp_path, filename=f"Concrete_Lab_Companion_v{__version__}.xlsx")
        info = publisher.build_release_info(tmp_path)

        assert info["version"] == __version__
        assert info["filename"] == f"Concrete_Lab_Companion_v{__version__}.xlsx"
        assert info["sha256"] == "ab" * 32
        assert info["data_source"] == "demo"
        assert info["qa_status"] == "warn"
        assert info["tests"] == {"total": 20, "implemented": 7, "pending": 13}
        assert info["generator"] == "scripts/make_release_manifest.py"

    def test_payload_is_json_serialisable_and_sorted(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path)
        info = publisher.build_release_info(tmp_path)
        text = json.dumps(info, indent=2, ensure_ascii=False, sort_keys=True)
        assert json.loads(text) == info
        keys = list(json.loads(text))
        assert keys == sorted(keys)


class TestRefusals:
    def test_missing_output_directory_is_refused(self, publisher, tmp_path: Path) -> None:
        with pytest.raises(publisher.PublishError, match="no manifest JSON"):
            publisher.build_release_info(tmp_path / "nowhere")

    def test_stale_version_is_refused(self, publisher, tmp_path: Path) -> None:
        """Publishing a page for a build that is not the current release is a defect."""
        write_artifacts(tmp_path, version="0.0.1")
        with pytest.raises(publisher.PublishError, match="rebuild before publishing"):
            publisher.build_release_info(tmp_path)

    def test_implausible_checksum_is_refused(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path, sha="not-a-hash")
        with pytest.raises(publisher.PublishError, match="implausible sha256"):
            publisher.build_release_info(tmp_path)

    def test_sidecar_disagreement_is_refused(self, publisher, tmp_path: Path) -> None:
        """Manifest and ``.sha256`` must describe the same bytes."""
        write_artifacts(tmp_path, sha="ab" * 32)
        manifest = next(tmp_path.glob("*.json"))
        data = json.loads(manifest.read_text(encoding="utf-8"))
        data["sha256"] = "cd" * 32
        manifest.write_text(json.dumps(data), encoding="utf-8")
        with pytest.raises(publisher.PublishError, match="changed after it was hashed"):
            publisher.build_release_info(tmp_path)

    def test_ambiguous_output_is_refused(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path, filename="a.xlsx")
        write_artifacts(tmp_path, filename="b.xlsx")
        with pytest.raises(publisher.PublishError, match="ambiguous manifest"):
            publisher.build_release_info(tmp_path)


class TestDriftDetection:
    def test_check_reports_a_stale_page(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path / "build")
        info = publisher.build_release_info(tmp_path / "build")

        dest = tmp_path / "release.json"
        publisher.write_release_info(info, dest)
        assert publisher.check_release_info(info, dest) is None

        for stale_field, stale_value in (("version", "9.9.9"),
                                         ("filename", "other.xlsx"),
                                         ("data_source", "user-input")):
            drift = publisher.check_release_info(dict(info, **{stale_field: stale_value}), dest)
            assert drift and stale_field in drift, f"{stale_field} drift went unnoticed"

    def test_a_rebuilt_checksum_is_not_drift(self, publisher, tmp_path: Path) -> None:
        """Rebuilds change the digest; that must not make the gate red forever."""
        write_artifacts(tmp_path)
        info = publisher.build_release_info(tmp_path)
        dest = tmp_path / "release.json"
        publisher.write_release_info(info, dest)
        assert publisher.check_release_info(dict(info, sha256="ff" * 32), dest) is None

    def test_check_reports_a_missing_page(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path)
        info = publisher.build_release_info(tmp_path)
        drift = publisher.check_release_info(info, tmp_path / "absent.json")
        assert drift and "does not exist" in drift

    def test_check_reports_a_corrupt_page(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path)
        info = publisher.build_release_info(tmp_path)
        dest = tmp_path / "release.json"
        dest.write_text("{not json", encoding="utf-8")
        assert "not readable JSON" in (publisher.check_release_info(info, dest) or "")

    def test_build_time_alone_is_not_drift(self, publisher, tmp_path: Path) -> None:
        """Rebuilds differ in timestamp; that must not look like a stale page."""
        write_artifacts(tmp_path)
        info = publisher.build_release_info(tmp_path)
        dest = tmp_path / "release.json"
        publisher.write_release_info(dict(info, build_time="1999-01-01T00:00:00Z",
                                          generated_at="1999-01-01T00:00:00Z"), dest)
        assert publisher.check_release_info(info, dest) is None


class TestCommandLine:
    def test_write_then_check_exits_zero(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path / "build")
        dest = tmp_path / "release.json"
        assert publisher.main(["--output", str(tmp_path / "build"), "--dest", str(dest)]) == 0
        assert dest.is_file()
        assert publisher.main(["--output", str(tmp_path / "build"), "--dest", str(dest),
                               "--check"]) == 0

    def test_check_without_writing_exits_one(self, publisher, tmp_path: Path) -> None:
        write_artifacts(tmp_path / "build")
        code = publisher.main(["--output", str(tmp_path / "build"),
                               "--dest", str(tmp_path / "absent.json"), "--check"])
        assert code == 1

    def test_unusable_build_exits_one(self, publisher, tmp_path: Path) -> None:
        assert publisher.main(["--output", str(tmp_path / "empty"),
                               "--dest", str(tmp_path / "release.json")]) == 1


class TestCommittedPageIsCurrent:
    def test_landing_release_json_passes_its_own_drift_check(self, publisher) -> None:
        """The committed page payload must describe the current package."""
        release = json.loads((REPO_ROOT / "landing" / "release.json").read_text(encoding="utf-8"))
        from concrete_lab import __version__

        assert release["version"] == __version__
        assert release["sha256"] and len(release["sha256"]) == publisher.SHA256_HEX_LENGTH

    def test_script_is_importable_without_side_effects(self) -> None:
        """Importing the publisher must not build, write or exit."""
        before = {path.name for path in REPO_ROOT.iterdir()}
        module = load_script()
        assert callable(module.build_release_info) and callable(module.main)
        assert {path.name for path in REPO_ROOT.iterdir()} == before, (
            "importing the script created files in the repository root"
        )

    def test_script_guards_its_entry_point(self) -> None:
        """``import`` must never run ``main()`` — the tests rely on that."""
        source = SCRIPT.read_text(encoding="utf-8")
        assert 'if __name__ == "__main__":' in source
