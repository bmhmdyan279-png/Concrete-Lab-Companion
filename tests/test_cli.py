"""End-to-end tests: build.py orchestrator as a real subprocess."""

import json
import subprocess
import sys
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parents[1]
BUILD = str(REPO_ROOT / "build.py")


def run_build_py(*args: str, cwd: Path = REPO_ROOT) -> subprocess.CompletedProcess:
    """Run the orchestrator exactly as a user would."""
    return subprocess.run(
        [sys.executable, BUILD, *args],
        cwd=str(cwd), capture_output=True, text=True, timeout=120,
    )


class TestBuildEndToEnd:
    def test_build_produces_verified_artifacts(self, tmp_path: Path) -> None:
        outcome = run_build_py("--output", str(tmp_path), "--no-protect")
        assert outcome.returncode == 0, outcome.stderr

        xlsx = tmp_path / "Concrete_Lab_Companion_v4.0.0.xlsx"
        sha_file = tmp_path / "Concrete_Lab_Companion_v4.0.0.xlsx.sha256"
        manifest = tmp_path / "Concrete_Lab_Companion_v4.0.0.json"
        assert xlsx.exists() and sha_file.exists() and manifest.exists()

        # the manifest the build finalised includes the structural tier
        data = json.loads(manifest.read_text(encoding="utf-8"))
        assert data["application_version"] == "4.0.0"
        assert data["qa"]["failed"] == 0
        assert "structural" in data["qa"]["tiers"]
        assert data["standards"]["ASTM C39/C39M"] == "2026"

        # the delivered workbook is display-only (no formulas)
        workbook = openpyxl.load_workbook(str(xlsx))
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    assert not (isinstance(cell.value, str) and cell.value.startswith("="))

    def test_sha256_matches_artifact(self, tmp_path: Path) -> None:
        from concrete_lab.utils import compute_sha256

        assert run_build_py("--output", str(tmp_path), "--no-protect").returncode == 0
        xlsx = tmp_path / "Concrete_Lab_Companion_v4.0.0.xlsx"
        recorded = (tmp_path / "Concrete_Lab_Companion_v4.0.0.xlsx.sha256") \
            .read_text(encoding="utf-8").split()[0]
        assert compute_sha256(xlsx) == recorded


class TestValidateEndToEnd:
    def test_validate_runs_real_qa_and_passes(self) -> None:
        # restricted path keeps the nested run fast and recursion-free
        outcome = run_build_py("--validate", "tests/test_domain.py")
        assert outcome.returncode == 0, outcome.stderr
        assert "VALIDATION MODE" in outcome.stderr

    def test_validate_fails_when_a_test_fails(self, tmp_path: Path) -> None:
        bad_test = tmp_path / "test_forced_failure.py"
        bad_test.write_text("def test_broken():\n    assert False\n", encoding="utf-8")
        outcome = run_build_py("--validate", str(bad_test))
        assert outcome.returncode == 1


class TestCliSurface:
    def test_version_flag(self) -> None:
        outcome = run_build_py("--version")
        assert outcome.returncode == 0
        assert "4.0.0" in outcome.stdout
