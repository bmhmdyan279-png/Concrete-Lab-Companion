"""Tests for the build manifest (traceability metadata)."""

from concrete_lab import __version__
from concrete_lab.qa.engine import QAReport
from concrete_lab.report.manifest import REQUIRED_MANIFEST_KEYS, build_manifest
from concrete_lab.standards import registry as standards_registry


def _manifest(**overrides) -> dict:
    defaults = dict(
        filename="Concrete_Lab_Companion_v4.0.0.xlsx",
        sha256="ab" * 32,
        sheets=("00_راهنما", "17_آزمایش_4-1"),
        qa_reports={
            "golden": QAReport("golden-suite", passed=14, failed=0, warnings=17),
            "structural": QAReport("structural", passed=6, failed=0),
        },
        protection_enabled=True,
    )
    defaults.update(overrides)
    return build_manifest(**defaults)


class TestManifest:
    def test_all_required_keys_present(self) -> None:
        manifest = _manifest()
        for key in REQUIRED_MANIFEST_KEYS:
            assert key in manifest, f"missing required key {key}"

    def test_application_and_environment_identity(self) -> None:
        manifest = _manifest()
        assert manifest["application_version"] == __version__
        assert manifest["build_id"]  # non-empty unique id
        assert manifest["python_version"]
        assert manifest["openpyxl_version"]
        assert manifest["build_time"].endswith("Z")

    def test_standards_map_uses_audited_editions(self) -> None:
        manifest = _manifest()
        assert manifest["standards"]["ASTM C39/C39M"] == "2026"  # audited, was 2022
        assert manifest["standards"]["ASTM C805/C805M"] == "2025"
        assert len(manifest["standards"]) == len(standards_registry.all_standards())

    def test_rulesets_carry_version_tags_and_scope(self) -> None:
        manifest = _manifest()
        assert manifest["rulesets"]["C39"]["ruleset_version"] == "C39-2026-v1"
        assert manifest["rulesets"]["C805"]["scope"] == "estimation"
        assert manifest["rulesets"]["ISIRI302"]["edition"] == "1394"

    def test_qa_section_combines_tiers(self) -> None:
        manifest = _manifest()
        qa = manifest["qa"]
        assert qa["passed"] == 20
        assert qa["failed"] == 0
        assert qa["warnings"] == 17
        assert qa["status"] == "warn"  # warnings keep it honest
        assert set(qa["tiers"]) == {"golden", "structural"}

    def test_failed_tier_flips_qa_status(self) -> None:
        manifest = _manifest(qa_reports={
            "golden": QAReport("golden-suite", passed=1, failed=2, failures=("x", "y")),
        })
        assert manifest["qa"]["status"] == "fail"
        assert manifest["qa"]["failed"] == 2

    def test_tests_progress_counts(self) -> None:
        manifest = _manifest()
        assert manifest["tests"]["total"] == 20
        assert manifest["tests"]["implemented"] == 7
        assert manifest["tests"]["pending"] == 13

    def test_protection_note_reminds_not_security(self) -> None:
        manifest = _manifest()
        assert manifest["protection"]["enabled"] is True
        assert "not a security boundary" in manifest["protection"]["note"]
