"""Build manifest: reproducibility and traceability metadata.

Implements the reviewer's manifest model: application version, build
identity, interpreter/library versions, the standards + ruleset
revisions used, QA outcomes, and artifact integrity data.
"""

from __future__ import annotations

import platform
import sys
from typing import Dict, Mapping, Optional, Sequence

import openpyxl

from concrete_lab import __version__
from concrete_lab.qa.engine import QAReport
from concrete_lab.specs.base import TEST_REGISTRY, implemented_tests
from concrete_lab.standards import registry as standards_registry
from concrete_lab.utils import new_build_id, utc_now_iso

#: Keys every manifest must carry (enforced by the structural QA tier).
REQUIRED_MANIFEST_KEYS: tuple = (
    "application_version",
    "build_id",
    "build_time",
    "python_version",
    "openpyxl_version",
    "standards",
    "rulesets",
    "qa",
    "filename",
    "sha256",
    "sheets",
)


def build_manifest(
    *,
    filename: str,
    sha256: str,
    sheets: Sequence[str],
    qa_reports: Mapping[str, QAReport],
    protection_enabled: bool,
) -> Dict[str, object]:
    """Assemble the manifest dictionary for one build.

    Args:
        filename: Artifact file name (not the full path).
        sha256: Artifact checksum.
        sheets: Worksheet titles in workbook order.
        qa_reports: QA tier name → report (e.g. ``golden``,
            ``structural``).
        protection_enabled: Whether worksheet protection was applied.

    Returns:
        A JSON-serialisable manifest containing every key in
        :data:`REQUIRED_MANIFEST_KEYS`.
    """
    standards = {spec.name: spec.edition for spec in standards_registry.all_standards()}
    rulesets = {}
    for code, module in sorted(standards_registry.RULESETS.items()):
        spec = standards_registry.get(code)
        rulesets[code] = {
            "edition": spec.edition,
            "ruleset_version": getattr(module, "RULESET_VERSION", None),
            "scope": getattr(module, "SCOPE").value if hasattr(module, "SCOPE") else None,
        }

    combined: Optional[QAReport] = None
    qa_section: Dict[str, object] = {}
    for name, report in qa_reports.items():
        qa_section[name] = report.to_dict()
        combined = report if combined is None else combined.merged_with(report)

    implemented = {spec.id for spec in implemented_tests()}

    return {
        "application_version": __version__,
        "build_id": new_build_id(),
        "build_time": utc_now_iso(),
        "python_version": platform.python_version(),
        "interpreter": sys.implementation.name,
        "openpyxl_version": openpyxl.__version__,
        "standards": standards,
        "rulesets": rulesets,
        "qa": {
            "tiers": qa_section,
            "passed": combined.passed if combined else 0,
            "failed": combined.failed if combined else 0,
            "warnings": combined.warnings if combined else 0,
            "status": combined.status.value if combined else "pending",
        },
        "tests": {
            "total": len(TEST_REGISTRY),
            "implemented": len(implemented),
            "implemented_ids": sorted(implemented),
            "pending": len(TEST_REGISTRY) - len(implemented),
        },
        "protection": {
            "enabled": protection_enabled,
            "note": "Worksheet protection prevents accidental modification. "
                    "It is not a security boundary.",
        },
        "filename": filename,
        "sha256": sha256,
        "sheets": list(sheets),
    }
