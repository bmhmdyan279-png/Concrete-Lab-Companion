"""Command-line interface of the Concrete Lab Companion.

Two real modes (no phantom features):

* ``--validate`` — executes QA for real: the in-process golden suite
  plus the pytest suite, and exits non-zero on any failure;
* build (default) — computes demo results with the domain engine,
  renders them (display-only), runs structural QA over the artifact
  and writes the manifest + SHA-256.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Dict, Optional, Sequence

from concrete_lab import __version__
from concrete_lab.config import AppConfig
from concrete_lab.domain.engine import run_demo_cases
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.qa.engine import QAEngine, QAReport
from concrete_lab.qa.golden import run_golden_suite
from concrete_lab.qa.structural import run_structural_checks
from concrete_lab.render.excel.protection import ProtectionManager
from concrete_lab.render.excel.renderer import ExcelRenderer
from concrete_lab.render.excel.assembler import assemble_workbook_model
from concrete_lab.report.manifest import build_manifest
from concrete_lab.utils import compute_sha256

logger = logging.getLogger("concrete_lab.cli")

#: Artifact naming pattern.
ARTIFACT_PATTERN: str = "Concrete_Lab_Companion_v{version}.xlsx"


def build_arg_parser() -> argparse.ArgumentParser:
    """Create the CLI parser (kept separate for tests and help output)."""
    parser = argparse.ArgumentParser(
        prog="concrete-lab-companion",
        description="Concrete Lab Companion — standards-compliant calculation & QA engine",
    )
    parser.add_argument("--output", default="output", help="Output directory (default: ./output)")
    parser.add_argument("--no-protect", action="store_true", help="Disable sheet protection")
    parser.add_argument("--password-env", default="WORKBOOK_PASSWORD",
                        help="Environment variable holding the protection password")
    parser.add_argument("--demo", action="store_true",
                        help="Compatibility flag: builds always display engine-computed demo results")
    parser.add_argument("--validate", nargs="*", default=None, metavar="PATH",
                        help="Run real QA (golden suite + pytest) and exit; "
                             "optionally restrict pytest to specific paths")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _log_qa_report(report: QAReport) -> None:
    """Log one QA tier as a compact table."""
    logger.info("─" * 46)
    logger.info("QA %-14s passed=%d failed=%d warnings=%d (%.2fs)",
                report.source, report.passed, report.failed, report.warnings,
                report.duration_seconds)
    logger.info("   status: %s %s", report.status.symbol, report.status.name)
    for failure in report.failures[:10]:
        logger.info("   ✗ %s", failure)


def run_validation(paths: Optional[Sequence[str]] = None) -> int:
    """Execute every QA tier and report PASS/FAIL/WARN.

    Args:
        paths: Optional pytest path restriction (test files/dirs).

    Returns:
        Process exit code: ``0`` unless some tier FAILED.
    """
    logger.info("═" * 46)
    logger.info(" VALIDATION MODE — running real QA")
    logger.info("═" * 46)

    golden = run_golden_suite()
    _log_qa_report(golden)

    pytest_report = QAEngine().run_pytest(pytest_args=list(paths or ()))
    _log_qa_report(pytest_report)

    combined = golden.merged_with(pytest_report)
    logger.info("═" * 46)
    logger.info(" QA TOTAL: %d passed, %d failed, %d warnings → %s",
                combined.passed, combined.failed, combined.warnings, combined.status.name)
    return 0 if combined.status is not ValidationStatus.FAIL else 1


def run_build(args: argparse.Namespace) -> int:
    """Build, render, QA and manifest one artifact set.

    Args:
        args: Parsed CLI arguments.

    Returns:
        Process exit code: ``0`` when golden + structural QA pass.
    """
    config = AppConfig.load()
    config.validate()
    if args.demo:
        logger.info("--demo accepted: every build displays engine-computed demo results")

    # Tier 1 — golden suite runs in-process and is embedded in the workbook.
    golden = run_golden_suite()
    _log_qa_report(golden)

    # Domain engine computes; renderer displays (never the reverse).
    results = run_demo_cases()
    model = assemble_workbook_model(results, config, {"golden": golden})
    workbook = ExcelRenderer().render(model)

    # Protection with an explicit, logged posture.
    protect = not args.no_protect
    password = config.protection_password(env=_env(args.password_env)) if protect else None
    protection_enabled = ProtectionManager.apply(workbook, password) if protect else False
    if protect and not protection_enabled:
        logger.warning("--no-protect was not set but no password was found; "
                       "set %s to enable protection", args.password_env)

    # Save artifact + checksum.
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = ARTIFACT_PATTERN.format(version=__version__)
    xlsx_path = output_dir / filename
    if xlsx_path.exists():
        try:
            xlsx_path.unlink()
        except PermissionError:
            logger.error("previous artifact is open; close it and retry: %s", xlsx_path)
            return 1
    workbook.save(str(xlsx_path))
    sha = compute_sha256(xlsx_path)
    (output_dir / (filename + ".sha256")).write_text(f"{sha}  {filename}\n", encoding="utf-8")
    logger.info("Saved: %s", xlsx_path)
    logger.info("SHA-256: %s", sha)

    # Manifest pass 1 (golden QA) so structural checks have something to verify…
    manifest_path = output_dir / (xlsx_path.stem + ".json")
    manifest = build_manifest(
        filename=filename, sha256=sha, sheets=model.sheet_titles,
        qa_reports={"golden": golden}, protection_enabled=protection_enabled,
    )
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    # Tier 3 — structural QA over the artifact set.
    structural = run_structural_checks(xlsx_path, manifest_path, expected_sheets=model.sheet_titles)
    _log_qa_report(structural)

    # Manifest pass 2 — final, including structural results.
    manifest = build_manifest(
        filename=filename, sha256=sha, sheets=model.sheet_titles,
        qa_reports={"golden": golden, "structural": structural},
        protection_enabled=protection_enabled,
    )
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info("Manifest: %s", manifest_path)

    combined = golden.merged_with(structural)
    logger.info("═" * 46)
    logger.info(" BUILD COMPLETE — QA: %d passed, %d failed, %d warnings → %s",
                combined.passed, combined.failed, combined.warnings, combined.status.name)
    return 0 if combined.status is not ValidationStatus.FAIL else 1


def _env(password_env_name: str) -> Dict[str, str]:
    """Environment mapping for password lookup (isolated for tests)."""
    import os

    return {password_env_name: os.environ.get(password_env_name, "")}


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point.

    Args:
        argv: Argument vector (defaults to ``sys.argv[1:]``).

    Returns:
        Process exit code.
    """
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = build_arg_parser().parse_args(argv)

    if args.validate is not None:
        return run_validation(paths=args.validate)
    return run_build(args)


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
