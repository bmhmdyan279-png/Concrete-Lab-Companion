"""In-process unit tests for the CLI.

``tests/test_cli.py`` drives ``build.py`` as a real subprocess, which is the
right end-to-end check but leaves the CLI's own branches unmeasured (the
child interpreter's coverage is not the parent's).  These tests call the
entry points directly so every decision in ``cli.py`` is actually executed:
argument parsing, data-source resolution, protection posture, error exits
and the two QA modes.
"""

import json
import logging
from pathlib import Path
from typing import List

import pytest

from concrete_lab import __version__
from concrete_lab.cli import (
    ARTIFACT_PATTERN,
    _env,
    _resolve_data,
    build_arg_parser,
    main,
    run_build,
    run_validation,
)
from concrete_lab.report.manifest import DATA_SOURCE_DEMO, DATA_SOURCE_USER

ARTIFACT = ARTIFACT_PATTERN.format(version=__version__)


def build_args(tmp_path: Path, *extra: str) -> List[str]:
    """A build invocation that writes into ``tmp_path`` and never protects."""
    return ["--output", str(tmp_path), "--no-protect", *extra]


def write_case_file(tmp_path: Path, cases: List[dict], name: str = "cases.json") -> Path:
    """Write a user case file and return its path."""
    path = tmp_path / name
    path.write_text(json.dumps({"cases": cases}, ensure_ascii=False), encoding="utf-8")
    return path


# ─── Argument surface ─────────────────────────────────────────────────────


class TestArgumentParser:
    def test_defaults(self) -> None:
        args = build_arg_parser().parse_args([])
        assert args.output == "output"
        assert args.no_protect is False
        assert args.demo is False
        assert args.input is None
        assert args.config is None
        assert args.validate is None
        assert args.password_env == "WORKBOOK_PASSWORD"

    def test_validate_accepts_zero_or_more_paths(self) -> None:
        assert build_arg_parser().parse_args(["--validate"]).validate == []
        assert build_arg_parser().parse_args(["--validate", "a.py", "b.py"]).validate == ["a.py", "b.py"]

    def test_demo_and_input_cannot_be_combined(self) -> None:
        """The point of keeping ``--demo`` is that this combination is an error."""
        with pytest.raises(SystemExit) as info:
            build_arg_parser().parse_args(["--demo", "--input", "x.json"])
        assert info.value.code == 2

    def test_version_flag_exits_with_the_package_version(self, capsys) -> None:
        with pytest.raises(SystemExit) as info:
            build_arg_parser().parse_args(["--version"])
        assert info.value.code == 0
        assert __version__ in capsys.readouterr().out

    def test_help_lists_the_real_flags(self, capsys) -> None:
        with pytest.raises(SystemExit):
            build_arg_parser().parse_args(["--help"])
        out = capsys.readouterr().out
        for flag in ("--output", "--no-protect", "--password-env", "--config",
                     "--input", "--demo", "--validate"):
            assert flag in out
        assert "examples" in out  # the epilog with usage examples


# ─── Data-source resolution ───────────────────────────────────────────────


class TestDataSourceResolution:
    def test_demo_is_the_default(self, tmp_path: Path) -> None:
        args = build_arg_parser().parse_args([])
        cases, source, path = _resolve_data(args)
        assert source == DATA_SOURCE_DEMO
        assert path is None
        assert len(cases) == 7

    def test_demo_flag_selects_the_same_dataset(self) -> None:
        args = build_arg_parser().parse_args(["--demo"])
        cases, source, _ = _resolve_data(args)
        assert source == DATA_SOURCE_DEMO and len(cases) == 7

    def test_input_file_switches_the_source(self, tmp_path: Path) -> None:
        case_file = write_case_file(tmp_path, [
            {"test_id": "1-2", "inputs": {"wet_mass_g": 1050, "dry_mass_g": 1000}},
        ])
        args = build_arg_parser().parse_args(["--input", str(case_file)])
        cases, source, name = _resolve_data(args)
        assert source == DATA_SOURCE_USER
        assert name == "cases.json"
        assert cases[0][0] == "1-2"


class TestEnvHelper:
    def test_returns_only_the_requested_variable(self, monkeypatch) -> None:
        monkeypatch.setenv("WORKBOOK_PASSWORD", "secret")
        assert _env("WORKBOOK_PASSWORD") == {"WORKBOOK_PASSWORD": "secret"}

    def test_missing_variable_is_an_empty_string(self, monkeypatch) -> None:
        monkeypatch.delenv("WORKBOOK_PASSWORD", raising=False)
        assert _env("WORKBOOK_PASSWORD") == {"WORKBOOK_PASSWORD": ""}

    def test_custom_variable_name_is_honoured(self, monkeypatch) -> None:
        monkeypatch.setenv("MY_PW", "x")
        assert _env("MY_PW") == {"MY_PW": "x"}


# ─── Build mode ───────────────────────────────────────────────────────────


class TestRunBuild:
    def test_demo_build_produces_all_three_artifacts(self, tmp_path: Path) -> None:
        assert main(build_args(tmp_path)) == 0
        assert (tmp_path / ARTIFACT).is_file()
        assert (tmp_path / (ARTIFACT + ".sha256")).is_file()
        manifest = json.loads((tmp_path / (Path(ARTIFACT).stem + ".json")).read_text(encoding="utf-8"))
        assert manifest["data_source"] == DATA_SOURCE_DEMO
        assert manifest["data_source_path"] is None
        assert manifest["qa"]["failed"] == 0
        assert set(manifest["qa"]["tiers"]) == {"golden", "structural"}

    def test_input_build_is_traceable_as_user_data(self, tmp_path: Path) -> None:
        case_file = write_case_file(tmp_path, [
            {"test_id": "3-1", "inputs": {"measured_height_mm": 210}},
        ])
        assert main(build_args(tmp_path, "--input", str(case_file))) == 0
        manifest = json.loads(
            (tmp_path / (Path(ARTIFACT).stem + ".json")).read_text(encoding="utf-8"))
        assert manifest["data_source"] == DATA_SOURCE_USER
        assert manifest["data_source_path"] == "cases.json"
        assert manifest["sheets"]  # non-empty

    def test_rejected_input_exits_non_zero_without_building(self, tmp_path: Path) -> None:
        case_file = write_case_file(tmp_path, [
            {"test_id": "3-1", "inputs": {"measured_height_mm": 900}},  # above the cone
        ])
        assert main(build_args(tmp_path, "--input", str(case_file))) == 1
        assert not (tmp_path / ARTIFACT).exists(), "nothing may be built from rejected data"

    def test_unusable_case_file_exits_non_zero(self, tmp_path: Path, caplog) -> None:
        broken = tmp_path / "broken.json"
        broken.write_text("{oops", encoding="utf-8")
        with caplog.at_level(logging.ERROR):
            assert main(build_args(tmp_path, "--input", str(broken))) == 1
        assert "not valid JSON" in caplog.text

    def test_missing_case_file_exits_non_zero(self, tmp_path: Path) -> None:
        assert main(build_args(tmp_path, "--input", str(tmp_path / "absent.json"))) == 1

    def test_custom_config_is_used_and_recorded(self, tmp_path: Path) -> None:
        custom = tmp_path / "custom.yaml"
        custom.write_text(
            f"project:\n  name: 'Custom Lab'\n  version: '{__version__}'\n", encoding="utf-8")
        assert main(build_args(tmp_path, "--config", str(custom))) == 0
        manifest = json.loads(
            (tmp_path / (Path(ARTIFACT).stem + ".json")).read_text(encoding="utf-8"))
        assert manifest["config_source"] == str(custom)

    def test_stale_config_version_is_refused(self, tmp_path: Path, caplog) -> None:
        """A config that disagrees with the package must stop the build politely."""
        stale = tmp_path / "stale.yaml"
        stale.write_text("project:\n  name: 'X'\n  version: '0.0.1'\n", encoding="utf-8")
        with caplog.at_level(logging.ERROR):
            assert main(build_args(tmp_path, "--config", str(stale))) == 1
        assert "disagrees with the package version" in caplog.text
        assert not (tmp_path / ARTIFACT).exists()

    def test_missing_config_file_is_refused(self, tmp_path: Path, caplog) -> None:
        with caplog.at_level(logging.ERROR):
            assert main(build_args(tmp_path, "--config", str(tmp_path / "absent.yaml"))) == 1
        assert "configuration rejected" in caplog.text

    def test_protection_is_enabled_from_the_environment(self, tmp_path: Path, monkeypatch) -> None:
        import openpyxl

        monkeypatch.setenv("WORKBOOK_PASSWORD", "s3cret")
        assert main(["--output", str(tmp_path)]) == 0  # note: protection NOT disabled
        manifest = json.loads(
            (tmp_path / (Path(ARTIFACT).stem + ".json")).read_text(encoding="utf-8"))
        assert manifest["protection"]["enabled"] is True

        workbook = openpyxl.load_workbook(str(tmp_path / ARTIFACT))
        assert workbook["00_راهنما"].protection.sheet is True
        assert workbook["_Standards"].protection.sheet is not True  # internal sheets stay open

    def test_protection_without_a_password_warns_and_continues(self, tmp_path: Path,
                                                               monkeypatch, caplog) -> None:
        monkeypatch.delenv("WORKBOOK_PASSWORD", raising=False)
        custom = tmp_path / "nopassword.yaml"
        custom.write_text(
            f"project:\n  name: 'X'\n  version: '{__version__}'\n", encoding="utf-8")
        with caplog.at_level(logging.WARNING):
            assert main(["--output", str(tmp_path), "--config", str(custom)]) == 0
        assert "no password was found" in caplog.text
        manifest = json.loads(
            (tmp_path / (Path(ARTIFACT).stem + ".json")).read_text(encoding="utf-8"))
        assert manifest["protection"]["enabled"] is False

    def test_locked_artifact_reports_a_usable_message(self, tmp_path: Path, monkeypatch,
                                                      caplog) -> None:
        """Excel holding the previous artifact open must not surface as a traceback."""
        assert main(build_args(tmp_path)) == 0

        real_unlink = Path.unlink

        def refuse_artifact(self, *args, **kwargs):
            if self.name == ARTIFACT:
                raise PermissionError(13, "Permission denied")
            return real_unlink(self, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", refuse_artifact)
        with caplog.at_level(logging.ERROR):
            assert main(build_args(tmp_path)) == 1
        assert "previous artifact is open" in caplog.text

    def test_output_directory_is_created(self, tmp_path: Path) -> None:
        nested = tmp_path / "a" / "b" / "c"
        assert main(build_args(nested)) == 0
        assert (nested / ARTIFACT).is_file()


# ─── Validate mode ────────────────────────────────────────────────────────


class TestRunValidation:
    def test_passing_selection_exits_zero(self) -> None:
        assert main(["--validate", "tests/test_domain.py"]) == 0

    def test_failing_selection_exits_one(self, tmp_path: Path) -> None:
        bad = tmp_path / "test_cli_unit_forced_failure.py"
        bad.write_text("def test_broken():\n    assert False\n", encoding="utf-8")
        assert main(["--validate", str(bad)]) == 1

    def test_run_validation_reports_both_tiers(self, caplog) -> None:
        with caplog.at_level(logging.INFO):
            code = run_validation(paths=["tests/test_domain.py"])
        assert code == 0
        assert "VALIDATION MODE" in caplog.text
        assert "QA TOTAL" in caplog.text

    def test_golden_warnings_are_logged_as_notes(self, caplog) -> None:
        """A warning without an explanation is not actionable."""
        with caplog.at_level(logging.INFO):
            run_validation(paths=["tests/test_domain.py"])
        assert "no engine implementation yet" in caplog.text


# ─── Entry point plumbing ─────────────────────────────────────────────────


class TestEntryPoint:
    def test_run_build_returns_an_int(self, tmp_path: Path) -> None:
        args = build_arg_parser().parse_args(build_args(tmp_path))
        assert isinstance(run_build(args), int)

    def test_version_is_consistent_with_the_artifact_name(self) -> None:
        assert ARTIFACT.startswith("Concrete_Lab_Companion_v")
        assert ARTIFACT.endswith(f"{__version__}.xlsx")
