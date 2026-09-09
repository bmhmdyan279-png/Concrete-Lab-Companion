"""Consistency tests: one fact, one place, verified everywhere.

Every critic who looked at this repository found the same class of defect,
and it was never a wrong formula — it was a *claim that disagreed with the
code*:

* ``config.yaml`` said ``implemented: 20`` while the engine implemented 7;
* ``config.yaml`` listed 31 worksheets from the retired v3 workbook while
  the v4 renderer emits 14;
* the landing page printed a SHA-256 digest that no rebuild could match;
* ``CITATION.cff`` shipped ``version: "soon"`` and ``doi: "soon"``;
* the README quoted a hand-maintained test count.

Hand-maintained copies of derived facts drift; this module makes drifting
fail the build instead.  Every number that describes the product is read
from the live registries and compared against each place that repeats it.
"""

import ast
import json
import re
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pytest
import yaml

from concrete_lab import __version__
from concrete_lab.cli import ARTIFACT_PATTERN
from concrete_lab.config import AppConfig
from concrete_lab.domain.engine import CALCULATORS
from concrete_lab.render.excel.assembler import assemble_workbook_model
from concrete_lab.specs.base import TEST_REGISTRY, all_tests, implemented_tests

REPO_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = REPO_ROOT / "src" / "concrete_lab"

CONFIG_FILE = REPO_ROOT / "config.yaml"
PYPROJECT_FILE = REPO_ROOT / "pyproject.toml"
REQUIREMENTS_FILE = REPO_ROOT / "requirements.txt"
REQUIREMENTS_DEV_FILE = REPO_ROOT / "requirements-dev.txt"
CITATION_FILE = REPO_ROOT / "CITATION.cff"
README_FILE = REPO_ROOT / "README.md"
CHANGELOG_FILE = REPO_ROOT / "CHANGELOG.md"
LANDING_HTML = REPO_ROOT / "landing" / "index.html"
LANDING_RELEASE = REPO_ROOT / "landing" / "release.json"
BUILD_WORKFLOW = REPO_ROOT / ".github" / "workflows" / "build.yml"
ERRATA_FILE = REPO_ROOT / "validation" / "errata.yaml"

#: Strings that mean "someone intended to fill this in later".
PLACEHOLDERS: Tuple[str, ...] = ("soon", "TODO", "TBD", "FIXME", "xxx", "placeholder")

#: Import name → distribution name for the runtime requirements.
RUNTIME_IMPORTS: Dict[str, str] = {"openpyxl": "openpyxl", "yaml": "pyyaml"}

#: SHA-256 hex digest, as printed on the landing page.
SHA256_RE = re.compile(r"\b[0-9a-f]{64}\b")

#: Persian digits, for claims the README states in Persian.
PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"


def to_persian_digits(number: int) -> str:
    """Render an integer with Persian digits (``7`` → ``۷``)."""
    return "".join(PERSIAN_DIGITS[int(digit)] for digit in str(number))


def python_files(root: Path = SRC_ROOT) -> Iterable[Path]:
    """Every ``.py`` file under ``root``."""
    return sorted(path for path in root.rglob("*.py"))


@pytest.fixture(scope="module")
def config_yaml() -> dict:
    """The raw repository ``config.yaml``."""
    return yaml.safe_load(CONFIG_FILE.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def pyproject() -> dict:
    """The parsed ``pyproject.toml`` (one read per session)."""
    try:
        import tomllib
    except ModuleNotFoundError:  # Python 3.9 / 3.10
        import tomli as tomllib  # type: ignore[no-redef]
    with open(PYPROJECT_FILE, "rb") as handle:
        return tomllib.load(handle)


@pytest.fixture(scope="module")
def landing_html() -> str:
    """The landing page source."""
    return LANDING_HTML.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def citation() -> dict:
    """The parsed ``CITATION.cff``."""
    return yaml.safe_load(CITATION_FILE.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def readme() -> str:
    """The README source."""
    return README_FILE.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def build_workflow() -> str:
    """The build workflow source."""
    return BUILD_WORKFLOW.read_text(encoding="utf-8")


@pytest.fixture(scope="module")
def rendered_sheets() -> Tuple[str, ...]:
    """Worksheet titles the display-only renderer actually emits."""
    from concrete_lab.domain.engine import run_demo_cases

    model = assemble_workbook_model(run_demo_cases(), AppConfig.load(), {})
    return model.sheet_titles


# ─── Version: one value, five places ──────────────────────────────────────


class TestVersionAgreement:
    def test_package_version_is_semantic(self) -> None:
        assert re.fullmatch(r"\d+\.\d+\.\d+", __version__), __version__

    def test_pyproject_matches_the_package(self) -> None:
        match = re.search(r'^version\s*=\s*"([^"]+)"',
                          PYPROJECT_FILE.read_text(encoding="utf-8"), re.MULTILINE)
        assert match, "pyproject.toml declares no [project].version"
        assert match.group(1) == __version__

    def test_config_yaml_matches_the_package(self, config_yaml: dict) -> None:
        assert config_yaml["project"]["version"] == __version__

    def test_appconfig_rejects_a_stale_version(self, tmp_path: Path) -> None:
        """``AppConfig.validate()`` must catch the drift, not just the tests."""
        stale = tmp_path / "config.yaml"
        stale.write_text(
            "project:\n  name: 'X'\n  version: '0.0.1'\n", encoding="utf-8"
        )
        with pytest.raises(Exception, match="disagrees with the package version"):
            AppConfig.load(stale).validate()

    def test_citation_matches_the_package(self) -> None:
        citation = yaml.safe_load(CITATION_FILE.read_text(encoding="utf-8"))
        assert str(citation["version"]) == __version__

    def test_landing_release_matches_the_package(self) -> None:
        release = json.loads(LANDING_RELEASE.read_text(encoding="utf-8"))
        assert release["version"] == __version__

    def test_changelog_documents_the_current_version(self) -> None:
        assert f"[{__version__}]" in CHANGELOG_FILE.read_text(encoding="utf-8"), (
            f"CHANGELOG.md has no [{__version__}] entry"
        )

    def test_no_artifact_name_pins_an_old_version(self) -> None:
        """A bumped version must not leave a stale artifact name behind.

        Only *artifact names* are checked, not prose: a docstring or a
        changelog that says "this was broken in an earlier release" is
        correct history.  A file that names an artifact carrying another
        version is a stale literal waiting to mislead someone.
        """
        pattern = re.compile(r"Concrete_Lab_Companion_v(\d+\.\d+\.\d+)")
        offenders: List[str] = []
        for path in REPO_ROOT.rglob("*"):
            if not path.is_file() or ".git" in path.parts or "output" in path.parts:
                continue
            if path.suffix not in {".py", ".yaml", ".yml", ".toml", ".json", ".html", ".ini", ".cfg"}:
                continue
            text = path.read_text(encoding="utf-8", errors="ignore")
            for found in set(pattern.findall(text)):
                if found != __version__:
                    offenders.append(f"{path.relative_to(REPO_ROOT)} → v{found}")
        assert not offenders, f"artifact names pinned to another version: {offenders}"

    def test_the_version_literal_lives_in_exactly_one_module(self) -> None:
        """``__init__.py`` owns the version; everything else must import it."""
        owners = [
            str(path.relative_to(REPO_ROOT))
            for path in python_files()
            if re.search(r"^__version__\s*=", path.read_text(encoding="utf-8"), re.MULTILINE)
        ]
        assert owners == ["src/concrete_lab/__init__.py"], owners


# ─── config.yaml describes the product that exists ────────────────────────


class TestConfigDescribesReality:
    def test_implemented_count_matches_the_engine(self, config_yaml: dict) -> None:
        sheets = config_yaml["sheets"]
        assert sheets["total_expected"] == len(all_tests())
        assert sheets["implemented"] == len(implemented_tests())
        assert sheets["implemented"] == len(CALCULATORS)
        assert sheets["pending"] == len(all_tests()) - len(implemented_tests())

    def test_key_sheets_match_the_rendered_workbook(
        self, config_yaml: dict, rendered_sheets: Tuple[str, ...]
    ) -> None:
        declared = list(config_yaml["sheets"]["key_sheets"])
        assert declared == list(rendered_sheets), (
            f"config.yaml key_sheets does not match the rendered workbook:\n"
            f"  declared: {declared}\n  rendered: {list(rendered_sheets)}"
        )

    def test_errata_counts_match_errata_yaml(self, config_yaml: dict) -> None:
        errata = yaml.safe_load(ERRATA_FILE.read_text(encoding="utf-8"))["errata"]
        statuses = [str(item.get("status")) for item in errata]
        declared = config_yaml["errata"]
        assert declared["total"] == len(errata)
        assert declared["confirmed"] == statuses.count("confirmed")
        assert declared["pending"] == statuses.count("pending")

    def test_declared_validation_paths_exist(self, config_yaml: dict) -> None:
        validation = config_yaml["validation"]
        assert (REPO_ROOT / validation["errata_file"]).is_file()
        assert (REPO_ROOT / validation["golden_cases_dir"]).is_dir()


# ─── The registries agree with each other ─────────────────────────────────


class TestRegistryAgreement:
    def test_calculators_are_exactly_the_implemented_specs(self) -> None:
        """``implemented=True`` and a calculator must never disagree."""
        spec_ids = {spec.id for spec in implemented_tests()}
        assert set(CALCULATORS) == spec_ids, (
            f"specs marked implemented: {sorted(spec_ids)}, "
            f"engine calculators: {sorted(CALCULATORS)}"
        )

    def test_every_spec_references_a_registered_standard(self) -> None:
        from concrete_lab.standards import registry as standards_registry

        codes = {spec.code for spec in standards_registry.all_standards()}
        unknown = sorted(
            f"{test.id}→{test.standard_code}"
            for test in all_tests() if test.standard_code not in codes
        )
        assert not unknown, f"specs reference unregistered standards: {unknown}"

    def test_sheet_names_are_unique_and_within_excel_limit(self) -> None:
        names = [spec.sheet_name for spec in all_tests()]
        assert len(names) == len(set(names)), "duplicate worksheet names in the catalogue"
        too_long = [name for name in names if len(name) > 31]
        assert not too_long, f"worksheet names exceed Excel's 31 characters: {too_long}"

    def test_catalogue_is_complete(self) -> None:
        assert len(TEST_REGISTRY) == 20
        assert len(all_tests()) == len(TEST_REGISTRY)


# ─── Dependencies: declared == used ───────────────────────────────────────


def _parse_requirements(path: Path) -> List[str]:
    """Distribution names from a requirements file (comments stripped)."""
    names = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line or line.startswith("-"):
            continue
        names.append(re.split(r"[<>=!~\[]", line, maxsplit=1)[0].strip().lower())
    return names


def _imported_top_level_modules() -> set:
    """Third-party modules actually imported by the package."""
    found = set()
    for path in python_files():
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                found.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                found.add(node.module.split(".")[0])
    return found


class TestDependencies:
    def test_runtime_requirements_are_all_imported(self) -> None:
        imported = _imported_top_level_modules()
        declared = _parse_requirements(REQUIREMENTS_FILE)
        unused = [name for name in declared
                  if name not in {RUNTIME_IMPORTS.get(mod, mod) for mod in imported}]
        assert not unused, (
            f"requirements.txt declares dependencies nothing imports: {unused} "
            f"(xlsxwriter was such a dead dependency until 4.1.0)"
        )

    def test_no_test_tooling_in_runtime_requirements(self) -> None:
        declared = set(_parse_requirements(REQUIREMENTS_FILE))
        leaked = declared & {"pytest", "pytest-cov", "coverage", "ruff", "pre-commit"}
        assert not leaked, f"test-only tools must live in requirements-dev.txt: {sorted(leaked)}"

    def test_dev_requirements_extend_the_runtime_ones(self) -> None:
        text = REQUIREMENTS_DEV_FILE.read_text(encoding="utf-8")
        assert "-r requirements.txt" in text, "requirements-dev.txt must include the runtime set"
        assert "pytest" in _parse_requirements(REQUIREMENTS_DEV_FILE)

    def test_requirements_files_agree_with_pyproject(self) -> None:
        pyproject = PYPROJECT_FILE.read_text(encoding="utf-8")
        for name in _parse_requirements(REQUIREMENTS_FILE):
            assert name in pyproject.lower(), (
                f"{name} is in requirements.txt but not in pyproject.toml dependencies"
            )


# ─── Packaging actually works ─────────────────────────────────────────────


class TestPackaging:
    def test_build_system_is_declared(self, pyproject: dict) -> None:
        backend = pyproject.get("build-system", {})
        assert backend.get("build-backend"), "no [build-system].build-backend — pip install . cannot work"
        assert backend.get("requires"), "no [build-system].requires"

    def test_runtime_dependencies_are_declared(self, pyproject: dict) -> None:
        dependencies = {re.split(r"[<>=!~\[]", dep, maxsplit=1)[0].strip().lower()
                        for dep in pyproject["project"].get("dependencies", [])}
        assert set(_parse_requirements(REQUIREMENTS_FILE)) <= dependencies

    def test_dev_extra_is_declared(self, pyproject: dict) -> None:
        optional = pyproject["project"].get("optional-dependencies", {})
        assert "dev" in optional and any("pytest" in dep for dep in optional["dev"])

    def test_console_script_points_at_the_cli(self, pyproject: dict) -> None:
        scripts = pyproject["project"].get("scripts", {})
        assert scripts.get("concrete-lab-companion") == "concrete_lab.cli:main"

    def test_src_layout_is_discoverable(self, pyproject: dict) -> None:
        find = pyproject.get("tool", {}).get("setuptools", {}).get("packages", {}).get("find", {})
        assert find.get("where") == ["src"], "src layout must be declared for pip install"

    def test_pytest_has_exactly_one_configuration_source(self) -> None:
        """Two pytest configs produced a real warning and two truths."""
        assert not (REPO_ROOT / "pytest.ini").exists(), (
            "pytest.ini must stay deleted: pytest config lives in pyproject.toml"
        )
        ini_options = PYPROJECT_FILE.read_text(encoding="utf-8")
        assert "[tool.pytest.ini_options]" in ini_options

    def test_license_file_and_declaration_agree(self, pyproject: dict) -> None:
        declared = json.dumps(pyproject["project"].get("license", ""), ensure_ascii=False)
        license_text = (REPO_ROOT / "LICENSE").read_text(encoding="utf-8")
        assert "CC BY-NC" in declared or "cc-by-nc" in declared.lower()
        assert "Attribution-NonCommercial" in license_text or "CC BY-NC" in license_text


# ─── Landing page: nothing hard-coded ─────────────────────────────────────


class TestLandingPage:
    def test_no_hardcoded_sha256(self, landing_html: str) -> None:
        digests = SHA256_RE.findall(landing_html)
        assert not digests, (
            f"landing page hard-codes SHA-256 digest(s) {digests}; a rebuilt artifact "
            f"would then look tampered with. Read release.json instead."
        )

    def test_no_hardcoded_current_version(self, landing_html: str) -> None:
        """The page must never pin the release it is serving."""
        assert __version__ not in landing_html, "the page must take its version from release.json"
        persian = ".".join(to_persian_digits(int(part)) for part in __version__.split("."))
        assert persian not in landing_html, f"the page pins the Persian version string {persian}"

    def test_page_loads_the_generated_release_info(self, landing_html: str) -> None:
        assert "release.json" in landing_html
        assert "loadReleaseInfo" in landing_html
        assert "hashValue" in landing_html

    def test_page_states_the_sample_report_limitation_up_front(self, landing_html: str) -> None:
        """The caveat belongs in the hero, not buried in step 2 of a guide."""
        hero = landing_html.split('id="features"', 1)[0]
        assert "گزارش نمونه" in hero and "ماشین‌حساب" in hero
        assert "--input" in hero

    def test_page_documents_the_user_input_path(self, landing_html: str) -> None:
        assert "python build.py --input your_data.json" in landing_html

    def test_release_json_is_well_formed(self) -> None:
        release = json.loads(LANDING_RELEASE.read_text(encoding="utf-8"))
        assert SHA256_RE.fullmatch(release["sha256"]), "sha256 must be a lowercase hex digest"
        assert release["filename"] == ARTIFACT_PATTERN.format(version=__version__)
        assert release["data_source"] in {"demo", "user-input"}
        assert release["qa_status"] in {"pass", "warn", "fail"}
        assert release["tests"]["implemented"] == len(implemented_tests())
        assert release["generator"] == "scripts/make_release_manifest.py"

# ─── Citation & documentation hygiene ─────────────────────────────────────


class TestCitation:
    def test_no_placeholder_values(self, citation: dict) -> None:
        offenders = []

        def walk(node: object, trail: str) -> None:
            if isinstance(node, dict):
                for key, value in node.items():
                    walk(value, f"{trail}.{key}")
            elif isinstance(node, list):
                for index, value in enumerate(node):
                    walk(value, f"{trail}[{index}]")
            elif isinstance(node, str):
                low = node.strip().lower()
                if low in PLACEHOLDERS or any(low == placeholder for placeholder in PLACEHOLDERS):
                    offenders.append(f"{trail}={node!r}")

        walk(citation, "CITATION.cff")
        assert not offenders, f"CITATION.cff still contains placeholders: {offenders}"

    def test_release_date_is_a_real_date(self, citation: dict) -> None:
        from datetime import date

        released = citation["date-released"]
        assert isinstance(released, date), f"date-released must be a date, got {released!r}"

    def test_author_identity_is_complete(self, citation: dict) -> None:
        for author in citation["authors"]:
            assert author.get("family-names") and author.get("given-names")
            assert "orcid" not in author or str(author["orcid"]).startswith("https://orcid.org/")

    def test_license_matches_the_license_file(self, citation: dict) -> None:
        assert citation["license"] == "CC-BY-NC-4.0"
        assert "CC BY-NC 4.0" in README_FILE.read_text(encoding="utf-8")


class TestReadmeClaims:
    def test_no_hand_maintained_test_count(self, readme: str) -> None:
        """``# ۲۱۵ تست`` went stale the moment a test was added or removed."""
        assert not re.search(r"\b\d{2,4}\s+تست\b", readme), "README quotes a hard-coded test count"
        assert not re.search(r"\b[۰-۹]{2,4}\s+تست\b", readme), "README quotes a hard-coded test count"

    def test_implemented_count_is_stated_correctly(self, readme: str) -> None:
        count = to_persian_digits(len(implemented_tests()))
        total = to_persian_digits(len(all_tests()))
        assert f"{count} آزمایش" in readme, f"README must state {count} implemented tests"
        assert f"{total} آزمایش" in readme, f"README must state {total} catalogue tests"

    def test_every_claimed_layer_exists(self, readme: str) -> None:
        for relative in ("src/concrete_lab/domain/", "src/concrete_lab/standards/",
                         "src/concrete_lab/specs/", "src/concrete_lab/qa/",
                         "src/concrete_lab/render/excel/", "src/concrete_lab/report/"):
            assert relative in readme
            assert (REPO_ROOT / relative).is_dir(), f"README documents a missing directory: {relative}"

    def test_documented_commands_are_real(self, readme: str) -> None:
        from concrete_lab.cli import build_arg_parser

        parser_text = build_arg_parser().format_help()
        for flag in re.findall(r"python3? build\.py (--[\w-]+)", readme):
            assert flag in parser_text, f"README documents an unknown flag: {flag}"

    def test_every_version_mentioned_is_the_current_one(self, readme: str) -> None:
        """A README that names a version must name the one being shipped."""
        persian = ".".join(to_persian_digits(int(part)) for part in __version__.split("."))
        # ``\d`` is Unicode-aware and would also match Persian digits, so both
        # scripts are matched with explicit character classes.
        latin = re.findall(r"\b[0-9]+\.[0-9]+\.[0-9]+\b", readme)
        persian_found = re.findall(r"[۰-۹]+\.[۰-۹]+\.[۰-۹]+", readme)
        assert all(found == __version__ for found in latin), (
            f"README mentions other versions: {sorted(set(latin))}"
        )
        assert all(found == persian for found in persian_found), (
            f"README mentions other Persian versions: {sorted(set(persian_found))}"
        )

    def test_no_dead_code_claim_is_false(self, readme: str) -> None:
        """README promises "no dead code"; these were the counter-examples."""
        assert not (REPO_ROOT / "tests" / "test_golden.py").exists()
        assert not (SRC_ROOT / "domain" / "results.py").exists()
        assert "بدون کد مرده" in readme


# ─── Architecture contract ────────────────────────────────────────────────


class TestArchitectureContract:
    def test_no_source_file_writes_excel_formulas(self) -> None:
        """The v4 headline claim, checked statically over every module."""
        pattern = re.compile(r"value\s*=\s*f?[\"']=")
        offenders = [
            f"{path.relative_to(REPO_ROOT)}:{number}"
            for path in python_files(SRC_ROOT)
            for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)
            if pattern.search(line)
        ]
        assert not offenders, f"formula-writing code found: {offenders}"

    #: Callables that *perform* science; a renderer importing these has broken
    #: the separation contract, even if it only imported them "for typing".
    FORBIDDEN_IN_RENDER = frozenset(
        {"calculate", "run_batch", "run_cases", "run_demo_cases", "CALCULATORS",
         "ld_correction_factor", "mean_rebound_number", "estimate_strength",
         "fineness_modulus", "curve_from_retained"}
    )

    def test_renderer_never_imports_the_science(self) -> None:
        """Display may know *about* results; it may never compute one."""
        for path in python_files(SRC_ROOT / "render"):
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if not isinstance(node, ast.ImportFrom) or not node.module:
                    continue
                imported = {alias.name for alias in node.names}
                leaked = imported & self.FORBIDDEN_IN_RENDER
                assert not leaked, f"{path.name} imports computation entry points: {sorted(leaked)}"
                assert not node.module.startswith(("concrete_lab.standards.astm",
                                                   "concrete_lab.standards.isiri")), (
                    f"{path.name} imports a scientific ruleset module ({node.module}); "
                    f"renderers display, they must not compute"
                )

    #: Development-only distributions.  Importing one at module level breaks a
    #: plain ``pip install`` of the package (found in 4.1.0: ``qa/engine.py``
    #: imported pytest, so the installed console script could not even start).
    DEV_ONLY_DISTRIBUTIONS = frozenset({"pytest", "coverage", "ruff", "_pytest"})

    def test_no_dev_dependency_is_imported_at_module_level(self) -> None:
        offenders = []
        for path in python_files():
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for node in tree.body:  # module level only — nested imports are fine
                modules = []
                if isinstance(node, ast.Import):
                    modules = [alias.name.split(".")[0] for alias in node.names]
                elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                    modules = [node.module.split(".")[0]]
                offenders += [
                    f"{path.relative_to(REPO_ROOT)}:{node.lineno} → {name}"
                    for name in modules if name in self.DEV_ONLY_DISTRIBUTIONS
                ]
        assert not offenders, (
            f"runtime code imports dev-only packages at module level: {offenders}"
        )

    def test_pytest_tier_reports_a_missing_runner_actionably(self, monkeypatch) -> None:
        """``--validate`` without pytest must explain itself, not traceback."""
        import builtins

        from concrete_lab.qa.engine import QAEngine, QAError

        real_import = builtins.__import__

        def deny(name, *args, **kwargs):
            if name == "pytest":
                raise ModuleNotFoundError("No module named 'pytest'")
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", deny)
        with pytest.raises(QAError, match=r"requirements-dev\.txt"):
            QAEngine().run_pytest([])

    def test_every_module_has_a_docstring(self) -> None:
        """README claims full docstrings; enforce it instead of hoping."""
        missing = []
        for path in python_files():
            tree = ast.parse(path.read_text(encoding="utf-8"))
            if not ast.get_docstring(tree):
                missing.append(str(path.relative_to(REPO_ROOT)))
        assert not missing, f"modules without a docstring: {missing}"

    def test_password_env_var_is_declared_once(self) -> None:
        from concrete_lab import config, constants

        assert config.PASSWORD_ENV_VAR is constants.PASSWORD_ENV_VAR or (
            config.PASSWORD_ENV_VAR == constants.PASSWORD_ENV_VAR
        )
        declarations = [
            f"{path.relative_to(REPO_ROOT)}:{number}"
            for path in python_files()
            for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1)
            if re.search(r'PASSWORD_ENV_VAR[^=]*=\s*["\']WORKBOOK_PASSWORD["\']', line)
        ]
        assert len(declarations) == 1, f"WORKBOOK_PASSWORD declared in: {declarations}"


# ─── CI must not pin what the code derives ────────────────────────────────


class TestContinuousIntegration:
    def test_workflow_does_not_pin_the_artifact_version(self, build_workflow: str) -> None:
        assert __version__ not in build_workflow, (
            "build.yml must glob the artifact instead of pinning a version string"
        )
        assert "uses: actions/upload-release-asset" not in build_workflow, (
            "actions/upload-release-asset is archived; use the preinstalled gh CLI"
        )
        assert "gh release upload" in build_workflow

    def test_workflow_installs_the_dev_requirements(self, build_workflow: str) -> None:
        assert "requirements-dev.txt" in build_workflow, (
            "pytest moved out of requirements.txt; CI must install the dev set"
        )

    def test_workflow_runs_the_real_qa_gate(self, build_workflow: str) -> None:
        assert "--validate" in build_workflow
        assert "make_release_manifest.py" in build_workflow
