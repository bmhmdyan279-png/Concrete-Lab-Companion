"""Tests for configuration resolution (``concrete_lab.config``).

The loader used to assume a single repository layout: ``parents[2]`` from
the module file.  Installed as a package that path points into
``site-packages`` and the load fails.  These tests pin the documented
resolution order — explicit path, environment variable, upward search,
code-level defaults — and the consistency rules ``validate()`` enforces.
"""

from pathlib import Path

import pytest

from concrete_lab import __version__, constants
from concrete_lab.config import (
    CONFIG_ENV_VAR,
    CONFIG_FILENAME,
    PASSWORD_ENV_VAR,
    AppConfig,
    ConfigError,
    default_config_path,
)

MINIMAL = "project:\n  name: 'Test Lab'\n  version: '{version}'\n"


def write_config(directory: Path, *, version: str = __version__, extra: str = "") -> Path:
    """Write a minimal valid config file and return its path."""
    path = directory / CONFIG_FILENAME
    path.write_text(MINIMAL.format(version=version) + extra, encoding="utf-8")
    return path


class TestDefaults:
    def test_defaults_need_no_file_at_all(self) -> None:
        config = AppConfig.defaults()
        assert config.name and config.version == __version__
        assert config.source == "defaults"
        assert config.validate() is True

    def test_defaults_take_their_palette_from_constants(self) -> None:
        """One source of truth for colours: constants.py."""
        colors = AppConfig.defaults().colors
        assert colors["header_fill"] == constants.COLOR_HEADER_FILL
        assert colors["pass_fill"] == constants.COLOR_PASS_FILL
        assert colors["fail_font"] == constants.COLOR_FAIL_FONT

    def test_defaults_have_no_password(self) -> None:
        assert AppConfig.defaults().sheet_protection_password is None


class TestResolutionOrder:
    def test_explicit_path_wins(self, tmp_path: Path, monkeypatch) -> None:
        explicit = write_config(tmp_path, extra="  language: 'en'\n")
        monkeypatch.setenv(CONFIG_ENV_VAR, "/nonexistent/from-env.yaml")
        assert AppConfig.load(explicit).language == "en"

    def test_environment_variable_is_next(self, tmp_path: Path, monkeypatch) -> None:
        from_env = tmp_path / "from-env.yaml"
        from_env.write_text(
            MINIMAL.format(version=__version__) + "standards:\n  primary: 'EN'\n",
            encoding="utf-8",
        )
        monkeypatch.setenv(CONFIG_ENV_VAR, str(from_env))
        assert AppConfig.load().standards_primary == "EN"
        assert AppConfig.load().source == str(from_env)

    def test_environment_is_injectable_for_tests(self, tmp_path: Path) -> None:
        from_env = write_config(tmp_path, extra="  language: 'en'\n")
        config = AppConfig.load(env={CONFIG_ENV_VAR: str(from_env)})
        assert config.language == "en"

    def test_upward_search_finds_the_repository_file(self) -> None:
        found = default_config_path()
        assert found is not None and found.name == CONFIG_FILENAME
        assert found.is_file()

    def test_missing_explicit_path_is_an_error(self, tmp_path: Path) -> None:
        with pytest.raises(ConfigError, match="not found"):
            AppConfig.load(tmp_path / "absent.yaml")

    def test_missing_environment_path_is_an_error(self, tmp_path: Path, monkeypatch) -> None:
        monkeypatch.setenv(CONFIG_ENV_VAR, str(tmp_path / "absent.yaml"))
        with pytest.raises(ConfigError, match="not found"):
            AppConfig.load()

    def test_no_discoverable_file_falls_back_to_defaults(self, tmp_path: Path, monkeypatch) -> None:
        """An installed package with no config.yaml must still build."""
        import concrete_lab.config as config_module

        monkeypatch.delenv(CONFIG_ENV_VAR, raising=False)
        monkeypatch.setattr(config_module, "default_config_path", lambda: None)
        assert AppConfig.load().source == "defaults"

    def test_undiscoverable_file_falls_back_to_defaults(self, tmp_path: Path, monkeypatch) -> None:
        import concrete_lab.config as config_module

        monkeypatch.delenv(CONFIG_ENV_VAR, raising=False)
        monkeypatch.setattr(config_module, "default_config_path",
                            lambda: tmp_path / "gone.yaml")
        assert AppConfig.load().source == "defaults"


class TestParsing:
    def test_empty_file_is_rejected(self, tmp_path: Path) -> None:
        empty = tmp_path / CONFIG_FILENAME
        empty.write_text("", encoding="utf-8")
        with pytest.raises(ConfigError, match=r"project\.name"):
            AppConfig.load(empty)

    def test_project_must_be_a_mapping(self, tmp_path: Path) -> None:
        bad = tmp_path / CONFIG_FILENAME
        bad.write_text("project: 'just a string'\n", encoding="utf-8")
        with pytest.raises(ConfigError, match=r"project\.name"):
            AppConfig.load(bad)

    def test_optional_sections_have_defaults(self, tmp_path: Path) -> None:
        config = AppConfig.load(write_config(tmp_path))
        assert config.language == "fa"
        assert config.standards_primary == "ASTM"
        assert config.standards_secondary == "ISIRI"
        assert config.tolerance_default == 0.1
        assert config.colors == {} and config.fonts == {}

    def test_raw_mapping_is_kept_for_forward_compatibility(self, tmp_path: Path) -> None:
        path = write_config(tmp_path, extra="future:\n  feature: true\n")
        assert AppConfig.load(path).extra["future"]["feature"] is True


class TestValidation:
    @pytest.mark.parametrize("version", ["1.2", "1.2.3.4", "v1.2.3", "1.x.3", ""])
    def test_non_semantic_version_is_rejected(self, tmp_path: Path, version: str) -> None:
        config = AppConfig(name="X", version=version)
        with pytest.raises(ConfigError, match="semantic"):
            config.validate()

    def test_version_must_match_the_package(self) -> None:
        with pytest.raises(ConfigError, match="disagrees with the package version"):
            AppConfig(name="X", version="0.0.1").validate()

    def test_negative_tolerance_is_rejected(self) -> None:
        config = AppConfig(name="X", version=__version__, tolerance_default=-0.1)
        with pytest.raises(ConfigError, match="tolerance_default"):
            config.validate()

    @pytest.mark.parametrize("language", ["de", "", "FA"])
    def test_unsupported_language_is_rejected(self, language: str) -> None:
        config = AppConfig(name="X", version=__version__, language=language)
        with pytest.raises(ConfigError, match="unsupported language"):
            config.validate()

    @pytest.mark.parametrize("colors", [{"header_fill": ""}, {"header_fill": "   "},
                                        {"header_fill": 1234}])
    def test_empty_or_non_string_colour_is_rejected(self, colors: dict) -> None:
        config = AppConfig(name="X", version=__version__, colors=colors)
        with pytest.raises(ConfigError, match=r"colors\.header_fill"):
            config.validate()

    def test_repository_config_is_valid(self) -> None:
        assert AppConfig.load().validate() is True


class TestPasswordResolution:
    def test_environment_wins_over_the_file(self) -> None:
        config = AppConfig(name="X", version=__version__, sheet_protection_password="from-file")
        assert config.protection_password(env={PASSWORD_ENV_VAR: "from-env"}) == "from-env"

    def test_file_value_is_the_fallback(self) -> None:
        config = AppConfig(name="X", version=__version__, sheet_protection_password="from-file")
        assert config.protection_password(env={}) == "from-file"

    def test_no_source_means_no_protection(self) -> None:
        config = AppConfig(name="X", version=__version__)
        assert config.protection_password(env={}) is None

    def test_empty_environment_value_is_ignored(self) -> None:
        config = AppConfig(name="X", version=__version__, sheet_protection_password="from-file")
        assert config.protection_password(env={PASSWORD_ENV_VAR: ""}) == "from-file"

    def test_environment_constant_is_shared_with_constants_module(self) -> None:
        assert PASSWORD_ENV_VAR == constants.PASSWORD_ENV_VAR
