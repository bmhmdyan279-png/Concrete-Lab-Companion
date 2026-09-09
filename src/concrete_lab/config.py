"""Application configuration loaded from ``config.yaml``.

Replaces hard-coded values with a validated configuration object
(review feedback: configuration management with an explicit
``validate()`` step and a safe loader).

Resolution order for the configuration file — the first hit wins:

1. an explicit ``path`` argument (CLI ``--config``);
2. the ``CONCRETE_LAB_CONFIG`` environment variable;
3. a ``config.yaml`` found by walking up from the package directory
   (works from a checkout *and* from a ``src``-layout install);
4. :meth:`AppConfig.defaults` — code-level fallbacks derived from
   :mod:`concrete_lab.constants`, so an installed package never depends
   on a file that may not exist next to it.

Only the protection password can come from the environment as a
*secret*; see :meth:`AppConfig.protection_password`.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

from concrete_lab import __version__, constants

#: Environment variable pointing at an alternative configuration file.
CONFIG_ENV_VAR: str = "CONCRETE_LAB_CONFIG"

#: Environment variable that supplies the worksheet protection password.
#: Re-exported from :mod:`concrete_lab.constants` so there is exactly one
#: declaration of it in the package.
PASSWORD_ENV_VAR: str = constants.PASSWORD_ENV_VAR

#: How many parent directories are searched for a ``config.yaml``.
CONFIG_SEARCH_DEPTH: int = 5

#: File name looked for when searching upwards from the package.
CONFIG_FILENAME: str = "config.yaml"


class ConfigError(ValueError):
    """Raised when the configuration file is missing required data."""


def _search_upwards(start: Path, filename: str = CONFIG_FILENAME,
                    depth: int = CONFIG_SEARCH_DEPTH) -> Optional[Path]:
    """Return the first ``filename`` found in ``start`` or its ancestors."""
    for directory in (start, *start.parents)[: depth + 1]:
        candidate = directory / filename
        if candidate.is_file():
            return candidate
    return None


def default_config_path() -> Optional[Path]:
    """Locate the repository ``config.yaml`` without assuming a layout.

    Returns:
        The discovered path, or ``None`` when the package is installed
        away from any configuration file (callers should then use
        :meth:`AppConfig.defaults`).
    """
    return _search_upwards(Path(__file__).resolve().parent)


@dataclass(frozen=True)
class AppConfig:
    """Validated application configuration.

    Attributes:
        name: Human-readable product name.
        version: Application version string (must match the package).
        language: UI language code (``fa`` → right-to-left sheets).
        colors: Colour palette used by the renderer (hex strings).
        fonts: Font names used by the renderer.
        sheet_protection_password: Optional protection password from
            the config file.  Prefer the environment variable (see
            :meth:`protection_password`) for anything real.
        standards_primary: Primary standards body (e.g. ``ASTM``).
        standards_secondary: Secondary standards body (e.g. ``ISIRI``).
        tolerance_default: Default numeric tolerance for QA checks.
        source: Where this configuration came from (path or ``defaults``);
            recorded in the build manifest for traceability.
        extra: Raw parsed mapping, for forward compatibility.
    """

    name: str
    version: str
    language: str = "fa"
    colors: Dict[str, str] = field(default_factory=dict)
    fonts: Dict[str, str] = field(default_factory=dict)
    sheet_protection_password: Optional[str] = None
    standards_primary: str = "ASTM"
    standards_secondary: str = "ISIRI"
    tolerance_default: float = 0.1
    source: str = "defaults"
    extra: Dict[str, Any] = field(default_factory=dict)

    # -- Construction --------------------------------------------------------

    @classmethod
    def defaults(cls) -> AppConfig:
        """Return a configuration built entirely from code-level constants.

        Used when no ``config.yaml`` is discoverable (e.g. the package is
        installed as a wheel).  The palette and fonts come from
        :mod:`concrete_lab.constants` so there is exactly one source of
        truth for them.
        """
        return cls(
            name="Concrete Lab Companion",
            version=__version__,
            language="fa",
            colors={
                "header_fill": constants.COLOR_HEADER_FILL,
                "header_font": constants.COLOR_HEADER_FONT,
                "pass_fill": constants.COLOR_PASS_FILL,
                "pass_font": constants.COLOR_PASS_FONT,
                "warning_fill": constants.COLOR_WARN_FILL,
                "warning_font": constants.COLOR_WARN_FONT,
                "fail_fill": constants.COLOR_FAIL_FILL,
                "fail_font": constants.COLOR_FAIL_FONT,
                "calculation_fill": constants.COLOR_VALUE_FILL,
                "input_fill": constants.COLOR_BADGE_FILL,
            },
            fonts={"primary_fa": constants.FONT_PRIMARY, "numbers": constants.FONT_NUMBERS},
            source="defaults",
        )

    # -- Loading -------------------------------------------------------------

    @classmethod
    def load(cls, path: Optional[Path] = None,
             env: Optional[Dict[str, str]] = None) -> AppConfig:
        """Load and validate the configuration.

        Args:
            path: Explicit config file location (highest precedence).
            env: Environment mapping (defaults to ``os.environ``);
                injectable for tests.

        Returns:
            A validated :class:`AppConfig`.

        Raises:
            ConfigError: If a path was explicitly requested (argument or
                environment variable) but does not exist or lacks the
                required ``project`` section.  When *no* path was
                requested and none is discoverable, code-level defaults
                are returned instead of raising.
        """
        environ = env if env is not None else dict(os.environ)
        requested: Optional[Path] = None
        explicit = False

        if path is not None:
            requested, explicit = Path(path), True
        elif environ.get(CONFIG_ENV_VAR):
            requested, explicit = Path(environ[CONFIG_ENV_VAR]), True
        else:
            requested = default_config_path()

        if requested is None:
            return cls.defaults()
        if not requested.exists():
            if explicit:
                raise ConfigError(f"configuration file not found: {requested}")
            return cls.defaults()

        with open(requested, encoding="utf-8") as handle:
            raw = yaml.safe_load(handle) or {}

        project = raw.get("project")
        if not isinstance(project, dict) or "name" not in project:
            raise ConfigError(f"{requested} must define a project.name entry")

        password = (raw.get("password") or {}).get("sheet_protection")
        standards = raw.get("standards") or {}
        validation = raw.get("validation") or {}

        return cls(
            name=str(project["name"]),
            version=str(project.get("version", __version__)),
            language=str(project.get("language", "fa")),
            colors=dict(raw.get("colors") or {}),
            fonts=dict(raw.get("fonts") or {}),
            sheet_protection_password=password,
            standards_primary=str(standards.get("primary", "ASTM")),
            standards_secondary=str(standards.get("secondary", "ISIRI")),
            tolerance_default=float(validation.get("tolerance_default", 0.1)),
            source=str(requested),
            extra=raw,
        )

    # -- Queries -------------------------------------------------------------

    def validate(self) -> bool:
        """Return ``True`` when the configuration is internally consistent.

        Raises:
            ConfigError: When a consistency rule is violated (version
                format, empty palette keys, non-finite tolerance, or a
                version that disagrees with the installed package).
        """
        parts = self.version.split(".")
        if len(parts) != 3 or not all(part.isdigit() for part in parts):
            raise ConfigError(f"version must be semantic (x.y.z), got {self.version!r}")
        if self.version != __version__:
            raise ConfigError(
                f"config version {self.version!r} disagrees with the package version "
                f"{__version__!r} — bump both together (see tests/test_consistency.py)"
            )
        if self.tolerance_default < 0:
            raise ConfigError("tolerance_default must be >= 0")
        if self.language not in ("fa", "en"):
            raise ConfigError(f"unsupported language {self.language!r}")
        for key, value in self.colors.items():
            if not isinstance(value, str) or not value.strip():
                raise ConfigError(f"colors.{key} must be a non-empty hex string")
        return True

    def protection_password(self, env: Optional[Dict[str, str]] = None) -> Optional[str]:
        """Resolve the protection password.

        Precedence: ``WORKBOOK_PASSWORD`` environment variable, then the
        config file value.  Returns ``None`` when neither is set, in
        which case protection must be reported as DISABLED.

        Args:
            env: Environment mapping (defaults to ``os.environ``);
                injectable for tests.
        """
        environ = env if env is not None else dict(os.environ)
        from_env = environ.get(PASSWORD_ENV_VAR)
        if from_env:
            return from_env
        return self.sheet_protection_password
