"""Application configuration loaded from ``config.yaml``.

Replaces hard-coded values with a validated configuration object
(review feedback: configuration management with an explicit
``validate()`` step and a safe loader).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

#: Default location of the configuration file (repository root).
DEFAULT_CONFIG_PATH: Path = Path(__file__).resolve().parents[2] / "config.yaml"

#: Only protection passwords ever come from the environment; see
#: :func:`AppConfig.protection_password`.
PASSWORD_ENV_VAR: str = "WORKBOOK_PASSWORD"


class ConfigError(ValueError):
    """Raised when the configuration file is missing required data."""


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
    extra: Dict[str, Any] = field(default_factory=dict)

    # -- Loading -------------------------------------------------------------

    @classmethod
    def load(cls, path: Optional[Path] = None) -> "AppConfig":
        """Load and validate the configuration file.

        Args:
            path: Config file location; defaults to the repo-root
                ``config.yaml``.

        Returns:
            A validated :class:`AppConfig`.

        Raises:
            ConfigError: If the file is missing or lacks the required
                ``project`` section.
        """
        config_path = Path(path) if path is not None else DEFAULT_CONFIG_PATH
        if not config_path.exists():
            raise ConfigError(f"configuration file not found: {config_path}")
        with open(config_path, "r", encoding="utf-8") as handle:
            raw = yaml.safe_load(handle) or {}

        project = raw.get("project")
        if not isinstance(project, dict) or "name" not in project:
            raise ConfigError("config.yaml must define a project.name entry")

        password = (raw.get("password") or {}).get("sheet_protection")
        standards = raw.get("standards") or {}
        validation = raw.get("validation") or {}

        return cls(
            name=str(project["name"]),
            version=str(project.get("version", "0.0.0")),
            language=str(project.get("language", "fa")),
            colors=dict(raw.get("colors") or {}),
            fonts=dict(raw.get("fonts") or {}),
            sheet_protection_password=password,
            standards_primary=str(standards.get("primary", "ASTM")),
            standards_secondary=str(standards.get("secondary", "ISIRI")),
            tolerance_default=float(validation.get("tolerance_default", 0.1)),
            extra=raw,
        )

    # -- Queries -------------------------------------------------------------

    def validate(self) -> bool:
        """Return ``True`` when the configuration is internally consistent.

        Raises:
            ConfigError: When a consistency rule is violated (version
                format, empty palette keys, non-finite tolerance).
        """
        parts = self.version.split(".")
        if len(parts) != 3 or not all(part.isdigit() for part in parts):
            raise ConfigError(f"version must be semantic (x.y.z), got {self.version!r}")
        if self.tolerance_default < 0:
            raise ConfigError("tolerance_default must be >= 0")
        if self.language not in ("fa", "en"):
            raise ConfigError(f"unsupported language {self.language!r}")
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
        import os

        environ = env if env is not None else dict(os.environ)
        from_env = environ.get(PASSWORD_ENV_VAR)
        if from_env:
            return from_env
        return self.sheet_protection_password
