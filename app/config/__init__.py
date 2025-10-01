"""Configuration helpers exposed at the package level."""

from pathlib import Path

from .settings import settings


def get_project_root(project_type: str) -> Path:
	"""Return the base directory for the given project type."""
	return settings.get_base_dir_for_project_type(project_type)


def get_templates_root() -> Path:
	"""Return the root directory where Word templates are stored."""
	candidate = settings.app_dir / "templates"
	candidate.mkdir(parents=True, exist_ok=True)
	return candidate


TEMPLATES_ROOT = get_templates_root()

__all__ = [
	"settings",
	"get_project_root",
	"TEMPLATES_ROOT",
]
