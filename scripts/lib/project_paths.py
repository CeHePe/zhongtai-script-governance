from __future__ import annotations

from pathlib import Path


ROOT_MARKERS = ("AGENTS.md", "README.txt", ".git")


def find_project_root(start: str | Path) -> Path:
    """Return the local project root without relying on the process cwd."""
    current = Path(start).resolve()
    if current.is_file():
        current = current.parent

    for candidate in (current, *current.parents):
        if any((candidate / marker).exists() for marker in ROOT_MARKERS):
            return candidate

    raise FileNotFoundError(f"Unable to locate project root from {start!s}")
