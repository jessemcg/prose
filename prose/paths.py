"""Filesystem paths for the project-root based Prose app."""

from __future__ import annotations

from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_DIR / "config.json"
SCRIPTS_DIR = PROJECT_DIR / "scripts"
PROMPTS_DIR = PROJECT_DIR / "prompts"
TEXT_DRAFT_CODEX_VTE_SCRIPT = SCRIPTS_DIR / "prose-text-draft-codex-vte.sh"
TEXT_DRAFT_PROJECTS_DIR = PROJECT_DIR.parent
