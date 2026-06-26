# Repository Guidelines

## Project Structure & Module Organization
- `prose/` is the normal system-Python package. Prose is intentionally not a uv project because it must use the local LibreOffice UNO Python bridge.
- `prose/cli.py` defines the `python3 -m prose app [document.odt]` command. Route launch behavior through this module rather than adding root-level scripts.
- `prose/app.py` owns `ProseApp` and `ProseWindow`, including the main GTK/Libadwaita UI, action registration, Writer interaction, Text Draft behavior, LLM calls, and proofreader flow.
- `prose/windows/settings.py` owns the Settings window. `prose/windows/editor_commands.py` owns the command-copy/run window.
- `prose/runtime.py` contains shared dataclasses, config helpers, action metadata, UNO setup helpers, and app-wide utility functions used across the package.
- `prose/paths.py` centralizes project-root paths such as `config.json`, `scripts/`, and `prompts/`.
- `config.json` stores local runtime settings and may contain API keys. Do not commit real credentials.
- `prompts/` holds baseline prompt text used by the app. `scripts/` holds bundled Text Draft external-action wrappers.

## Build, Test, and Development Commands
- `python3 -m prose app` runs the desktop app locally.
- `python3 -m prose app /path/to/document.odt` launches Prose and opens a Writer document.
- `python3 -m prose --help` checks the package CLI.
- `python3 -m py_compile prose/*.py prose/windows/*.py` is the quick syntax check.
- There is no uv environment, no pyproject, and no separate build step.

## Coding Style & Naming Conventions
- Follow standard Python conventions: 4-space indentation, snake_case for functions/variables, PascalCase for classes.
- Prefer type hints and dataclasses as used in the existing package.
- Keep UI strings and config keys centralized in shared package modules.
- Keep changes narrow and consistent with existing GTK/Libadwaita patterns.

## Testing Guidelines
- No automated Prose test suite is present. Validate changes with `python3 -m py_compile prose/*.py prose/windows/*.py`.
- For UI or UNO changes, run `python3 -m prose app` and exercise the affected flow manually.
- If launcher integration changes, validate the dependent project or desktop script as part of the same change.

## Integration Notes
- Desktop files live outside this repo in `/home/jesse/Dropbox/MCGLAW/config_files/Desktop_Files`; the launcher should `cd` into this project and run `python3 -m prose app "$@"`.
- CurrentCaseTui opens `.odt` files through Prose and should call `python3 -m prose app <document>` with this project as the working directory.
- Keep `APP_ID`, `ACTION_OBJECT_PATH`, and existing GApplication action names stable because Focus and hotkey workflows call them directly.

## Configuration & Security Notes
- Do not commit real API keys; keep `config.json` local or use placeholders.
- LibreOffice integration expects a local UNO socket and a normal non-Flatpak/non-Snap LibreOffice install.
