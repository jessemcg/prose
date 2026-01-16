# Repository Guidelines

## Project Structure & Module Organization
- `prose.py` is the main GTK/Libadwaita application and contains the UI plus LibreOffice/UNO integration.
- `config.json` stores local runtime settings (API URLs, model IDs, keys, prompts, last editor source file).
- `prompts/` holds baseline prompt text used by the app.
- `focus_sample_for_prose_coding_tasks.py` is a reference app used for patterns; it is not executed by Prose.

## Build, Test, and Development Commands
- `python3 prose.py` runs the desktop app locally.
- There is no separate build step or test runner configured in this repo.

## Coding Style & Naming Conventions
- Follow standard Python conventions (PEP 8): 4-space indentation, snake_case for functions/variables, PascalCase for classes.
- Prefer type hints and dataclasses as used in `prose.py`.
- Keep UI strings and config keys centralized near the top of the file.
- No formatter or linter is configured; keep changes minimal and consistent with existing style.

## Testing Guidelines
- No automated tests are present. Validate changes by running `python3 prose.py` and exercising the affected UI flow.
- If you add tests in the future, keep them in a new `tests/` directory and name files `test_*.py`.

## Commit & Pull Request Guidelines
- The Git history is empty, so no established commit style exists. Use concise, imperative messages (e.g., "Add config validation").
- PRs should include a short summary, manual testing notes, and screenshots or screen recordings for UI changes.
- Call out any `config.json` changes or new config keys explicitly.

## Configuration & Security Notes
- Do not commit real API keys; keep `config.json` local or use placeholders.
- LibreOffice integration expects a local UNO socket; if it changes, document it in the PR.
