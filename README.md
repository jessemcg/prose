# Prose

![Prose icon](prose.png)

Prose is a GTK/Libadwaita desktop app that connects to LibreOffice Writer through UNO and runs AI-powered writing helpers (proofread, improve, shorten, thesaurus, references, and more).

## Features
- LibreOffice Writer integration via UNO socket
- Configurable prompts and model settings per tool
- Prompt editing UI with saved defaults in `config.json`
- Built-in helpers for citations, thesaurus, and reference lookups

## Requirements
- Python 3
- GTK 4 + Libadwaita
- LibreOffice with UNO socket listening on `127.0.0.1:2004`

## Run
```bash
python3 prose.py
```

## Configuration
- `config.json` stores API URLs, keys, model IDs, prompts, and the last editor source file
- `prompts/` contains baseline prompt text used by the app

## Notes
- Do not commit real API keys; keep `config.json` local or use placeholders
- The app expects a local UNO socket; if the port/host changes, update it in `prose.py`
