# Prose

<img src="prose.png" alt="Prose" width="128" align="left">

Prose is a GTK4/Libadwaita desktop application for AI-assisted editing inside LibreOffice Writer. It connects to a local LibreOffice Writer process through the UNO bridge, opens Writer documents in a dedicated Prose LibreOffice profile, and runs prompt-driven tools such as proofreading, SpellingStyle, shorten, translate, reference lookup, and related editing helpers.

## Important Requirements

### LibreOffice must be a normal `.deb` or `.rpm` install

Prose depends on LibreOffice's local UNO bridge and Python UNO modules. For that reason, LibreOffice must be installed as a standard system package such as:

- a `.deb` package on Debian/Ubuntu-based systems
- an `.rpm` package on Fedora/RHEL/openSUSE-based systems

Do not use:

- Flatpak LibreOffice
- Snap LibreOffice

Those sandboxed packaging formats commonly interfere with one or more of the pieces Prose needs:

- direct access to LibreOffice's `python-uno` bridge files
- local process launching with the expected `soffice` binary path
- local filesystem/profile access used by `-env:UserInstallation=...`
- reliable local UNO socket communication

If LibreOffice is installed through Flatpak or Snap, Prose may fail to:

- import `python-uno`
- connect to Writer through UNO
- open documents in the dedicated Prose LibreOffice profile
- reuse the user's LibreOffice profile content

If you want Prose to work reliably, install a normal non-sandboxed LibreOffice package first.

### You also need a separate speech-to-text app

Prose does not perform speech recognition itself. If you want to drive Prose from dictated text, you must use a separate speech-to-text application that writes recognized text to a plain text file at a fixed location in memory or on disk.

Prose reads that file as its source text input. In practice this means:

- your speech-to-text app must save output to a text file
- that file must live at a stable path Prose can access
- you must point Prose's `Source text file` setting to that path

Without an external speech-to-text app that saves to a text file, features such as `SpellingStyle` and `Direct Input` will not have source text to process.

## How Prose Works

Prose uses LibreOffice Writer as the live editing surface and sends selected document text or external source text to an LLM. Depending on the action, it then:

- inserts generated text into Writer
- rewrites selected text
- proofreads page ranges
- creates shorter or alternative drafts
- performs reference-style lookups

The app launches LibreOffice with a dedicated Prose profile located at:

`~/.config/libreoffice-prose-profile`

This keeps Prose's Writer session isolated from your ordinary day-to-day LibreOffice profile. The Settings window also includes a command to copy your normal LibreOffice profile into the Prose profile if you want Prose to inherit your dictionaries, AutoText, templates, or other user settings.

## Core Dependencies

- Python 3
- GTK 4
- Libadwaita 1
- LibreOffice Writer installed as a normal `.deb` or `.rpm`
- LibreOffice Python UNO bridge files available locally
- network/API access for whichever LLM provider you configure
- Tavily CLI installed locally for `Reference` and `Ask Field`

## Repository Layout

- `prose.py`: main application, UI, Writer integration, and AI actions
- `config.json`: local runtime settings such as API URLs, model IDs, keys, prompts, and last source file
- `prompts/`: baseline prompt text used by the app
- `prose.png`: app icon used in the README

## Running Prose

Start the app with:

```bash
python3 prose.py
```

There is no separate build step in this repository.

## First-Time Setup

### 1. Install normal LibreOffice

Install LibreOffice from your distribution's standard package system, not from Flatpak or Snap.

Examples:

- Debian/Ubuntu: install the standard `libreoffice` packages from `apt`
- Fedora/openSUSE/RHEL-family: install the standard LibreOffice packages from `dnf`, `yum`, or `zypper`

After installation, confirm that LibreOffice's program directory exists in a normal system path such as:

- `/usr/lib/libreoffice/program`
- `/usr/lib64/libreoffice/program`

### 2. Confirm Prose can find `python-uno`

Prose tries to auto-detect LibreOffice's Python bridge files. If it cannot, open Settings and set:

- `LibreOffice Python path`

This should usually point to LibreOffice's `program` directory, for example:

`/usr/lib/libreoffice/program`

### 3. Prepare your speech-to-text workflow

If you want dictation-driven editing:

- run a separate speech-to-text tool outside Prose
- configure that tool to save recognized text into a plain text file
- choose that file in Prose under `Settings -> Source text file`

This source file is what Prose uses for actions like:

- `Direct Input`
- `Direct Input No Trailing Space`
- `SpellingStyle`
- `Improve 1`
- `Improve 2`

### 4. Configure your model settings

Open Settings and enter the API information for the tools you want to use. Each tool family has its own prompt and model settings. Depending on the provider, you may need:

- API URL
- model ID
- API key
- prompt text
- Tavily API key for reference-oriented tools

`Reference` and `Ask Field` now use local Tavily CLI retrieval instead of provider-side MCP. Install the CLI once with:

```bash
uv tool install tavily-cli
```

Prose passes the Tavily API key from Settings to the CLI through `TAVILY_API_KEY`, so you do not need to run `tvly login` separately.

## LibreOffice Integration Details

Prose connects to LibreOffice through a local UNO socket at:

`127.0.0.1:2004`

When needed, Prose starts LibreOffice with:

- Writer mode
- a dedicated Prose profile
- local UNO accept arguments

Because the app launches and controls a local LibreOffice process directly, sandboxed LibreOffice packaging is a poor fit. A standard system package install is the supported setup.

## Using the App

### Launch Writer

Use `Launch Writer` to open a Writer document in the Prose-managed LibreOffice session.

### Choose a source text file

In Settings, set `Source text file` to the plain text file that Prose should read when using external-input tools.

### Run editing actions

Available actions include:

- `SpellingStyle`
- `Improve 1`
- `Improve 2`
- `Proof Reading`
- `Shorten`
- `Translate`
- `Topic Sentence`
- `Introduction`
- `Conclusion`
- `Concl. No Issues`
- `Concl. Section`
- `Reference`
- `Ask Field`
- `Combine Cites`

Some actions operate on the current Writer document or selection. Others use the configured external source text file.

`Reference` and `Ask Field` first run a local `tvly search --json` command and then send those search results to the configured model endpoint. This makes them work with standard OpenAI-compatible `/chat/completions` and `/responses` endpoints, including Fireworks-hosted models.

### Add Case

`Add Case` is designed for a selected case citation in Writer.

When you run it, Prose:

- reads the currently selected citation in Writer
- appends the citation and a shortened case-reference form to the concordance file
- creates a LibreOffice AutoText entry for that citation in the `Cases` AutoText group if one does not already exist

This is useful if you want to build up reusable case citations and related concordance entries while editing a brief.

The concordance file can then be used by LibreOffice's table-of-authorities workflow. In general, the concordance entries map citation text in the document to an authority category such as `Cases`. Once your citations have been added to the concordance file, LibreOffice can use that concordance data to mark matching authorities and generate a table of authorities from those marks. In other words, `Add Case` helps populate the citation data that LibreOffice can later use when building the table of authorities section of the brief.

### Optional: bind Prose commands to hotkeys

If you want a faster workflow, wire Prose commands to keyboard shortcuts in a hotkey tool such as `ckb-next`.

Prose includes an `Editor Commands` window that can copy runnable GApplication command lines for supported actions. A practical workflow is:

- open `Editor Commands` from the Prose menu
- find the action you want to trigger
- click `Copy Command`
- paste that command into a hotkey or macro action in `ckb-next`

This is especially useful for actions such as:

- `SpellingStyle`
- `Improve 1`
- `Improve 2`
- `Direct Input`

The main requirement is that Prose must already be running when the external hotkey tool sends the command.

## Profile Import

Settings includes a `Copy Normal Profile to Prose` action. This copies your normal LibreOffice profile into Prose's dedicated profile so Prose can inherit items such as:

- dictionaries
- AutoText
- templates
- custom LibreOffice preferences

This action:

- replaces the current Prose LibreOffice profile
- stops Prose's LibreOffice listener first if needed
- creates a backup of the previous Prose profile before replacing it

After a successful copy, launch Writer again from Prose to use the imported profile.

## Configuration Notes

`config.json` is local application state. It stores items such as:

- API URLs
- API keys
- model IDs
- prompts
- Tavily keys
- last selected source text file
- LibreOffice Python path

Do not commit real API keys.

## Troubleshooting

### `python-uno unavailable`

Likely causes:

- LibreOffice is installed via Flatpak or Snap
- LibreOffice is not installed in a standard system location
- the `LibreOffice Python path` setting is unset or wrong

Fix:

- install a normal `.deb` or `.rpm` LibreOffice package
- set `LibreOffice Python path` to LibreOffice's `program` directory

### Writer will not open through Prose

Likely causes:

- Prose cannot connect to the UNO bridge
- `python-uno` is unavailable
- LibreOffice is installed in a sandboxed packaging format

Fix:

- replace Flatpak/Snap LibreOffice with a normal `.deb` or `.rpm`
- confirm the LibreOffice Python path in Settings
- relaunch Prose and try `Launch Writer` again

### `Choose a source text file first`

This means you triggered an action that depends on external source text and Prose does not yet have a valid text file path configured.

Fix:

- configure your external speech-to-text app to save to a text file
- point Prose's `Source text file` setting to that file

## Development

Run locally with:

```bash
python3 prose.py
```

There is no automated test suite in this repository at the moment. A minimal verification step for changes is:

```bash
python3 -m py_compile prose.py
python3 prose.py
```

## Security

- keep `config.json` local
- do not commit live API credentials
- remember that source text files may contain sensitive dictated or legal text
