#!/usr/bin/python3

from __future__ import annotations

import difflib
import json
import importlib
import os
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import gi

from .paths import CONFIG_FILE, PROJECT_DIR

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
from gi.repository import Adw, Gdk, Gio, GLib, Gtk, Pango  # type: ignore

uno = None  # type: ignore
PropertyValue = None  # type: ignore[assignment]
NoConnectException = Exception  # type: ignore[assignment]
XModel = None  # type: ignore[assignment]
ControlCharacter = None  # type: ignore[assignment]
XTextDocument = None  # type: ignore[assignment]


APP_ID = "com.mcglaw.Prose"
ACTION_OBJECT_PATH = "/com/mcglaw/Prose"
APP_NAME = "Prose"
GLib.set_application_name(APP_NAME)

TEXT_DRAFT_EXTERNAL_ACTION_ENV_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
STYLE_RULES_TOKEN = "{{STYLE_RULES}}"
STYLE_RULES_HINT_TEXT = (
    f"Use {STYLE_RULES_TOKEN} to insert the shared style rules from the Style Rules page."
)
CONFIG_KEY_PROOFREAD_API_URL = "api_url"
CONFIG_KEY_PROOFREAD_MODEL_ID = "model_id"
CONFIG_KEY_PROOFREAD_API_KEY = "api_key"
CONFIG_KEY_PROOFREAD_PROMPT = "proofread_prompt"
CONFIG_KEY_PROOFREAD_DISABLE_REASONING = "proofread_disable_reasoning"
CONFIG_KEY_PROOFREAD_SUGGESTIONS_JSON_PATH = "proofread_suggestions_json_path"
CONFIG_KEY_SPELLING_API_URL = "spellingstyle_api_url"
CONFIG_KEY_SPELLING_MODEL_ID = "spellingstyle_model_id"
CONFIG_KEY_SPELLING_API_KEY = "spellingstyle_api_key"
CONFIG_KEY_SPELLING_PROMPT = "spellingstyle_prompt"
CONFIG_KEY_SPELLING_DISABLE_REASONING = "spellingstyle_disable_reasoning"
CONFIG_KEY_TEXT_DRAFT_SPELLING_PROMPT = "text_draft_spellingstyle_prompt"
CONFIG_KEY_CITATION_VALIDATOR_API_URL = "citation_validator_api_url"
CONFIG_KEY_CITATION_VALIDATOR_MODEL_ID = "citation_validator_model_id"
CONFIG_KEY_CITATION_VALIDATOR_API_KEY = "citation_validator_api_key"
CONFIG_KEY_CITATION_VALIDATOR_PROMPT = "citation_validator_prompt"
CONFIG_KEY_CITATION_VALIDATOR_DISABLE_REASONING = "citation_validator_disable_reasoning"
CONFIG_KEY_IMPROVE1_API_URL = "improve1_api_url"
CONFIG_KEY_IMPROVE1_MODEL_ID = "improve1_model_id"
CONFIG_KEY_IMPROVE1_API_KEY = "improve1_api_key"
CONFIG_KEY_IMPROVE1_PROMPT = "improve1_prompt"
CONFIG_KEY_IMPROVE1_DISABLE_REASONING = "improve1_disable_reasoning"
CONFIG_KEY_IMPROVE2_API_URL = "improve2_api_url"
CONFIG_KEY_IMPROVE2_MODEL_ID = "improve2_model_id"
CONFIG_KEY_IMPROVE2_API_KEY = "improve2_api_key"
CONFIG_KEY_IMPROVE2_PROMPT = "improve2_prompt"
CONFIG_KEY_IMPROVE2_DISABLE_REASONING = "improve2_disable_reasoning"
CONFIG_KEY_COMBINE_CITES_API_URL = "combine_cites_api_url"
CONFIG_KEY_COMBINE_CITES_MODEL_ID = "combine_cites_model_id"
CONFIG_KEY_COMBINE_CITES_API_KEY = "combine_cites_api_key"
CONFIG_KEY_COMBINE_CITES_PROMPT = "combine_cites_prompt"
CONFIG_KEY_COMBINE_CITES_DISABLE_REASONING = "combine_cites_disable_reasoning"
CONFIG_KEY_THESAURUS_API_URL = "thesaurus_api_url"
CONFIG_KEY_THESAURUS_MODEL_ID = "thesaurus_model_id"
CONFIG_KEY_THESAURUS_API_KEY = "thesaurus_api_key"
CONFIG_KEY_THESAURUS_PROMPT = "thesaurus_prompt"
CONFIG_KEY_THESAURUS_DISABLE_REASONING = "thesaurus_disable_reasoning"
CONFIG_KEY_REFERENCE_API_URL = "reference_api_url"
CONFIG_KEY_REFERENCE_MODEL_ID = "reference_model_id"
CONFIG_KEY_REFERENCE_API_KEY = "reference_api_key"
CONFIG_KEY_REFERENCE_TAVILY_API_KEY = "reference_tavily_api_key"
CONFIG_KEY_REFERENCE_PROMPT = "reference_prompt"
CONFIG_KEY_REFERENCE_ASK_PROMPT = "reference_ask_prompt"
CONFIG_KEY_REFERENCE_DISABLE_REASONING = "reference_disable_reasoning"
CONFIG_KEY_ASK_API_URL = "ask_api_url"
CONFIG_KEY_ASK_MODEL_ID = "ask_model_id"
CONFIG_KEY_ASK_API_KEY = "ask_api_key"
CONFIG_KEY_ASK_TAVILY_API_KEY = "ask_tavily_api_key"
CONFIG_KEY_ASK_PROMPT = "ask_prompt"
CONFIG_KEY_ASK_DISABLE_REASONING = "ask_disable_reasoning"
CONFIG_KEY_SHORTEN_API_URL = "shorten_api_url"
CONFIG_KEY_SHORTEN_MODEL_ID = "shorten_model_id"
CONFIG_KEY_SHORTEN_API_KEY = "shorten_api_key"
CONFIG_KEY_SHORTEN_PROMPT = "shorten_prompt"
CONFIG_KEY_SHORTEN_DISABLE_REASONING = "shorten_disable_reasoning"
CONFIG_KEY_INTRO_API_URL = "introduction_api_url"
CONFIG_KEY_INTRO_MODEL_ID = "introduction_model_id"
CONFIG_KEY_INTRO_API_KEY = "introduction_api_key"
CONFIG_KEY_INTRO_PROMPT = "introduction_prompt"
CONFIG_KEY_INTRO_DISABLE_REASONING = "introduction_disable_reasoning"
CONFIG_KEY_INTRO_REPLY_API_URL = "introduction_reply_api_url"
CONFIG_KEY_INTRO_REPLY_MODEL_ID = "introduction_reply_model_id"
CONFIG_KEY_INTRO_REPLY_API_KEY = "introduction_reply_api_key"
CONFIG_KEY_INTRO_REPLY_PROMPT = "introduction_reply_prompt"
CONFIG_KEY_INTRO_REPLY_DISABLE_REASONING = "introduction_reply_disable_reasoning"
CONFIG_KEY_CONCLUSION_API_URL = "conclusion_api_url"
CONFIG_KEY_CONCLUSION_MODEL_ID = "conclusion_model_id"
CONFIG_KEY_CONCLUSION_API_KEY = "conclusion_api_key"
CONFIG_KEY_CONCLUSION_PROMPT = "conclusion_prompt"
CONFIG_KEY_CONCLUSION_DISABLE_REASONING = "conclusion_disable_reasoning"
CONFIG_KEY_CONCL_NO_ISSUES_API_URL = "concl_no_issues_api_url"
CONFIG_KEY_CONCL_NO_ISSUES_MODEL_ID = "concl_no_issues_model_id"
CONFIG_KEY_CONCL_NO_ISSUES_API_KEY = "concl_no_issues_api_key"
CONFIG_KEY_CONCL_NO_ISSUES_PROMPT = "concl_no_issues_prompt"
CONFIG_KEY_CONCL_NO_ISSUES_DISABLE_REASONING = "concl_no_issues_disable_reasoning"
CONFIG_KEY_TOPIC_SENTENCE_API_URL = "topic_sentence_api_url"
CONFIG_KEY_TOPIC_SENTENCE_MODEL_ID = "topic_sentence_model_id"
CONFIG_KEY_TOPIC_SENTENCE_API_KEY = "topic_sentence_api_key"
CONFIG_KEY_TOPIC_SENTENCE_PROMPT = "topic_sentence_prompt"
CONFIG_KEY_TOPIC_SENTENCE_DISABLE_REASONING = "topic_sentence_disable_reasoning"
CONFIG_KEY_CONCL_SECTION_API_URL = "concl_section_api_url"
CONFIG_KEY_CONCL_SECTION_MODEL_ID = "concl_section_model_id"
CONFIG_KEY_CONCL_SECTION_API_KEY = "concl_section_api_key"
CONFIG_KEY_CONCL_SECTION_PROMPT = "concl_section_prompt"
CONFIG_KEY_CONCL_SECTION_DISABLE_REASONING = "concl_section_disable_reasoning"
CONFIG_KEY_TRANSLATE_API_URL = "translate_api_url"
CONFIG_KEY_TRANSLATE_MODEL_ID = "translate_model_id"
CONFIG_KEY_TRANSLATE_API_KEY = "translate_api_key"
CONFIG_KEY_TRANSLATE_PROMPT = "translate_prompt"
CONFIG_KEY_TRANSLATE_DISABLE_REASONING = "translate_disable_reasoning"
CONFIG_KEY_MODEL_PROFILES = "model_profiles"
CONFIG_KEY_COMMAND_DEFAULT_PROFILES = "command_default_profiles"
CONFIG_KEY_EDITOR_COMMAND_DEFAULT_PROFILES = "editor_command_default_profiles"
CONFIG_KEY_RT_PREFIX = "rt_prefix"
CONFIG_KEY_CT_PREFIX = "ct_prefix"
CONFIG_KEY_WORD_SUBSTITUTIONS = "word_substitutions"
CONFIG_KEY_EDITOR_SOURCE_FILE = "editor_source_file"
CONFIG_KEY_LAST_ODT_FILE = "last_odt_file"
CONFIG_KEY_TEXT_DRAFT_TEMPLATE_DIR = "text_draft_template_dir"
CONFIG_KEY_LIBREOFFICE_PYTHON_PATH = "libreoffice_python_path"
CONFIG_KEY_CONCORDANCE_FILE_PATH = "concordance_file_path"
CONFIG_KEY_EDITOR_PINNED_ACTIONS = "editor_pinned_actions"
CONFIG_KEY_TEXT_DRAFT_PINNED_ACTIONS = "text_draft_pinned_actions"
CONFIG_KEY_SHARED_STYLE_RULES = "shared_style_rules"
CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION = "text_draft_external_action"
CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTIONS = "text_draft_external_actions"
TEXT_DRAFT_EXTERNAL_ACTION_DRAFT_FILE_TOKEN = "{draft_file}"
DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON = "utilities-terminal-symbolic"

DEFAULT_SHARED_STYLE_RULES = """## STYLE RULES

Priority rule: if text appears inside quotation marks, preserve it exactly as written. Do not change spelling, punctuation, capitalization, spacing, or characters inside quotation marks, even if another style rule would otherwise apply.

### 1. Tone and formality
- Use formal legal writing.
- Do not use contractions.
- Write "is not," "cannot," and "should not" instead of "isn't," "can't," and "shouldn't."

### 2. Mother / father
- Capitalize "Mother" or "Father" only when the word begins a sentence.
- Use lowercase "mother" or "father" when the word appears in the middle of a sentence.
- Do not insert "the" before "mother" or "father" unless the SOURCE text already includes "the."

Examples:
"Mother told the juvenile court that father struck her in the face."
"The juvenile court advised mother and father to be quiet during the hearing."

### 3. Agency / department
- Refer to the entity as "DCFS," "the agency," or "the department."
- Capitalize "agency" or "department" only when the word begins a sentence.

### 4. Foster care
- Always write "foster care" as two words.
- Do not hyphenate it.

### 5. Parents
- If the SOURCE text uses "parents," keep "parents."

### 6. Attorney and role titles
- Write position titles in lowercase.
- Examples: "children's counsel," "county counsel."

### 7. Dates
- If the SOURCE text provides a specific calendar date, write it in long-form U.S. style.
- Example: "June 15, 2023" or "August 25, 2025."
- Convert numeric dates to long-form U.S. style.
- Example: "02/24/2024" becomes "February 24, 2024."
- Do not convert relative or vague time expressions into calendar dates.
- Preserve terms such as "yesterday," "today," "tomorrow," "last week," "earlier this month," and "mid-April 2025" as written.
- If a modifier appears before a month, keep the modifier and hyphenate it.
- Examples: "early-January 2025," "mid-March 2024," "late-May 2022."
- Never invent, infer, or calculate a date that does not appear in the SOURCE text.

### 8. Kinship terms
- Use "maternal" for the mother's side of the family.
- Use "paternal" for the father's side of the family.
- Do not write phrases such as "mother's mother" or "father's father" when "maternal" or "paternal" would express the relationship.

### 9. Initials in names
- If a person's or child's name appears as initials, add a period after each letter.
- Do not add spaces between the letters.
- "BR" becomes "B.R."
- "AL" becomes "A.L."
- "FP" becomes "F.P."
- "TWT" becomes "T.W.T."

Examples:
"The juvenile court placed the children, BR and AL, in foster care" becomes "The juvenile court placed the children, B.R. and A.L., in foster care."
"The children's legal guardian, FP, requested custody of the children" becomes "The children's legal guardian, F.P., requested custody of the children."
"The parents of the children are TWT and FT" becomes "The parents of the children are T.W.T. and F.T."

### 10. Statutes
- Write "section" in lowercase.
- Write subdivision letters and numbers in lowercase inside parentheses.

Examples:
"The juvenile court sustained the findings under section 300, subdivision (b)(1)."
"The juvenile court ordered the children removed under section 361, subdivision (c)."
"Evidence Code section 720 authorizes the juvenile court to appoint an expert."

### 11. Special capitalizations
- Assume there is one respondent. Write "Respondent's Brief."
- Treat "Phoenix H." as a case name. Write "Phoenix H. Brief."

### 12. Quotations
- Preserve quoted content exactly.
- Do not normalize or replace any character inside quotation marks.
- Outside quoted SOURCE text, never use straight ASCII quotation marks or apostrophes.
- Prohibited characters in normal output:
  - U+0022 (")
  - U+0027 (')
- Use only typographic quotation marks and apostrophes in normal output:
  - U+201C LEFT DOUBLE QUOTATION MARK (“)
  - U+201D RIGHT DOUBLE QUOTATION MARK (”)
  - U+2018 LEFT SINGLE QUOTATION MARK (‘)
  - U+2019 RIGHT SINGLE QUOTATION MARK (’)
- Possessives must use U+2019.
- Example: “Jesse’s”
- If quoted SOURCE text contains ASCII quotation marks or apostrophes, preserve them unchanged inside that quoted material because the preservation rule controls.

### 13. Use of "that"
- Prefer adding "that" after reporting verbs when the verb introduces a clause.
- Common examples of reporting verbs include "said," "explained," "reported," "testified," and "believed."
- Use "that" when it makes the sentence clearer, smoother, or less ambiguous.
- Example pattern: "Mother reported that father struck her."

### Quick checklist
Before finalizing, confirm that:
- there are no contractions;
- "mother" and "father" follow the capitalization rule;
- dates are in long-form U.S. style when a specific date is given;
- initials contain periods with no spaces;
- statutes use lowercase "section" and lowercase subdivision references;
- quoted text is unchanged;
- all quotation marks and apostrophes outside quoted SOURCE text are non-ASCII typographic characters;
- "that" appears after reporting verbs when needed for clarity;"""

DEFAULT_PROMPT = (
    "You are a meticulous legal proofreader. Improve clarity, fix grammar, and preserve legal meaning. "
    "Do not shorten or embellish facts. Respond with concrete edits, not generic advice.\n\n"
    f"{STYLE_RULES_TOKEN}"
)
PROOFREAD_CHUNK_PARAGRAPH_LIMIT = 12
PROOFREAD_CHUNK_CHAR_LIMIT = 6500
DEFAULT_SPELLINGSTYLE_PROMPT = (
    "Revise the source text for spelling, grammar, and style. Preserve meaning and facts.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the revised text."
)
DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT = (
    "Revise the source text for spelling, grammar, punctuation, and style. Improve clarity and flow for "
    "general-purpose writing while preserving meaning and facts. Do not add legal phrasing or legal style rules "
    "unless the source text already requires them.\n\n"
    "Return only the revised text."
)
DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT = (
    "You validate page-number text for appellate record citations.\n\n"
    "Convert the source text into only the page number or page range it represents. "
    "Return only ASCII digits, or two ASCII digit groups separated by a hyphen for a range. "
    "Convert spoken digits and number words to numerals, including 'oh' and 'o' as 0 when used in a digit sequence. "
    "Do not add RT, CT, prefixes, parentheses, punctuation, labels, explanations, or extra words. "
    "If there is no usable page number, return NONE.\n\n"
    "Examples:\n"
    "Two, oh three -> 203\n"
    "page thirty five -> 35\n"
    "one twenty to one twenty two -> 120-122\n"
    "no page given -> NONE"
)
DEFAULT_IMPROVE_PROMPT = (
    "Improve the following text for clarity and precision while preserving meaning.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the improved text."
)
DEFAULT_IMPROVE1_PROMPT = DEFAULT_IMPROVE_PROMPT
DEFAULT_IMPROVE2_PROMPT = (
    "Rephrase the following text while preserving meaning.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the rephrased text."
)
DEFAULT_COMBINE_CITES_PROMPT = ""
DEFAULT_THESAURUS_PROMPT = (
    "Return JSON only with an 'alternatives' array of synonyms or short phrases for the selected text. "
    "Example: {\"alternatives\": [\"option\", \"best choice\"]}."
)
DEFAULT_REFERENCE_PROMPT = (
    "Using only the Tavily search results provided below, provide the definition of a word or phrase "
    "and use the word or phrase in a sentence. Consult up to two sources and cite the url for each source. "
    "Present the information like this:\n\n"
    "DEFINITION\n"
    "This is the definition.\n\n"
    "EXAMPLE SENTENCE\n"
    "This is an example sentence.\n\n"
    "SOURCES\n"
    "One url per line.\n\n"
    "Do not use markdown formatting."
)
DEFAULT_REFERENCE_ASK_PROMPT = "Answer the following question."
DEFAULT_ASK_PROMPT = (
    "Using only the Tavily search results provided below, answer the question. "
    "Present the information like this:\n\n"
    "ANSWER\n"
    "This is the answer.\n\n"
    "SOURCES\n"
    "One url per line.\n\n"
    "Do not use markdown formatting."
)
DEFAULT_SHORTEN_PROMPT = (
    "Shorten the selected text by removing unnecessary facts while preserving meaning.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the shortened text."
)
DEFAULT_INTRO_PROMPT = (
    "Write a concise introduction for the brief based on the provided argument section.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the introduction."
)
DEFAULT_INTRO_REPLY_PROMPT = (
    "Write a concise introduction for the reply brief based on the provided argument section.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the introduction."
)
DEFAULT_CONCLUSION_PROMPT = (
    "Write a conclusion for the brief based on the provided argument section.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the conclusion."
)
DEFAULT_CONCL_NO_ISSUES_PROMPT = (
    "Write a conclusion for the brief based on the provided issues-considered section.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the conclusion."
)
DEFAULT_TOPIC_SENTENCE_PROMPT = (
    "Write a concise topic sentence that captures the central meaning of the paragraph.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the topic sentence."
)

DEFAULT_CONCL_SECTION_PROMPT = (
    "Write a concise conclusion for the provided section.\n\n"
    f"{STYLE_RULES_TOKEN}\n\n"
    "Return only the conclusion."
)
DEFAULT_TRANSLATE_PROMPT = (
    "Translate the provided text into Spanish. Preserve meaning, tone, line breaks, numbering, and punctuation. "
    "Do not summarize and do not omit content."
)
DEFAULT_RT_PREFIX = "2"
DEFAULT_CT_PREFIX = "2"
MAX_WORD_SUBSTITUTIONS = 3
MAX_PINNED_EDITOR_ACTIONS = 5
TEXT_DRAFT_TEMP_FLUSH_DELAY_MS = 200
STACKED_CHOICES_MIN_HEIGHT = 620
STACKED_CHOICES_MAX_HEIGHT = 900
STACKED_CHOICES_COMPACT_SOURCE_CHARS = 700
STACKED_CHOICES_FULL_SOURCE_CHARS = 2600
STACKED_CHOICES_COMPACT_SOURCE_LINES = 8
STACKED_CHOICES_FULL_SOURCE_LINES = 28
STACKED_CHOICES_ESTIMATED_CHARS_PER_LINE = 92
UNSET_PROFILE_LABEL = "Choose a profile..."
MODEL_PROFILE_IDS = ("profile1", "profile2", "profile3", "profile4")
DEFAULT_MODEL_PROFILE_NICKNAMES = {
    "profile1": "Profile 1",
    "profile2": "Profile 2",
    "profile3": "Profile 3",
    "profile4": "Profile 4",
}
PROFILE_BACKED_COMMAND_TITLES = {
    "spelling": "SpellingStyle",
    "improve-generated": "Improve Generated",
    "rephrase-generated": "Rephrase Generated",
    "improve-selected": "Improve Selected",
    "choices-improve": "Choices Improve",
    "choices-rephrase-1": "Choices Rephrase 1",
    "choices-rephrase-2": "Choices Rephrase 2",
    "combine": "Combine Cites",
    "thesaurus": "Thesaurus",
    "shorten": "Shorten",
    "intro": "Introduction",
    "intro-reply": "Introduction for Reply",
    "conclusion": "Conclusion",
    "concl-no-issues": "Conclusion No Issues",
    "topic-sentence": "Topic Sentence",
    "concl-section": "Section Conclusion",
    "translate": "Translate",
}
PROFILE_BACKED_COMMAND_KEYS = tuple(PROFILE_BACKED_COMMAND_TITLES.keys())
MULTI_DRAFT_DIFF_VARIANTS = frozenset({"improve", "rephrase", "rephrase-1", "rephrase-2"})
STACKED_CHOICE_VARIANTS = ("improve", "rephrase-1", "rephrase-2")
REGENERATE_INSERT_MODE_BY_ACTION = {
    "improve-generated": "improve",
    "rephrase-generated": "improve",
    "improve-selected": "improve",
    "shorten": "improve",
    "topic-sentence": "editor",
    "intro": "editor",
    "intro-reply": "editor",
    "conclusion": "editor",
    "concl-no-issues": "editor",
    "concl-section": "editor",
}
REGENERATE_SOURCE_BUFFER_ACTION_KEYS = frozenset({"improve-generated", "rephrase-generated", "improve-selected", "shorten"})
CITATION_NUMBER_WORDS = {
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
}

LIBREOFFICE_PROFILE = Path.home() / ".config" / "libreoffice-prose-profile"
DEFAULT_NORMAL_LIBREOFFICE_PROFILE = Path.home() / ".config" / "libreoffice" / "4"
UNO_BRIDGE_URL = "uno:socket,host=127.0.0.1,port=2004;urp;StarOffice.ComponentContext"
SOFFICE_PATH = Path("/usr/lib/libreoffice/program/soffice")
DEFAULT_LIBREOFFICE_PYTHON_PATH = Path("/usr/lib/libreoffice/program")
LIBREOFFICE_PYTHON_CANDIDATES = (
    Path("/usr/lib64/libreoffice/program"),
    Path("/usr/lib/libreoffice/program"),
    Path("/opt/libreoffice/program"),
    Path("/opt/libreoffice7.6/program"),
    Path("/opt/libreoffice7.5/program"),
    Path("/Applications/LibreOffice.app/Contents/Resources"),
    Path("/Applications/LibreOffice.app/Contents/MacOS"),
)
SPELLING_OUTPUT_FONT_SIZE_PX = 18
SPELLING_OUTPUT_PADDING_PX = 12
SPELLING_OUTPUT_CORNER_RADIUS_PX = 10
REFERENCE_OUTPUT_FONT_SIZE_PX = SPELLING_OUTPUT_FONT_SIZE_PX
REFERENCE_URL_RE = re.compile(r"https?://[^\s)\]]+")
TEXT_DRAFT_TEMPLATE_EMAIL_RE = re.compile(
    r"\[\[email:\s*(?P<label>.+?)\s*<(?P<address>[^<>\s]+@[^<>\s]+)>\s*\]\]"
)
TEXT_DRAFT_TEMPLATE_EMAIL_FALLBACK_RE = re.compile(r"(?P<address>[^<>\s]+@[^<>\s]+)")
TAVILY_MAX_RESULTS = 5
TAVILY_MAX_SOURCES = 2
TAVILY_SOURCE_EXCERPT_CHARS = 2000
TAVILY_CLI_INSTALL_HINT = "Install it with `uv tool install tavily-cli`."
ODT_IMPORT_XML_MEMBERS = frozenset({"content.xml", "styles.xml"})
ODT_PAGE_STYLE_ATTR_RE = re.compile(r'\s(?:style:master-page-name|style:page-number)="[^"]*"')
ODT_MASTER_PAGE_TAG_RE = re.compile(r"<style:master-page\b[^>]*>")
ODT_MASTER_PAGE_NEXT_STYLE_RE = re.compile(r'\sstyle:next-style-name="[^"]*"')
OBSOLETE_REASONING_CONFIG_KEYS = (
    "proofread_kimi_reasoning",
    "proofread_deepseek_reasoning",
    "spellingstyle_kimi_reasoning",
    "spellingstyle_deepseek_reasoning",
    "improve1_kimi_reasoning",
    "improve1_deepseek_reasoning",
    "improve2_kimi_reasoning",
    "improve2_deepseek_reasoning",
    "combine_cites_kimi_reasoning",
    "combine_cites_deepseek_reasoning",
    "thesaurus_kimi_reasoning",
    "thesaurus_deepseek_reasoning",
    "reference_kimi_reasoning",
    "reference_deepseek_reasoning",
    "ask_kimi_reasoning",
    "ask_deepseek_reasoning",
    "shorten_kimi_reasoning",
    "shorten_deepseek_reasoning",
    "introduction_kimi_reasoning",
    "introduction_deepseek_reasoning",
    "conclusion_kimi_reasoning",
    "conclusion_deepseek_reasoning",
    "concl_no_issues_kimi_reasoning",
    "concl_no_issues_deepseek_reasoning",
    "topic_sentence_kimi_reasoning",
    "topic_sentence_deepseek_reasoning",
    "concl_section_kimi_reasoning",
    "concl_section_deepseek_reasoning",
    "translate_kimi_reasoning",
    "translate_deepseek_reasoning",
)

_UNO_BOOTSTRAPPED = False
_UNO_IMPORT_ERROR: str | None = None
_UNO_IMPORT_SOURCE = "uninitialized"
_UNO_IMPORT_PATH: Path | None = None
_UNO_ATTEMPTED_PATHS: list[Path] = []


def _coerce_bool_config(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "on"}:
            return True
        if normalized in {"0", "false", "no", "off"}:
            return False
    if isinstance(value, (int, float)):
        return bool(value)
    return default


def _model_looks_kimi(model_id: str) -> bool:
    normalized = (model_id or "").strip().lower()
    return "kimi" in normalized or "moonshot" in normalized


def _model_looks_deepseek(model_id: str) -> bool:
    normalized = (model_id or "").strip().lower()
    return "deepseek" in normalized


def _model_looks_minimax(model_id: str) -> bool:
    normalized = (model_id or "").strip().lower()
    return "minimax" in normalized


def _api_url_looks_fireworks(api_url: str) -> bool:
    normalized = (api_url or "").strip().lower()
    return "fireworks.ai" in normalized


def _apply_disable_reasoning_to_body(
    body: dict[str, Any],
    *,
    model_id: str,
    disable_reasoning: bool,
) -> None:
    if not disable_reasoning:
        return
    if _model_looks_deepseek(model_id):
        body["reasoning_effort"] = "none"
    elif _model_looks_kimi(model_id):
        body["thinking"] = {"type": "disabled"}
    elif _model_looks_minimax(model_id):
        body["reasoning_effort"] = "low"
    else:
        body["reasoning_effort"] = "none"


def _resolve_tavily_cli_path() -> str | None:
    found = shutil.which("tvly")
    if found:
        return found
    candidates = (
        Path.home() / ".local" / "bin" / "tvly",
        Path.home() / ".local" / "share" / "uv" / "tools" / "tavily-cli" / "bin" / "tvly",
    )
    for candidate in candidates:
        try:
            if candidate.is_file() and os.access(candidate, os.X_OK):
                return str(candidate)
        except OSError:
            continue
    return None


def _read_config() -> dict[str, Any]:
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _write_config(data: dict[str, Any]) -> None:
    try:
        for key in OBSOLETE_REASONING_CONFIG_KEYS:
            data.pop(key, None)
        CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass


def load_shared_style_rules() -> str:
    raw = _read_config()
    if CONFIG_KEY_SHARED_STYLE_RULES in raw:
        return str(raw.get(CONFIG_KEY_SHARED_STYLE_RULES, "") or "").strip()
    return DEFAULT_SHARED_STYLE_RULES


def save_shared_style_rules(shared_style_rules: str) -> None:
    data = _read_config()
    data[CONFIG_KEY_SHARED_STYLE_RULES] = str(shared_style_rules or "").strip()
    _write_config(data)


def _expand_shared_prompt_parts(prompt_text: str) -> str:
    text = str(prompt_text or "").strip()
    if not text or STYLE_RULES_TOKEN not in text:
        return text
    expanded = text.replace(STYLE_RULES_TOKEN, load_shared_style_rules())
    return re.sub(r"\n{3,}", "\n\n", expanded).strip()


def _replace_prompt_block_with_style_token(prompt_text: str, start_marker: str, end_marker: str) -> str:
    text = str(prompt_text or "").strip()
    if not text or STYLE_RULES_TOKEN in text:
        return text
    start = text.find(start_marker)
    end = text.find(end_marker)
    if start < 0 or end <= start:
        return text
    prefix = text[:start].rstrip()
    suffix = text[end:].lstrip("\n")
    return "\n\n".join(part for part in (prefix, STYLE_RULES_TOKEN, suffix) if part).strip()


def _insert_style_rules_token_before_marker(prompt_text: str, marker: str) -> str:
    text = str(prompt_text or "").strip()
    if not text or STYLE_RULES_TOKEN in text:
        return text
    marker_index = text.find(marker)
    if marker_index < 0:
        return text
    prefix = text[:marker_index].rstrip()
    suffix = text[marker_index:].lstrip("\n")
    return "\n\n".join(part for part in (prefix, STYLE_RULES_TOKEN, suffix) if part).strip()


def _migrate_shared_style_rules_prompts() -> None:
    data = _read_config()
    if not data:
        return

    changed = False
    if CONFIG_KEY_SHARED_STYLE_RULES not in data:
        data[CONFIG_KEY_SHARED_STYLE_RULES] = DEFAULT_SHARED_STYLE_RULES
        changed = True

    replacements: dict[str, tuple[str, str]] = {
        CONFIG_KEY_PROOFREAD_PROMPT: ("## STYLE RULES", "## SOURCE text"),
        CONFIG_KEY_SPELLING_PROMPT: ("## STYLE RULES", "## OUTPUT"),
        CONFIG_KEY_IMPROVE1_PROMPT: ("## STYLE RULES", "## OUTPUT"),
    }
    insertions: dict[str, str] = {
        CONFIG_KEY_IMPROVE2_PROMPT: "## RULES",
        CONFIG_KEY_SHORTEN_PROMPT: "## RULES",
        CONFIG_KEY_INTRO_PROMPT: "Below are ten sample",
        CONFIG_KEY_INTRO_REPLY_PROMPT: "Below are ten sample",
        CONFIG_KEY_CONCLUSION_PROMPT: "Below are ten sample",
        CONFIG_KEY_CONCL_NO_ISSUES_PROMPT: "Below are ten sample",
        CONFIG_KEY_TOPIC_SENTENCE_PROMPT: "## RULES",
        CONFIG_KEY_CONCL_SECTION_PROMPT: "## RULES",
    }

    for key, (start_marker, end_marker) in replacements.items():
        prompt = data.get(key)
        if not isinstance(prompt, str) or not prompt.strip():
            continue
        updated = _replace_prompt_block_with_style_token(prompt, start_marker, end_marker)
        if updated != prompt:
            data[key] = updated
            changed = True

    for key, marker in insertions.items():
        prompt = data.get(key)
        if not isinstance(prompt, str) or not prompt.strip():
            continue
        updated = _insert_style_rules_token_before_marker(prompt, marker)
        if updated != prompt:
            data[key] = updated
            changed = True

    improve2_prompt = str(data.get(CONFIG_KEY_IMPROVE2_PROMPT, "") or "").strip()
    if (
        not improve2_prompt
        or improve2_prompt == DEFAULT_IMPROVE_PROMPT
        or "Rewrite the SOURCE text to be more clear, concise, and easier to read without changing its meaning." in improve2_prompt
        or "Rewrite the SOURCE text to be more clear, more concise, and easier to read without changing its meaning." in improve2_prompt
        or improve2_prompt.startswith("Improve the following text for clarity and precision while preserving meaning.")
    ) and improve2_prompt != DEFAULT_IMPROVE2_PROMPT:
        data[CONFIG_KEY_IMPROVE2_PROMPT] = DEFAULT_IMPROVE2_PROMPT
        changed = True

    if changed:
        _write_config(data)


def _normal_libreoffice_profile_path() -> Path:
    root = DEFAULT_NORMAL_LIBREOFFICE_PROFILE.expanduser().resolve(strict=False)
    user_dir = root / "user"
    if user_dir.exists():
        return root
    legacy_user_dir = Path.home() / ".config" / "libreoffice"
    if (legacy_user_dir / "user").exists():
        return legacy_user_dir.expanduser().resolve(strict=False)
    return root


def load_libreoffice_python_path() -> Path | None:
    raw = _read_config()
    path = raw.get(CONFIG_KEY_LIBREOFFICE_PYTHON_PATH)
    if isinstance(path, str) and path.strip():
        return Path(path).expanduser().resolve(strict=False)
    return DEFAULT_LIBREOFFICE_PYTHON_PATH


def save_libreoffice_python_path(path: Path | None) -> None:
    data = _read_config()
    if path:
        data[CONFIG_KEY_LIBREOFFICE_PYTHON_PATH] = str(path.expanduser().resolve(strict=False))
    else:
        data.pop(CONFIG_KEY_LIBREOFFICE_PYTHON_PATH, None)
    _write_config(data)


def load_concordance_file_path() -> Path | None:
    raw = _read_config()
    path = raw.get(CONFIG_KEY_CONCORDANCE_FILE_PATH)
    if isinstance(path, str) and path.strip():
        return Path(path).expanduser().resolve(strict=False)
    return None


def save_concordance_file_path(path: Path | None) -> None:
    data = _read_config()
    if path:
        data[CONFIG_KEY_CONCORDANCE_FILE_PATH] = str(path.expanduser().resolve(strict=False))
    else:
        data.pop(CONFIG_KEY_CONCORDANCE_FILE_PATH, None)
    _write_config(data)


def _candidate_uno_paths(configured_path: Path | None = None) -> list[Path]:
    seen: set[Path] = set()
    candidates: list[Path] = []

    def _add(path: Path | None) -> None:
        if not path:
            return
        resolved = path.expanduser().resolve(strict=False)
        if resolved in seen:
            return
        seen.add(resolved)
        candidates.append(resolved)

    _add(configured_path)

    env_pythonpath = os.environ.get("PYTHONPATH", "")
    for raw in env_pythonpath.split(os.pathsep):
        if raw.strip():
            _add(Path(raw.strip()))

    soffice_dir = SOFFICE_PATH.parent
    if SOFFICE_PATH.exists():
        _add(soffice_dir)

    for candidate in LIBREOFFICE_PYTHON_CANDIDATES:
        _add(candidate)

    return candidates


def _import_uno_from_candidates(configured_path: Path | None = None, *, force_retry: bool = False) -> bool:
    global _UNO_BOOTSTRAPPED, _UNO_IMPORT_ERROR, _UNO_IMPORT_SOURCE, _UNO_IMPORT_PATH, _UNO_ATTEMPTED_PATHS
    global uno, PropertyValue, NoConnectException, XModel, ControlCharacter, XTextDocument

    if _UNO_BOOTSTRAPPED and not force_retry and uno is not None:
        return True

    _UNO_BOOTSTRAPPED = True
    _UNO_IMPORT_ERROR = None
    _UNO_IMPORT_SOURCE = "direct"
    _UNO_IMPORT_PATH = None
    _UNO_ATTEMPTED_PATHS = []

    try:
        importlib.invalidate_caches()
        uno = importlib.import_module("uno")  # type: ignore[assignment]
        from com.sun.star.beans import PropertyValue as _PropertyValue  # type: ignore
        from com.sun.star.connection import NoConnectException as _NoConnectException  # type: ignore
        from com.sun.star.frame import XModel as _XModel  # type: ignore
        from com.sun.star.text import ControlCharacter as _ControlCharacter, XTextDocument as _XTextDocument  # type: ignore

        PropertyValue = _PropertyValue  # type: ignore[assignment]
        NoConnectException = _NoConnectException  # type: ignore[assignment]
        XModel = _XModel  # type: ignore[assignment]
        ControlCharacter = _ControlCharacter  # type: ignore[assignment]
        XTextDocument = _XTextDocument  # type: ignore[assignment]
        return True
    except Exception as exc:  # noqa: BLE001
        _UNO_IMPORT_ERROR = str(exc) or exc.__class__.__name__

    for candidate in _candidate_uno_paths(configured_path):
        _UNO_ATTEMPTED_PATHS.append(candidate)
        uno_py = candidate / "uno.py"
        uno_pkg = candidate / "uno"
        if not uno_py.exists() and not uno_pkg.exists():
            continue
        candidate_str = str(candidate)
        if candidate_str not in sys.path:
            sys.path.append(candidate_str)
        try:
            importlib.invalidate_caches()
            sys.modules.pop("uno", None)
            uno = importlib.import_module("uno")  # type: ignore[assignment]
            from com.sun.star.beans import PropertyValue as _PropertyValue  # type: ignore
            from com.sun.star.connection import NoConnectException as _NoConnectException  # type: ignore
            from com.sun.star.frame import XModel as _XModel  # type: ignore
            from com.sun.star.text import ControlCharacter as _ControlCharacter, XTextDocument as _XTextDocument  # type: ignore

            PropertyValue = _PropertyValue  # type: ignore[assignment]
            NoConnectException = _NoConnectException  # type: ignore[assignment]
            XModel = _XModel  # type: ignore[assignment]
            ControlCharacter = _ControlCharacter  # type: ignore[assignment]
            XTextDocument = _XTextDocument  # type: ignore[assignment]
            _UNO_IMPORT_SOURCE = "configured" if configured_path and candidate == configured_path.resolve(strict=False) else "auto"
            _UNO_IMPORT_PATH = candidate
            _UNO_IMPORT_ERROR = None
            return True
        except Exception as exc:  # noqa: BLE001
            _UNO_IMPORT_ERROR = str(exc) or exc.__class__.__name__

    uno = None  # type: ignore[assignment]
    PropertyValue = None  # type: ignore[assignment]
    NoConnectException = Exception  # type: ignore[assignment]
    XModel = None  # type: ignore[assignment]
    ControlCharacter = None  # type: ignore[assignment]
    XTextDocument = None  # type: ignore[assignment]
    return False


def _uno_status_message(configured_path: Path | None = None) -> str:
    if uno is not None:
        if _UNO_IMPORT_SOURCE == "configured" and _UNO_IMPORT_PATH:
            return f"python-uno loaded from configured path: {_UNO_IMPORT_PATH}"
        if _UNO_IMPORT_SOURCE == "auto" and _UNO_IMPORT_PATH:
            return f"python-uno auto-detected at {_UNO_IMPORT_PATH}"
        return "python-uno available."

    if configured_path:
        return f"python-uno unavailable. Configured path not usable: {configured_path}"
    if _UNO_ATTEMPTED_PATHS:
        return "python-uno unavailable. Set a LibreOffice Python path in Settings."
    return "python-uno unavailable."


def _strip_odt_page_style_metadata(xml_text: str) -> tuple[str, int]:
    stripped, attr_count = ODT_PAGE_STYLE_ATTR_RE.subn("", xml_text)

    def _strip_master_page_next_style(match: re.Match[str]) -> str:
        return ODT_MASTER_PAGE_NEXT_STYLE_RE.sub("", match.group(0))

    stripped, master_count = ODT_MASTER_PAGE_TAG_RE.subn(_strip_master_page_next_style, stripped)
    return stripped, attr_count + master_count


def _action_command(
    action_name: str,
    param: str | None = None,
    object_path: str = ACTION_OBJECT_PATH,
) -> str:
    params = "[]" if param is None else f"[{param}]"
    return shlex.join(
        [
            "gdbus",
            "call",
            "--session",
            "--dest",
            APP_ID,
            "--object-path",
            object_path,
            "--method",
            "org.gtk.Actions.Activate",
            action_name,
            params,
            "{}",
        ]
    )


def _format_action_param(variant: GLib.Variant) -> str:
    printed = variant.print_(True)
    if variant.is_of_type(GLib.VariantType.new("i")):
        return f"<int32 {printed}>"
    return f"<{printed}>"


@dataclass
class ModelProfile:
    key: str
    nickname: str
    abbreviation: str
    api_url: str
    model_id: str
    api_key: str
    disable_reasoning: bool
    priority_service_tier: bool

    def display_name(self) -> str:
        return self.nickname.strip() or _default_profile_nickname(self.key)

    def is_configured(self) -> bool:
        return bool(self.api_url.strip() and self.api_key.strip())


@dataclass
class ProofreadSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool
    suggestions_json_path: Path | None

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_PROMPT)
        )


@dataclass
class SpellingStyleSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_SPELLINGSTYLE_PROMPT)
        )


@dataclass
class CitationValidatorSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT)
        )


@dataclass
class Improve1Settings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_IMPROVE1_PROMPT)
        )


@dataclass
class Improve2Settings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_IMPROVE2_PROMPT)
        )


@dataclass
class CombineCitesSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_COMBINE_CITES_PROMPT)
        )


@dataclass
class ThesaurusSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_THESAURUS_PROMPT)
        )


@dataclass
class ReferenceSettings:
    api_url: str
    model_id: str
    api_key: str
    tavily_api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (
                self.api_url,
                self.api_key,
                self.tavily_api_key,
                self.prompt or DEFAULT_REFERENCE_PROMPT,
            )
        )


@dataclass
class AskSettings:
    api_url: str
    model_id: str
    api_key: str
    tavily_api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (
                self.api_url,
                self.api_key,
                self.tavily_api_key,
                self.prompt or DEFAULT_ASK_PROMPT,
            )
        )


@dataclass
class ShortenSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_SHORTEN_PROMPT)
        )


@dataclass
class IntroductionSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_INTRO_PROMPT)
        )


@dataclass
class ConclusionSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_CONCLUSION_PROMPT)
        )


@dataclass
class ConclNoIssuesSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_CONCL_NO_ISSUES_PROMPT)
        )


@dataclass
class TopicSentenceSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_TOPIC_SENTENCE_PROMPT)
        )


@dataclass
class ConclSectionSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_CONCL_SECTION_PROMPT)
        )


@dataclass
class TranslateSettings:
    api_url: str
    model_id: str
    api_key: str
    prompt: str
    disable_reasoning: bool

    def is_configured(self) -> bool:
        return all(
            value.strip()
            for value in (self.api_url, self.api_key, self.prompt or DEFAULT_TRANSLATE_PROMPT)
        )


@dataclass
class RegenerateContext:
    action_key: str
    command_title: str
    source_text: str
    insert_mode: str


@dataclass
class MultiDraftChoice:
    key: str
    label: str
    profile: ModelProfile
    buffer: Gtk.TextBuffer
    text_view: Gtk.TextView
    status_label: Gtk.Label
    insert_button: Gtk.Button
    variant: str | None = None
    source_text: str = ""
    text: str = ""
    deletion_marker_tag: Gtk.TextTag | None = None
    complete: bool = False
    failed: bool = False


@dataclass(frozen=True)
class MultiDraftRequest:
    key: str
    label: str
    profile: ModelProfile
    payload_builder: Callable[[str, ModelProfile], dict[str, Any]]
    prompt_text: str
    default_prompt: str
    request_title: str
    variant: str | None = None
    stacked: bool = False


@dataclass
class TextDraftExternalAction:
    enabled: bool
    label: str
    command: list[str]
    cwd: Path | None
    env: dict[str, str]
    icon_name: str = DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON
    tooltip: str = ""
    success_message: str = ""

    def is_configured(self) -> bool:
        return self.enabled and bool(self.command)


@dataclass
class TextDraftExternalActionEditorWidgets:
    enabled_row: Adw.SwitchRow
    label_row: Adw.EntryRow
    icon_row: Adw.EntryRow
    tooltip_row: Adw.EntryRow
    success_message_row: Adw.EntryRow
    command_row: Adw.EntryRow
    cwd_row: Adw.EntryRow
    env_buffer: Gtk.TextBuffer


@dataclass
class PromptEditorWidgets:
    api_url_row: Adw.EntryRow
    model_row: Adw.EntryRow
    api_key_row: Adw.PasswordEntryRow
    tavily_api_key_row: Adw.PasswordEntryRow | None
    disable_reasoning_row: Adw.SwitchRow
    prompt_buffer: Gtk.TextBuffer
    ask_prompt_buffer: Gtk.TextBuffer | None
    default_profile_dropdowns: dict[str, Gtk.DropDown] | None = None


@dataclass
class ModelProfileEditorWidgets:
    nickname_row: Adw.EntryRow
    abbreviation_row: Adw.EntryRow
    api_url_row: Adw.EntryRow
    model_row: Adw.EntryRow
    api_key_row: Adw.PasswordEntryRow
    disable_reasoning_row: Adw.SwitchRow
    priority_service_tier_row: Adw.SwitchRow


@dataclass
class WordSubstitution:
    original: str
    replacement: str


@dataclass
class TavilySearchSource:
    title: str
    url: str
    excerpt: str
    domain: str


@dataclass
class TavilySearchBundle:
    query: str
    sources: list[TavilySearchSource]
    notice: str | None = None


@dataclass
class PrefixSettings:
    rt_prefix: str
    ct_prefix: str
    substitutions: list[WordSubstitution]


@dataclass(frozen=True)
class QuickActionDefinition:
    key: str
    label: str
    title: str
    action_name: str
    description: str
    supports_profiles: bool = False


EDITOR_QUICK_ACTIONS = (
    QuickActionDefinition(
        key="improve-generated",
        label="Improve Generated",
        title="Improve Generated",
        action_name="improve-generated",
        description="Rewrite the latest SpellingStyle output using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="rephrase-generated",
        label="Rephrase Generated",
        title="Rephrase Generated",
        action_name="rephrase-generated",
        description="Rephrase the latest SpellingStyle output using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="generated-choices",
        label="Generated Choices",
        title="Generated Choices",
        action_name="generated-choices",
        description="Compare Improve, Rephrase 1, and Rephrase 2 outputs for the latest SpellingStyle output.",
    ),
    QuickActionDefinition(
        key="selected-choices",
        label="Selected Choices",
        title="Selected Choices",
        action_name="selected-choices",
        description="Compare Improve, Rephrase 1, and Rephrase 2 outputs for selected Writer text.",
    ),
    QuickActionDefinition(
        key="improve-selected",
        label="Improve Selected",
        title="Improve Selected",
        action_name="improve-selected",
        description="Rewrite selected text in Writer using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="shorten",
        label="Shorten Selected",
        title="Shorten Selected",
        action_name="transform-shorten",
        description="Shorten selected text while preserving meaning.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="topic-sentence",
        label="Topic Sentence",
        title="Topic Sentence",
        action_name="transform-topic-sentence",
        description="Generate a topic sentence from the current paragraph.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="concl-section",
        label="Section Conclusion",
        title="Section Conclusion",
        action_name="transform-concl-section",
        description="Write a concise conclusion for the current section.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="concl-section-choices",
        label="Section Conclusion Choices",
        title="Section Conclusion Choices",
        action_name="transform-concl-section-choices",
        description="Compare section conclusions using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="add-case",
        label="Add Case",
        title="Add Case",
        action_name="add-case",
        description="Add the selected case citation to concordance and AutoText.",
    ),
    QuickActionDefinition(
        key="import-socf",
        label="Import SOCF",
        title="Import SOCF",
        action_name="import-socf",
        description="Insert a SOCF ODT at the cursor without importing page-number reset metadata.",
    ),
    QuickActionDefinition(
        key="intro",
        label="Introduction",
        title="Introduction",
        action_name="transform-introduction",
        description="Write an introduction from the current argument section.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="intro-choices",
        label="Introduction Choices",
        title="Introduction Choices",
        action_name="transform-introduction-choices",
        description="Compare introductions using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="intro-reply",
        label="Introduction Reply",
        title="Introduction Reply",
        action_name="transform-introduction-reply",
        description="Write an introduction for a reply brief from the current argument section.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="intro-reply-choices",
        label="Introduction Reply Choices",
        title="Introduction Reply Choices",
        action_name="transform-introduction-reply-choices",
        description="Compare reply introductions using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="conclusion",
        label="Conclusion",
        title="Conclusion",
        action_name="transform-conclusion",
        description="Write a conclusion from the current argument section.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="conclusion-choices",
        label="Conclusion Choices",
        title="Conclusion Choices",
        action_name="transform-conclusion-choices",
        description="Compare conclusions using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="concl-no-issues",
        label="Conclusion No Issues",
        title="Conclusion No Issues",
        action_name="transform-concl-no-issues",
        description="Write a conclusion from the current issues-considered section.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="concl-no-issues-choices",
        label="Conclusion No Issues Choices",
        title="Conclusion No Issues Choices",
        action_name="transform-concl-no-issues-choices",
        description="Compare no-issues conclusions using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="quotes",
        label="Quotes",
        title="Wrap Selection in Quotes",
        action_name="transform-wrap-quotes",
        description="Wrap selected text in curly quotes.",
    ),
    QuickActionDefinition(
        key="translate",
        label="Translate",
        title="Translate Document",
        action_name="transform-translate",
        description="Translate the full Writer document into Spanish.",
    ),
)
EDITOR_QUICK_ACTION_BY_KEY = {definition.key: definition for definition in EDITOR_QUICK_ACTIONS}
EDITOR_QUICK_ACTION_BY_ACTION_NAME = {
    definition.action_name: definition for definition in EDITOR_QUICK_ACTIONS
}
EDITOR_PROFILE_ACTION_KEYS = tuple(
    definition.key for definition in EDITOR_QUICK_ACTIONS if definition.supports_profiles
)
EDITOR_SINGLE_DRAFT_ACTION_KEYS = frozenset(
    {
        "intro",
        "intro-reply",
        "conclusion",
        "concl-no-issues",
    }
)
EDITOR_DRAFT_CHOICE_ACTION_KEYS = (
    "generated-choices",
    "selected-choices",
    "intro-choices",
    "intro-reply-choices",
    "conclusion-choices",
    "concl-no-issues-choices",
    "concl-section-choices",
)
EDITOR_DEDICATED_QUICK_ACTION_KEYS = frozenset(EDITOR_DRAFT_CHOICE_ACTION_KEYS)
EDITOR_TOOLBAR_EXCLUDED_ACTION_KEYS = EDITOR_SINGLE_DRAFT_ACTION_KEYS | EDITOR_DEDICATED_QUICK_ACTION_KEYS
DEFAULT_EDITOR_PINNED_ACTION_IDS = (
    "shorten",
    "topic-sentence",
    "concl-section",
)
DEFAULT_TEXT_DRAFT_PINNED_ACTION_IDS = (
    "text-draft-improve-generated",
    "text-draft-rephrase-generated",
    "text-draft-improve-selected",
    "text-draft-keep-original",
)


def _editor_command_items() -> list[tuple[str, str, str | None, str, bool]]:
    commands: list[tuple[str, str, str | None, str, bool]] = [
        ("Launch Writer", "launch-writer", None, "Open a Writer document via UNO.", False),
        ("Direct Input", "direct-input", None, "Insert the source file text into Writer.", False),
        ("Input RT", "input-rt", None, "Insert a formatted RT citation from the source file.", False),
        ("Input CT", "input-ct", None, "Insert a formatted CT citation from the source file.", False),
        ("Speech Find", "speech-find", None, "Find the source file speech text in Writer.", False),
        (
            "Direct Input No Trailing Space",
            "direct-input-no-trailing-space",
            None,
            "Insert the source file text into Writer without trailing spaces.",
            False,
        ),
        ("Combine Cites", "combine-cites", None, "Combine the first run of adjacent citations.", False),
        ("SpellingStyle", "spellingstyle", None, "Stream model output into Writer.", False),
        ("Keep Original", "keep-original", None, "Restore the last SpellingStyle output.", False),
        ("Reference Lookup", "reference-lookup", None, "Look up a definition for the selected text.", False),
        ("Focus Ask Field", "focus-ask", None, "Focus the Ask question field in the Editor view.", False),
    ]
    commands.extend(
        (
            definition.title,
            definition.action_name,
            None,
            definition.description,
            definition.supports_profiles,
        )
        for definition in EDITOR_QUICK_ACTIONS
        if definition.key not in EDITOR_SINGLE_DRAFT_ACTION_KEYS
    )
    return commands


TEXT_DRAFT_QUICK_ACTIONS = (
    QuickActionDefinition(
        key="text-draft-improve-generated",
        label="Improve Generated",
        title="Improve Generated",
        action_name="text-draft-improve-generated",
        description="Rewrite the latest Text Draft Original Output using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="text-draft-generated-choices",
        label="Generated Choices",
        title="Generated Choices",
        action_name="text-draft-generated-choices",
        description="Compare improve and rephrase outputs for the latest Text Draft Original Output using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="text-draft-rephrase-generated",
        label="Rephrase Generated",
        title="Rephrase Generated",
        action_name="text-draft-rephrase-generated",
        description="Rephrase the latest Text Draft Original Output using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="text-draft-improve-selected",
        label="Improve Selected",
        title="Improve Selected",
        action_name="text-draft-improve-selected",
        description="Rewrite selected text in the Text Draft buffer using the selected model profile.",
        supports_profiles=True,
    ),
    QuickActionDefinition(
        key="text-draft-selected-choices",
        label="Selected Choices",
        title="Selected Choices",
        action_name="text-draft-selected-choices",
        description="Compare improve and rephrase outputs for selected Text Draft text using the Choices profiles.",
    ),
    QuickActionDefinition(
        key="text-draft-keep-original",
        label="Keep Original",
        title="Keep Original",
        action_name="text-draft-keep-original",
        description="Restore the last Text Draft generated range from Original Output.",
    ),
    QuickActionDefinition(
        key="text-draft-wrap-quotes",
        label="Quotes",
        title="Wrap Quotes",
        action_name="text-draft-wrap-quotes",
        description="Wrap the selected Text Draft text in curly quotes.",
    ),
)
TEXT_DRAFT_QUICK_ACTION_BY_KEY = {definition.key: definition for definition in TEXT_DRAFT_QUICK_ACTIONS}
TEXT_DRAFT_QUICK_ACTION_BY_ACTION_NAME = {
    definition.action_name: definition for definition in TEXT_DRAFT_QUICK_ACTIONS
}


def _text_draft_command_items() -> list[tuple[str, str, str | None, str, bool]]:
    commands = [
        (
            "SpellingStyle",
            "text-draft-spellingstyle",
            None,
            "Stream model output into the local Text Draft buffer.",
            False,
        ),
        *[
        (
            definition.title,
            definition.action_name,
            None,
            definition.description,
            definition.supports_profiles,
        )
        for definition in TEXT_DRAFT_QUICK_ACTIONS
        ],
    ]
    commands.append(
        ("Copy Draft", "text-draft-copy", None, "Copy the full Text Draft buffer to the clipboard.", False)
    )
    commands.append(
        (
            "External Draft Action",
            "text-draft-external-action",
            None,
            "Run the configured external command with the full Text Draft buffer.",
            False,
        )
    )
    return commands


TEXT_DRAFT_VIEW_ACTION_NAMES = frozenset(
    {
        "text-draft-spellingstyle",
        "text-draft-copy",
        "text-draft-external-action",
        *TEXT_DRAFT_QUICK_ACTION_BY_ACTION_NAME.keys(),
    }
)
EDITOR_VIEW_ACTION_NAMES = frozenset(
    {
        "launch-writer",
        "open-last-odt",
        "choose-source-file",
        "editor-commands",
        "direct-input",
        "direct-input-no-trailing-space",
        "input-rt",
        "input-ct",
        "speech-find",
        "combine-cites",
        "spellingstyle",
        "keep-original",
        "reference-lookup",
        "focus-ask",
        "paste-clean-italics",
        "improve",
        "improve1",
        "improve2",
        *EDITOR_QUICK_ACTION_BY_ACTION_NAME.keys(),
    }
)


def _main_view_name_for_action(action_name: str) -> str | None:
    if action_name in TEXT_DRAFT_VIEW_ACTION_NAMES:
        return "text-draft"
    if action_name in EDITOR_VIEW_ACTION_NAMES:
        return "editor"
    return None


def _profile_default_lookup_key_for_action_name(action_name: str) -> str | None:
    definition = EDITOR_QUICK_ACTION_BY_ACTION_NAME.get(action_name)
    if definition is not None:
        return definition.key
    text_draft_definition = TEXT_DRAFT_QUICK_ACTION_BY_ACTION_NAME.get(action_name)
    if text_draft_definition is not None:
        if text_draft_definition.key == "text-draft-improve-generated":
            return "improve-generated"
        if text_draft_definition.key == "text-draft-rephrase-generated":
            return "rephrase-generated"
        if text_draft_definition.key == "text-draft-improve-selected":
            return "improve-selected"
    if action_name == "improve":
        return "improve-generated"
    if action_name == "improve-selected":
        return "improve-selected"
    return None


def _profile_default_source_keys(action_key: str) -> tuple[str, ...]:
    if action_key == "improve-generated":
        return ("improve-generated", "improve")
    if action_key == "choices-improve":
        return ("choices-improve", "choices-profile-1", "improve-generated", "improve")
    if action_key == "choices-rephrase-1":
        return ("choices-rephrase-1", "choices-profile-1", "rephrase-generated")
    if action_key == "choices-rephrase-2":
        return ("choices-rephrase-2", "choices-profile-2", "rephrase-generated")
    return (action_key,)


def _normalize_editor_quick_action_key(action_key: str) -> str:
    if action_key == "improve":
        return "improve-generated"
    return action_key


def _credential_signature(api_url: str, model_id: str, api_key: str) -> tuple[str, str, str] | None:
    cleaned_api_url = api_url.strip()
    cleaned_model_id = model_id.strip()
    cleaned_api_key = api_key.strip()
    if not cleaned_api_url or not cleaned_api_key:
        return None
    return cleaned_api_url, cleaned_model_id, cleaned_api_key


def _default_profile_nickname(profile_key: str) -> str:
    fallback = DEFAULT_MODEL_PROFILE_NICKNAMES.get(profile_key)
    if fallback:
        return fallback
    match = re.fullmatch(r"profile(\d+)", profile_key or "")
    if match:
        return f"Profile {match.group(1)}"
    return profile_key.title()


def _match_profile_key_for_settings(settings: Any, profiles: list[ModelProfile]) -> str | None:
    signature = _credential_signature(
        str(getattr(settings, "api_url", "") or ""),
        str(getattr(settings, "model_id", "") or ""),
        str(getattr(settings, "api_key", "") or ""),
    )
    if signature is None:
        return None
    for profile in profiles:
        if _credential_signature(profile.api_url, profile.model_id, profile.api_key) == signature:
            return profile.key
    return None


def _sanitize_model_profile(raw: Any, key: str, fallback_nickname: str) -> ModelProfile:
    data = raw if isinstance(raw, dict) else {}
    nickname = str(data.get("nickname", fallback_nickname) or "").strip() or fallback_nickname
    return ModelProfile(
        key=key,
        nickname=nickname,
        abbreviation=str(data.get("abbreviation", "") or "").strip(),
        api_url=str(data.get("api_url", "") or "").strip(),
        model_id=str(data.get("model_id", "") or "").strip(),
        api_key=str(data.get("api_key", "") or "").strip(),
        disable_reasoning=_coerce_bool_config(data.get("disable_reasoning"), False),
        priority_service_tier=_coerce_bool_config(data.get("priority_service_tier"), False),
    )


def _legacy_profile_from_improve(prefix: str, nickname: str) -> ModelProfile:
    raw = _read_config()
    return ModelProfile(
        key="",
        nickname=nickname,
        abbreviation="",
        api_url=str(raw.get(f"{prefix}_api_url", "") or "").strip(),
        model_id=str(raw.get(f"{prefix}_model_id", "") or "").strip(),
        api_key=str(raw.get(f"{prefix}_api_key", "") or "").strip(),
        disable_reasoning=_coerce_bool_config(raw.get(f"{prefix}_disable_reasoning"), False),
        priority_service_tier=False,
    )


def _legacy_default_profile_key_for_command(command_key: str, raw: dict[str, Any] | None = None) -> str | None:
    data = raw if raw is not None else _read_config()
    for container_key in (CONFIG_KEY_COMMAND_DEFAULT_PROFILES, CONFIG_KEY_EDITOR_COMMAND_DEFAULT_PROFILES):
        container = data.get(container_key)
        if not isinstance(container, dict):
            continue
        candidate = str(container.get(command_key, "") or "").strip()
        if re.fullmatch(r"profile\d+", candidate):
            return candidate
    return None


def _raw_model_profile_by_key(raw_profiles: Any, profile_key: str) -> ModelProfile | None:
    if not isinstance(raw_profiles, list):
        return None
    match = re.fullmatch(r"profile(\d+)", profile_key or "")
    if match is None:
        return None
    index = int(match.group(1)) - 1
    if index < 0 or index >= len(raw_profiles):
        return None
    return _sanitize_model_profile(raw_profiles[index], profile_key, _default_profile_nickname(profile_key))


def load_model_profiles() -> list[ModelProfile]:
    raw = _read_config()
    raw_profiles = raw.get(CONFIG_KEY_MODEL_PROFILES)
    if isinstance(raw_profiles, list) and raw_profiles:
        profiles: list[ModelProfile] = []
        for index, key in enumerate(MODEL_PROFILE_IDS):
            fallback = DEFAULT_MODEL_PROFILE_NICKNAMES[key]
            entry = raw_profiles[index] if index < len(raw_profiles) else {}
            profiles.append(_sanitize_model_profile(entry, key, fallback))
        return profiles

    legacy_first = _legacy_profile_from_improve("improve1", "Improve 1")
    legacy_first.key = "profile1"
    legacy_second = _legacy_profile_from_improve("improve2", "Improve 2")
    legacy_second.key = "profile2"
    return [
        legacy_first,
        legacy_second,
        ModelProfile(
            key="profile3",
            nickname=DEFAULT_MODEL_PROFILE_NICKNAMES["profile3"],
            abbreviation="",
            api_url="",
            model_id="",
            api_key="",
            disable_reasoning=False,
            priority_service_tier=False,
        ),
        ModelProfile(
            key="profile4",
            nickname=DEFAULT_MODEL_PROFILE_NICKNAMES["profile4"],
            abbreviation="",
            api_url="",
            model_id="",
            api_key="",
            disable_reasoning=False,
            priority_service_tier=False,
        ),
    ]


def save_model_profiles(profiles: list[ModelProfile]) -> None:
    data = _read_config()
    data[CONFIG_KEY_MODEL_PROFILES] = [
        {
            "nickname": profile.display_name(),
            "abbreviation": profile.abbreviation.strip(),
            "api_url": profile.api_url,
            "model_id": profile.model_id,
            "api_key": profile.api_key,
            "disable_reasoning": bool(profile.disable_reasoning),
            "priority_service_tier": bool(profile.priority_service_tier),
        }
        for profile in profiles[: len(MODEL_PROFILE_IDS)]
    ]
    _write_config(data)


def _sanitize_editor_action_profile_defaults(raw: Any) -> dict[str, str | None]:
    defaults: dict[str, str | None] = {}
    source = raw if isinstance(raw, dict) else {}
    for key in PROFILE_BACKED_COMMAND_KEYS:
        selected_key = None
        for source_key in _profile_default_source_keys(key):
            candidate = str(source.get(source_key, "") or "").strip()
            if candidate in MODEL_PROFILE_IDS:
                selected_key = candidate
                break
        defaults[key] = selected_key
    return defaults


def load_editor_action_profile_defaults(
    profiles: list[ModelProfile],
    settings_by_key: dict[str, Any],
) -> dict[str, str | None]:
    raw = _read_config()
    defaults = _sanitize_editor_action_profile_defaults(raw.get(CONFIG_KEY_COMMAND_DEFAULT_PROFILES))
    legacy_defaults = _sanitize_editor_action_profile_defaults(raw.get(CONFIG_KEY_EDITOR_COMMAND_DEFAULT_PROFILES))
    for key in PROFILE_BACKED_COMMAND_KEYS:
        if defaults[key] is None and legacy_defaults[key] is not None:
            defaults[key] = legacy_defaults[key]
        if defaults[key] is None:
            matched_profile_key = _match_profile_key_for_settings(settings_by_key.get(key), profiles)
            if matched_profile_key is not None:
                defaults[key] = matched_profile_key
    if defaults.get("rephrase-generated") is None and defaults.get("improve-generated") is not None:
        defaults["rephrase-generated"] = defaults["improve-generated"]
    return defaults


def save_editor_action_profile_defaults(defaults: dict[str, str | None]) -> None:
    data = _read_config()
    sanitized = _sanitize_editor_action_profile_defaults(defaults)
    data[CONFIG_KEY_COMMAND_DEFAULT_PROFILES] = {
        key: value
        for key, value in sanitized.items()
        if value in MODEL_PROFILE_IDS
    }
    _write_config(data)


def load_proofread_settings() -> ProofreadSettings:
    raw = _read_config()
    suggestions_path_raw = raw.get(CONFIG_KEY_PROOFREAD_SUGGESTIONS_JSON_PATH)
    if isinstance(suggestions_path_raw, str) and suggestions_path_raw.strip():
        suggestions_json_path = Path(suggestions_path_raw).expanduser().resolve(strict=False)
    else:
        suggestions_json_path = None
    settings = ProofreadSettings(
        api_url=str(raw.get(CONFIG_KEY_PROOFREAD_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_PROOFREAD_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_PROOFREAD_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_PROOFREAD_PROMPT, DEFAULT_PROMPT) or DEFAULT_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_PROOFREAD_DISABLE_REASONING), False),
        suggestions_json_path=suggestions_json_path,
    )
    if settings.api_url.strip() and settings.api_key.strip():
        return settings

    legacy_profile_key = _legacy_default_profile_key_for_command("proof", raw)
    legacy_profile = _raw_model_profile_by_key(raw.get(CONFIG_KEY_MODEL_PROFILES), legacy_profile_key or "")
    if legacy_profile is None or not legacy_profile.is_configured():
        return settings

    return ProofreadSettings(
        api_url=legacy_profile.api_url,
        model_id=legacy_profile.model_id,
        api_key=legacy_profile.api_key,
        prompt=settings.prompt or DEFAULT_PROMPT,
        disable_reasoning=legacy_profile.disable_reasoning,
        suggestions_json_path=settings.suggestions_json_path,
    )


def save_proofread_settings(settings: ProofreadSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_PROOFREAD_API_URL] = settings.api_url
    data[CONFIG_KEY_PROOFREAD_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_PROOFREAD_API_KEY] = settings.api_key
    data[CONFIG_KEY_PROOFREAD_PROMPT] = settings.prompt or DEFAULT_PROMPT
    data[CONFIG_KEY_PROOFREAD_DISABLE_REASONING] = bool(settings.disable_reasoning)
    data[CONFIG_KEY_PROOFREAD_SUGGESTIONS_JSON_PATH] = (
        str(settings.suggestions_json_path.expanduser().resolve(strict=False))
        if settings.suggestions_json_path
        else ""
    )
    _write_config(data)


def _sanitize_word_substitutions(raw: Any) -> list[WordSubstitution]:
    substitutions: list[WordSubstitution] = []
    if isinstance(raw, list):
        for item in raw[:MAX_WORD_SUBSTITUTIONS]:
            if not isinstance(item, dict):
                continue
            original = str(item.get("original", "") or "").strip()
            replacement = str(item.get("replacement", "") or "").strip()
            substitutions.append(WordSubstitution(original=original, replacement=replacement))
    while len(substitutions) < MAX_WORD_SUBSTITUTIONS:
        substitutions.append(WordSubstitution(original="", replacement=""))
    return substitutions


def load_prefix_settings() -> PrefixSettings:
    raw = _read_config()
    substitutions = _sanitize_word_substitutions(raw.get(CONFIG_KEY_WORD_SUBSTITUTIONS))
    return PrefixSettings(
        rt_prefix=str(raw.get(CONFIG_KEY_RT_PREFIX, DEFAULT_RT_PREFIX) or "").strip(),
        ct_prefix=str(raw.get(CONFIG_KEY_CT_PREFIX, DEFAULT_CT_PREFIX) or "").strip(),
        substitutions=substitutions,
    )


def save_prefix_settings(settings: PrefixSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_RT_PREFIX] = settings.rt_prefix
    data[CONFIG_KEY_CT_PREFIX] = settings.ct_prefix
    data[CONFIG_KEY_WORD_SUBSTITUTIONS] = [
        {"original": entry.original.strip(), "replacement": entry.replacement.strip()}
        for entry in settings.substitutions[:MAX_WORD_SUBSTITUTIONS]
    ]
    _write_config(data)


def load_spellingstyle_settings() -> SpellingStyleSettings:
    raw = _read_config()
    return SpellingStyleSettings(
        api_url=str(raw.get(CONFIG_KEY_SPELLING_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_SPELLING_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_SPELLING_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_SPELLING_PROMPT, DEFAULT_SPELLINGSTYLE_PROMPT) or DEFAULT_SPELLINGSTYLE_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_SPELLING_DISABLE_REASONING), False),
    )


def save_spellingstyle_settings(settings: SpellingStyleSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_SPELLING_API_URL] = settings.api_url
    data[CONFIG_KEY_SPELLING_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_SPELLING_API_KEY] = settings.api_key
    data[CONFIG_KEY_SPELLING_PROMPT] = settings.prompt or DEFAULT_SPELLINGSTYLE_PROMPT
    data[CONFIG_KEY_SPELLING_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_text_draft_spellingstyle_prompt() -> str:
    raw = _read_config()
    return str(
        raw.get(CONFIG_KEY_TEXT_DRAFT_SPELLING_PROMPT, DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT)
        or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
    ).strip()


def save_text_draft_spellingstyle_prompt(prompt: str) -> None:
    data = _read_config()
    data[CONFIG_KEY_TEXT_DRAFT_SPELLING_PROMPT] = prompt or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
    _write_config(data)


def load_citation_validator_settings() -> CitationValidatorSettings:
    raw = _read_config()
    has_saved_connection = any(
        str(raw.get(key, "") or "").strip()
        for key in (
            CONFIG_KEY_CITATION_VALIDATOR_API_URL,
            CONFIG_KEY_CITATION_VALIDATOR_MODEL_ID,
            CONFIG_KEY_CITATION_VALIDATOR_API_KEY,
        )
    )
    default_api_url = raw.get(CONFIG_KEY_SPELLING_API_URL, "") if not has_saved_connection else ""
    default_model_id = raw.get(CONFIG_KEY_SPELLING_MODEL_ID, "") if not has_saved_connection else ""
    default_api_key = raw.get(CONFIG_KEY_SPELLING_API_KEY, "") if not has_saved_connection else ""
    default_disable_reasoning = raw.get(CONFIG_KEY_SPELLING_DISABLE_REASONING, False) if not has_saved_connection else False
    return CitationValidatorSettings(
        api_url=str(raw.get(CONFIG_KEY_CITATION_VALIDATOR_API_URL, default_api_url) or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_CITATION_VALIDATOR_MODEL_ID, default_model_id) or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_CITATION_VALIDATOR_API_KEY, default_api_key) or "").strip(),
        prompt=str(
            raw.get(CONFIG_KEY_CITATION_VALIDATOR_PROMPT, DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT)
            or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT
        ).strip(),
        disable_reasoning=bool(raw.get(CONFIG_KEY_CITATION_VALIDATOR_DISABLE_REASONING, default_disable_reasoning)),
    )


def save_citation_validator_settings(settings: CitationValidatorSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_CITATION_VALIDATOR_API_URL] = settings.api_url
    data[CONFIG_KEY_CITATION_VALIDATOR_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_CITATION_VALIDATOR_API_KEY] = settings.api_key
    data[CONFIG_KEY_CITATION_VALIDATOR_PROMPT] = settings.prompt or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT
    data[CONFIG_KEY_CITATION_VALIDATOR_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_improve1_settings() -> Improve1Settings:
    raw = _read_config()
    return Improve1Settings(
        api_url=str(raw.get(CONFIG_KEY_IMPROVE1_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_IMPROVE1_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_IMPROVE1_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_IMPROVE1_PROMPT, DEFAULT_IMPROVE1_PROMPT) or DEFAULT_IMPROVE1_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_IMPROVE1_DISABLE_REASONING), False),
    )


def save_improve1_settings(settings: Improve1Settings) -> None:
    data = _read_config()
    data[CONFIG_KEY_IMPROVE1_API_URL] = settings.api_url
    data[CONFIG_KEY_IMPROVE1_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_IMPROVE1_API_KEY] = settings.api_key
    data[CONFIG_KEY_IMPROVE1_PROMPT] = settings.prompt or DEFAULT_IMPROVE1_PROMPT
    data[CONFIG_KEY_IMPROVE1_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_improve2_settings() -> Improve2Settings:
    raw = _read_config()
    return Improve2Settings(
        api_url=str(raw.get(CONFIG_KEY_IMPROVE2_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_IMPROVE2_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_IMPROVE2_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_IMPROVE2_PROMPT, DEFAULT_IMPROVE2_PROMPT) or DEFAULT_IMPROVE2_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_IMPROVE2_DISABLE_REASONING), False),
    )


def save_improve2_settings(settings: Improve2Settings) -> None:
    data = _read_config()
    data[CONFIG_KEY_IMPROVE2_API_URL] = settings.api_url
    data[CONFIG_KEY_IMPROVE2_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_IMPROVE2_API_KEY] = settings.api_key
    data[CONFIG_KEY_IMPROVE2_PROMPT] = settings.prompt or DEFAULT_IMPROVE2_PROMPT
    data[CONFIG_KEY_IMPROVE2_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_combine_cites_settings() -> CombineCitesSettings:
    raw = _read_config()
    return CombineCitesSettings(
        api_url=str(raw.get(CONFIG_KEY_COMBINE_CITES_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_COMBINE_CITES_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_COMBINE_CITES_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_COMBINE_CITES_PROMPT, DEFAULT_COMBINE_CITES_PROMPT) or DEFAULT_COMBINE_CITES_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_COMBINE_CITES_DISABLE_REASONING), False),
    )


def save_combine_cites_settings(settings: CombineCitesSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_COMBINE_CITES_API_URL] = settings.api_url
    data[CONFIG_KEY_COMBINE_CITES_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_COMBINE_CITES_API_KEY] = settings.api_key
    data[CONFIG_KEY_COMBINE_CITES_PROMPT] = settings.prompt or DEFAULT_COMBINE_CITES_PROMPT
    data[CONFIG_KEY_COMBINE_CITES_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_thesaurus_settings() -> ThesaurusSettings:
    raw = _read_config()
    return ThesaurusSettings(
        api_url=str(raw.get(CONFIG_KEY_THESAURUS_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_THESAURUS_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_THESAURUS_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_THESAURUS_PROMPT, DEFAULT_THESAURUS_PROMPT) or DEFAULT_THESAURUS_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_THESAURUS_DISABLE_REASONING), False),
    )


def save_thesaurus_settings(settings: ThesaurusSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_THESAURUS_API_URL] = settings.api_url
    data[CONFIG_KEY_THESAURUS_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_THESAURUS_API_KEY] = settings.api_key
    data[CONFIG_KEY_THESAURUS_PROMPT] = settings.prompt or DEFAULT_THESAURUS_PROMPT
    data[CONFIG_KEY_THESAURUS_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_reference_settings() -> ReferenceSettings:
    raw = _read_config()
    return ReferenceSettings(
        api_url=str(raw.get(CONFIG_KEY_REFERENCE_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_REFERENCE_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_REFERENCE_API_KEY, "") or "").strip(),
        tavily_api_key=str(raw.get(CONFIG_KEY_REFERENCE_TAVILY_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_REFERENCE_PROMPT, DEFAULT_REFERENCE_PROMPT) or DEFAULT_REFERENCE_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_REFERENCE_DISABLE_REASONING), False),
    )


def save_reference_settings(settings: ReferenceSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_REFERENCE_API_URL] = settings.api_url
    data[CONFIG_KEY_REFERENCE_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_REFERENCE_API_KEY] = settings.api_key
    data[CONFIG_KEY_REFERENCE_TAVILY_API_KEY] = settings.tavily_api_key
    data[CONFIG_KEY_REFERENCE_PROMPT] = settings.prompt or DEFAULT_REFERENCE_PROMPT
    data[CONFIG_KEY_REFERENCE_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_ask_settings() -> AskSettings:
    raw = _read_config()
    prompt = str(raw.get(CONFIG_KEY_ASK_PROMPT, "") or "").strip()
    if not prompt:
        prompt = str(raw.get(CONFIG_KEY_REFERENCE_ASK_PROMPT, DEFAULT_ASK_PROMPT) or DEFAULT_ASK_PROMPT).strip()
    return AskSettings(
        api_url=str(raw.get(CONFIG_KEY_ASK_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_ASK_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_ASK_API_KEY, "") or "").strip(),
        tavily_api_key=str(raw.get(CONFIG_KEY_ASK_TAVILY_API_KEY, "") or "").strip(),
        prompt=prompt,
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_ASK_DISABLE_REASONING), False),
    )


def save_ask_settings(settings: AskSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_ASK_API_URL] = settings.api_url
    data[CONFIG_KEY_ASK_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_ASK_API_KEY] = settings.api_key
    data[CONFIG_KEY_ASK_TAVILY_API_KEY] = settings.tavily_api_key
    data[CONFIG_KEY_ASK_PROMPT] = settings.prompt or DEFAULT_ASK_PROMPT
    data[CONFIG_KEY_ASK_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_shorten_settings() -> ShortenSettings:
    raw = _read_config()
    return ShortenSettings(
        api_url=str(raw.get(CONFIG_KEY_SHORTEN_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_SHORTEN_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_SHORTEN_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_SHORTEN_PROMPT, DEFAULT_SHORTEN_PROMPT) or DEFAULT_SHORTEN_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_SHORTEN_DISABLE_REASONING), False),
    )


def save_shorten_settings(settings: ShortenSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_SHORTEN_API_URL] = settings.api_url
    data[CONFIG_KEY_SHORTEN_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_SHORTEN_API_KEY] = settings.api_key
    data[CONFIG_KEY_SHORTEN_PROMPT] = settings.prompt or DEFAULT_SHORTEN_PROMPT
    data[CONFIG_KEY_SHORTEN_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_introduction_settings() -> IntroductionSettings:
    raw = _read_config()
    return IntroductionSettings(
        api_url=str(raw.get(CONFIG_KEY_INTRO_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_INTRO_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_INTRO_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_INTRO_PROMPT, DEFAULT_INTRO_PROMPT) or DEFAULT_INTRO_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_INTRO_DISABLE_REASONING), False),
    )


def save_introduction_settings(settings: IntroductionSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_INTRO_API_URL] = settings.api_url
    data[CONFIG_KEY_INTRO_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_INTRO_API_KEY] = settings.api_key
    data[CONFIG_KEY_INTRO_PROMPT] = settings.prompt or DEFAULT_INTRO_PROMPT
    data[CONFIG_KEY_INTRO_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_introduction_reply_settings() -> IntroductionSettings:
    raw = _read_config()
    return IntroductionSettings(
        api_url=str(raw.get(CONFIG_KEY_INTRO_REPLY_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_INTRO_REPLY_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_INTRO_REPLY_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_INTRO_REPLY_PROMPT, DEFAULT_INTRO_REPLY_PROMPT) or DEFAULT_INTRO_REPLY_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_INTRO_REPLY_DISABLE_REASONING), False),
    )


def save_introduction_reply_settings(settings: IntroductionSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_INTRO_REPLY_API_URL] = settings.api_url
    data[CONFIG_KEY_INTRO_REPLY_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_INTRO_REPLY_API_KEY] = settings.api_key
    data[CONFIG_KEY_INTRO_REPLY_PROMPT] = settings.prompt or DEFAULT_INTRO_REPLY_PROMPT
    data[CONFIG_KEY_INTRO_REPLY_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_conclusion_settings() -> ConclusionSettings:
    raw = _read_config()
    return ConclusionSettings(
        api_url=str(raw.get(CONFIG_KEY_CONCLUSION_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_CONCLUSION_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_CONCLUSION_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_CONCLUSION_PROMPT, DEFAULT_CONCLUSION_PROMPT) or DEFAULT_CONCLUSION_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_CONCLUSION_DISABLE_REASONING), False),
    )


def save_conclusion_settings(settings: ConclusionSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_CONCLUSION_API_URL] = settings.api_url
    data[CONFIG_KEY_CONCLUSION_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_CONCLUSION_API_KEY] = settings.api_key
    data[CONFIG_KEY_CONCLUSION_PROMPT] = settings.prompt or DEFAULT_CONCLUSION_PROMPT
    data[CONFIG_KEY_CONCLUSION_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_concl_no_issues_settings() -> ConclNoIssuesSettings:
    raw = _read_config()
    return ConclNoIssuesSettings(
        api_url=str(raw.get(CONFIG_KEY_CONCL_NO_ISSUES_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_CONCL_NO_ISSUES_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_CONCL_NO_ISSUES_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_CONCL_NO_ISSUES_PROMPT, DEFAULT_CONCL_NO_ISSUES_PROMPT) or DEFAULT_CONCL_NO_ISSUES_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_CONCL_NO_ISSUES_DISABLE_REASONING), False),
    )


def save_concl_no_issues_settings(settings: ConclNoIssuesSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_CONCL_NO_ISSUES_API_URL] = settings.api_url
    data[CONFIG_KEY_CONCL_NO_ISSUES_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_CONCL_NO_ISSUES_API_KEY] = settings.api_key
    data[CONFIG_KEY_CONCL_NO_ISSUES_PROMPT] = settings.prompt or DEFAULT_CONCL_NO_ISSUES_PROMPT
    data[CONFIG_KEY_CONCL_NO_ISSUES_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_topic_sentence_settings() -> TopicSentenceSettings:
    raw = _read_config()
    return TopicSentenceSettings(
        api_url=str(raw.get(CONFIG_KEY_TOPIC_SENTENCE_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_TOPIC_SENTENCE_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_TOPIC_SENTENCE_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_TOPIC_SENTENCE_PROMPT, DEFAULT_TOPIC_SENTENCE_PROMPT) or DEFAULT_TOPIC_SENTENCE_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_TOPIC_SENTENCE_DISABLE_REASONING), False),
    )


def save_topic_sentence_settings(settings: TopicSentenceSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_TOPIC_SENTENCE_API_URL] = settings.api_url
    data[CONFIG_KEY_TOPIC_SENTENCE_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_TOPIC_SENTENCE_API_KEY] = settings.api_key
    data[CONFIG_KEY_TOPIC_SENTENCE_PROMPT] = settings.prompt or DEFAULT_TOPIC_SENTENCE_PROMPT
    data[CONFIG_KEY_TOPIC_SENTENCE_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_concl_section_settings() -> ConclSectionSettings:
    raw = _read_config()
    return ConclSectionSettings(
        api_url=str(raw.get(CONFIG_KEY_CONCL_SECTION_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_CONCL_SECTION_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_CONCL_SECTION_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_CONCL_SECTION_PROMPT, DEFAULT_CONCL_SECTION_PROMPT) or DEFAULT_CONCL_SECTION_PROMPT)
        .strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_CONCL_SECTION_DISABLE_REASONING), False),
    )


def save_concl_section_settings(settings: ConclSectionSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_CONCL_SECTION_API_URL] = settings.api_url
    data[CONFIG_KEY_CONCL_SECTION_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_CONCL_SECTION_API_KEY] = settings.api_key
    data[CONFIG_KEY_CONCL_SECTION_PROMPT] = settings.prompt or DEFAULT_CONCL_SECTION_PROMPT
    data[CONFIG_KEY_CONCL_SECTION_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_translate_settings() -> TranslateSettings:
    raw = _read_config()
    return TranslateSettings(
        api_url=str(raw.get(CONFIG_KEY_TRANSLATE_API_URL, "") or "").strip(),
        model_id=str(raw.get(CONFIG_KEY_TRANSLATE_MODEL_ID, "") or "").strip(),
        api_key=str(raw.get(CONFIG_KEY_TRANSLATE_API_KEY, "") or "").strip(),
        prompt=str(raw.get(CONFIG_KEY_TRANSLATE_PROMPT, DEFAULT_TRANSLATE_PROMPT) or DEFAULT_TRANSLATE_PROMPT).strip(),
        disable_reasoning=_coerce_bool_config(raw.get(CONFIG_KEY_TRANSLATE_DISABLE_REASONING), False),
    )


def save_translate_settings(settings: TranslateSettings) -> None:
    data = _read_config()
    data[CONFIG_KEY_TRANSLATE_API_URL] = settings.api_url
    data[CONFIG_KEY_TRANSLATE_MODEL_ID] = settings.model_id
    data[CONFIG_KEY_TRANSLATE_API_KEY] = settings.api_key
    data[CONFIG_KEY_TRANSLATE_PROMPT] = settings.prompt or DEFAULT_TRANSLATE_PROMPT
    data[CONFIG_KEY_TRANSLATE_DISABLE_REASONING] = bool(settings.disable_reasoning)
    _write_config(data)


def load_editor_source_file() -> Path | None:
    raw = _read_config()
    path = raw.get(CONFIG_KEY_EDITOR_SOURCE_FILE)
    if isinstance(path, str) and path.strip():
        return Path(path).expanduser().resolve(strict=False)
    return None


def save_editor_source_file(path: Path | None) -> None:
    data = _read_config()
    if path:
        data[CONFIG_KEY_EDITOR_SOURCE_FILE] = str(path.expanduser().resolve(strict=False))
    else:
        data.pop(CONFIG_KEY_EDITOR_SOURCE_FILE, None)
    _write_config(data)


def load_text_draft_template_dir() -> Path | None:
    raw = _read_config()
    path = raw.get(CONFIG_KEY_TEXT_DRAFT_TEMPLATE_DIR)
    if isinstance(path, str) and path.strip():
        return Path(path).expanduser().resolve(strict=False)
    return None


def save_text_draft_template_dir(path: Path | None) -> None:
    data = _read_config()
    if path:
        data[CONFIG_KEY_TEXT_DRAFT_TEMPLATE_DIR] = str(path.expanduser().resolve(strict=False))
    else:
        data.pop(CONFIG_KEY_TEXT_DRAFT_TEMPLATE_DIR, None)
    _write_config(data)


def _parse_text_draft_external_action(raw: Any) -> TextDraftExternalAction | None:
    if not isinstance(raw, dict):
        return None

    enabled = _coerce_bool_config(raw.get("enabled"), False)
    label = str(raw.get("label") or "External").strip() or "External"
    icon_name = str(raw.get("icon_name") or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON).strip()
    if not icon_name:
        icon_name = DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON
    tooltip = str(raw.get("tooltip") or "").strip()
    success_message = str(raw.get("success_message") or "").strip()

    command: list[str] = []
    raw_command = raw.get("command")
    if isinstance(raw_command, list):
        command = [str(part) for part in raw_command if str(part)]
    elif isinstance(raw_command, str) and raw_command.strip():
        try:
            command = shlex.split(raw_command)
        except ValueError:
            command = []

    cwd = None
    raw_cwd = raw.get("cwd")
    if isinstance(raw_cwd, str) and raw_cwd.strip():
        cwd = Path(raw_cwd).expanduser().resolve(strict=False)

    env: dict[str, str] = {}
    raw_env = raw.get("env")
    if isinstance(raw_env, dict):
        env = {str(key): str(value) for key, value in raw_env.items() if str(key)}

    return TextDraftExternalAction(
        enabled=enabled,
        label=label,
        command=command,
        cwd=cwd,
        env=env,
        icon_name=icon_name,
        tooltip=tooltip,
        success_message=success_message,
    )


def _format_text_draft_external_action_env(env: dict[str, str]) -> str:
    return "\n".join(f"{key}={value}" for key, value in env.items())


def _parse_text_draft_external_action_env(text: str) -> tuple[dict[str, str] | None, str | None]:
    env: dict[str, str] = {}
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue
        if "=" not in line:
            return None, f"line {line_number} must be KEY=value."
        key, value = line.split("=", 1)
        key = key.strip()
        if not TEXT_DRAFT_EXTERNAL_ACTION_ENV_NAME_RE.fullmatch(key):
            return None, f"line {line_number} has an invalid environment variable name."
        env[key] = value.strip()
    return env, None


def load_text_draft_external_actions() -> list[TextDraftExternalAction]:
    raw = _read_config()
    raw_actions = raw.get(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTIONS)
    if isinstance(raw_actions, list):
        actions = [
            action
            for action in (_parse_text_draft_external_action(item) for item in raw_actions)
            if action is not None and not _is_ignored_text_draft_codex_action(action)
        ]
        return actions

    legacy_action = _parse_text_draft_external_action(raw.get(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION))
    if legacy_action is None or _is_ignored_text_draft_codex_action(legacy_action):
        return []
    return [legacy_action]


def _text_draft_external_action_has_payload(action: TextDraftExternalAction) -> bool:
    label = action.label.strip()
    has_custom_label = bool(label and label != "External")
    icon_name = action.icon_name.strip()
    has_custom_icon = bool(icon_name and icon_name != DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON)
    return (
        bool(action.command)
        or has_custom_label
        or action.cwd is not None
        or bool(action.env)
        or has_custom_icon
        or bool(action.tooltip.strip())
        or bool(action.success_message.strip())
    )


def _text_draft_external_action_payload(action: TextDraftExternalAction) -> dict[str, Any]:
    label = action.label.strip()
    icon_name = action.icon_name.strip() or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON
    payload: dict[str, Any] = {
        "enabled": bool(action.enabled),
        "label": label or "External",
        "command": [str(part) for part in action.command if str(part)],
        "icon_name": icon_name,
    }
    if action.cwd is not None:
        payload["cwd"] = str(action.cwd.expanduser().resolve(strict=False))
    if action.env:
        payload["env"] = dict(action.env)
    if action.tooltip.strip():
        payload["tooltip"] = action.tooltip.strip()
    if action.success_message.strip():
        payload["success_message"] = action.success_message.strip()
    return payload


def save_text_draft_external_actions(actions: list[TextDraftExternalAction]) -> None:
    data = _read_config()
    payload = [
        _text_draft_external_action_payload(action)
        for action in actions
        if _text_draft_external_action_has_payload(action)
        and not _is_ignored_text_draft_codex_action(action)
    ]
    data.pop(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION, None)
    if payload:
        data[CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTIONS] = payload
    else:
        data.pop(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTIONS, None)
    _write_config(data)


def load_text_draft_external_action() -> TextDraftExternalAction:
    actions = load_text_draft_external_actions()
    if actions:
        return actions[0]
    return TextDraftExternalAction(enabled=False, label="", command=[], cwd=None, env={})


def save_text_draft_external_action(action: TextDraftExternalAction) -> None:
    save_text_draft_external_actions([action])


def _is_ignored_text_draft_codex_action(action: TextDraftExternalAction) -> bool:
    if action.label.strip().lower() == "codex":
        return True
    return any(
        script_name in part
        for part in action.command
        for script_name in (
            "prose-text-draft-codex-ghostty.sh",
            "prose-text-draft-codex-vte.sh",
        )
    )


def load_last_odt_file() -> Path | None:
    raw = _read_config()
    path = raw.get(CONFIG_KEY_LAST_ODT_FILE)
    if isinstance(path, str) and path.strip():
        return Path(path).expanduser().resolve(strict=False)
    return None


def save_last_odt_file(path: Path | None) -> None:
    data = _read_config()
    if path:
        data[CONFIG_KEY_LAST_ODT_FILE] = str(path.expanduser().resolve(strict=False))
    else:
        data.pop(CONFIG_KEY_LAST_ODT_FILE, None)
    _write_config(data)


def _default_editor_pinned_actions() -> list[str]:
    return list(DEFAULT_EDITOR_PINNED_ACTION_IDS)


def _sanitize_editor_pinned_actions(raw: Any) -> list[str]:
    pinned: list[str] = []
    seen: set[str] = set()
    if not isinstance(raw, list):
        return pinned
    for item in raw:
        key = _normalize_editor_quick_action_key(str(item or "").strip())
        if key not in EDITOR_QUICK_ACTION_BY_KEY or key in EDITOR_TOOLBAR_EXCLUDED_ACTION_KEYS or key in seen:
            continue
        pinned.append(key)
        seen.add(key)
        if len(pinned) >= MAX_PINNED_EDITOR_ACTIONS:
            break
    return pinned


def _ordered_editor_quick_action_keys(pinned_action_ids: Iterable[str]) -> list[str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for key in pinned_action_ids:
        key = _normalize_editor_quick_action_key(key)
        if key not in EDITOR_QUICK_ACTION_BY_KEY or key in EDITOR_TOOLBAR_EXCLUDED_ACTION_KEYS or key in seen:
            continue
        ordered.append(key)
        seen.add(key)
    for definition in EDITOR_QUICK_ACTIONS:
        if definition.key in seen or definition.key in EDITOR_TOOLBAR_EXCLUDED_ACTION_KEYS:
            continue
        ordered.append(definition.key)
    return ordered


def _default_text_draft_pinned_actions() -> list[str]:
    return list(DEFAULT_TEXT_DRAFT_PINNED_ACTION_IDS)


def _sanitize_text_draft_pinned_actions(raw: Any) -> list[str]:
    pinned: list[str] = []
    seen: set[str] = set()
    if not isinstance(raw, list):
        return pinned
    for item in raw:
        key = str(item or "").strip()
        if key not in TEXT_DRAFT_QUICK_ACTION_BY_KEY or key in seen:
            continue
        pinned.append(key)
        seen.add(key)
        if len(pinned) >= MAX_PINNED_EDITOR_ACTIONS:
            break
    return pinned


def _ordered_text_draft_quick_action_keys(pinned_action_ids: Iterable[str]) -> list[str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for key in pinned_action_ids:
        if key not in TEXT_DRAFT_QUICK_ACTION_BY_KEY or key in seen:
            continue
        ordered.append(key)
        seen.add(key)
    for definition in TEXT_DRAFT_QUICK_ACTIONS:
        if definition.key in seen:
            continue
        ordered.append(definition.key)
    return ordered


def load_editor_pinned_actions() -> list[str]:
    raw = _read_config()
    if CONFIG_KEY_EDITOR_PINNED_ACTIONS not in raw:
        return _default_editor_pinned_actions()
    stored = raw.get(CONFIG_KEY_EDITOR_PINNED_ACTIONS)
    if not isinstance(stored, list):
        return _default_editor_pinned_actions()
    return _sanitize_editor_pinned_actions(stored)


def save_editor_pinned_actions(action_ids: Iterable[str]) -> None:
    data = _read_config()
    data[CONFIG_KEY_EDITOR_PINNED_ACTIONS] = _sanitize_editor_pinned_actions(list(action_ids))
    _write_config(data)


def load_text_draft_pinned_actions() -> list[str]:
    raw = _read_config()
    if CONFIG_KEY_TEXT_DRAFT_PINNED_ACTIONS not in raw:
        return _default_text_draft_pinned_actions()
    stored = raw.get(CONFIG_KEY_TEXT_DRAFT_PINNED_ACTIONS)
    if not isinstance(stored, list):
        return _default_text_draft_pinned_actions()
    return _sanitize_text_draft_pinned_actions(stored)


def save_text_draft_pinned_actions(action_ids: Iterable[str]) -> None:
    data = _read_config()
    data[CONFIG_KEY_TEXT_DRAFT_PINNED_ACTIONS] = _sanitize_text_draft_pinned_actions(list(action_ids))
    _write_config(data)


@dataclass
class Suggestion:
    title: str
    page: int | None
    snippet: str
    replacement: str
    reasoning: str


@dataclass(frozen=True)
class ProofreadChunk:
    index: int
    paragraphs: list[str]


@dataclass(frozen=True)
class ProofreadParagraph:
    text: str
    page: int | None


@dataclass(frozen=True)
class TextDraftTemplateCategory:
    key: str
    name: str
    template_paths: list[Path]


__all__ = [name for name in globals() if not name.startswith("__")]
