#!/usr/bin/python3

from __future__ import annotations

import difflib
import json
import importlib
import os
import re
import signal
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

gi.require_version("Gtk", "4.0")
gi.require_version("Adw", "1")
from gi.repository import Adw, Gdk, Gio, GLib, Gtk, Pango  # type: ignore

Vte = None  # type: ignore[assignment]
try:
    gi.require_version("Vte", "3.91")
    from gi.repository import Vte as VteModule  # type: ignore

    Vte = VteModule  # type: ignore[assignment]
except (ImportError, ValueError):
    Vte = None  # type: ignore[assignment]

uno = None  # type: ignore
PropertyValue = None  # type: ignore[assignment]
NoConnectException = Exception  # type: ignore[assignment]
XModel = None  # type: ignore[assignment]
ControlCharacter = None  # type: ignore[assignment]
XTextDocument = None  # type: ignore[assignment]


APP_ID = "com.mcglaw.Prose"
ACTION_OBJECT_PATH = "/org/gtk/Application"
APP_NAME = "Prose"
GLib.set_application_name(APP_NAME)

CONFIG_FILE = Path(__file__).with_name("config.json")
TEXT_DRAFT_CODEX_VTE_SCRIPT = (
    Path(__file__).resolve().parent / "scripts" / "prose-text-draft-codex-vte.sh"
)
TEXT_DRAFT_PROJECTS_DIR = Path(__file__).resolve().parent.parent
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
CONFIG_KEY_TEXT_DRAFT_CODEX_PROJECT_ACTION = "text_draft_codex_project_action"
CONFIG_KEY_TEXT_DRAFT_CODEX_SESSION_ACTION = "text_draft_codex_session_action"
CONFIG_KEY_SHARED_STYLE_RULES = "shared_style_rules"
CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION = "text_draft_external_action"
CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTIONS = "text_draft_external_actions"
TEXT_DRAFT_EXTERNAL_ACTION_DRAFT_FILE_TOKEN = "{draft_file}"
TEXT_DRAFT_CODEX_REASONING_EFFORTS = ("minimal", "low", "medium", "high", "xhigh")
DEFAULT_TEXT_DRAFT_CODEX_REASONING_EFFORT = "medium"
TEXT_DRAFT_CASE_SUGGESTION_LIMIT = 25
TEXT_DRAFT_CASE_SUGGESTION_MAX_HEIGHT = 220
DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON = "utilities-terminal-symbolic"
PROSE_TERMINAL_DARK_FOREGROUND = "#f2f4f8"
PROSE_TERMINAL_DARK_BACKGROUND = "#3d3d3d"
PROSE_TERMINAL_DARK_SELECTION = "#3d536b"
PROSE_TERMINAL_DARK_CURSOR = "#8ab4f8"
PROSE_TERMINAL_DARK_CURSOR_FOREGROUND = "#111318"
PROSE_TERMINAL_DARK_PALETTE = (
    "#1f2329",
    "#ff7b86",
    "#7bd88f",
    "#f4cf65",
    "#8ab4f8",
    "#c58af9",
    "#6fd6e8",
    "#e6e9ef",
    "#7f8b99",
    "#ff9aa2",
    "#9be7ad",
    "#f8dd85",
    "#a8c7fa",
    "#d7aefb",
    "#8de8f7",
    "#ffffff",
)
PROSE_TERMINAL_LIGHT_FOREGROUND = "#20242c"
PROSE_TERMINAL_LIGHT_BACKGROUND = "#f5f5f5"
PROSE_TERMINAL_LIGHT_SELECTION = "#d7e4f5"
PROSE_TERMINAL_LIGHT_CURSOR = "#1f66d1"
PROSE_TERMINAL_LIGHT_CURSOR_FOREGROUND = "#ffffff"
PROSE_TERMINAL_LIGHT_PALETTE = (
    "#2d333b",
    "#c2414b",
    "#2f8f4e",
    "#8a6a00",
    "#1f66d1",
    "#8c4ac9",
    "#1f7a8c",
    "#f5f7fa",
    "#66717f",
    "#d95560",
    "#3fae63",
    "#a78300",
    "#327fe3",
    "#a35bd8",
    "#2695aa",
    "#ffffff",
)

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


def _rgba_color(spec: str) -> Gdk.RGBA:
    color = Gdk.RGBA()
    if not color.parse(spec):
        raise ValueError(f"Invalid color: {spec}")
    return color


def _apply_prose_terminal_theme(terminal: Any) -> None:
    dark = Adw.StyleManager.get_default().get_dark()
    if dark:
        foreground_spec = PROSE_TERMINAL_DARK_FOREGROUND
        background_spec = PROSE_TERMINAL_DARK_BACKGROUND
        selection_spec = PROSE_TERMINAL_DARK_SELECTION
        cursor_spec = PROSE_TERMINAL_DARK_CURSOR
        cursor_foreground_spec = PROSE_TERMINAL_DARK_CURSOR_FOREGROUND
        palette_specs = PROSE_TERMINAL_DARK_PALETTE
    else:
        foreground_spec = PROSE_TERMINAL_LIGHT_FOREGROUND
        background_spec = PROSE_TERMINAL_LIGHT_BACKGROUND
        selection_spec = PROSE_TERMINAL_LIGHT_SELECTION
        cursor_spec = PROSE_TERMINAL_LIGHT_CURSOR
        cursor_foreground_spec = PROSE_TERMINAL_LIGHT_CURSOR_FOREGROUND
        palette_specs = PROSE_TERMINAL_LIGHT_PALETTE

    foreground = _rgba_color(foreground_spec)
    background = _rgba_color(background_spec)
    palette = [_rgba_color(spec) for spec in palette_specs]
    terminal.set_colors(foreground, background, palette)
    terminal.set_color_background(background)
    terminal.set_color_foreground(foreground)
    terminal.set_clear_background(True)
    terminal.set_color_cursor(_rgba_color(cursor_spec))
    terminal.set_color_cursor_foreground(_rgba_color(cursor_foreground_spec))
    terminal.set_color_highlight(_rgba_color(selection_spec))
    terminal.set_color_highlight_foreground(foreground)


@dataclass
class ModelProfile:
    key: str
    nickname: str
    abbreviation: str
    api_url: str
    model_id: str
    api_key: str
    disable_reasoning: bool

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
    codex_reasoning_effort: str = DEFAULT_TEXT_DRAFT_CODEX_REASONING_EFFORT

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
    codex_reasoning_dropdown: Gtk.DropDown | None
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
        ),
        ModelProfile(
            key="profile4",
            nickname=DEFAULT_MODEL_PROFILE_NICKNAMES["profile4"],
            abbreviation="",
            api_url="",
            model_id="",
            api_key="",
            disable_reasoning=False,
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
        codex_reasoning_effort=_sanitize_text_draft_codex_reasoning_effort(
            raw.get("codex_reasoning_effort")
        ),
    )


def _sanitize_text_draft_codex_reasoning_effort(raw: Any) -> str:
    reasoning_effort = str(raw or "").strip().lower()
    if reasoning_effort in TEXT_DRAFT_CODEX_REASONING_EFFORTS:
        return reasoning_effort
    return DEFAULT_TEXT_DRAFT_CODEX_REASONING_EFFORT


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
            if action is not None
        ]
        return actions

    legacy_action = _parse_text_draft_external_action(raw.get(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION))
    return [legacy_action] if legacy_action is not None else []


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
    if _is_text_draft_codex_external_action(action):
        payload["codex_reasoning_effort"] = _sanitize_text_draft_codex_reasoning_effort(
            action.codex_reasoning_effort
        )
    return payload


def save_text_draft_external_actions(actions: list[TextDraftExternalAction]) -> None:
    data = _read_config()
    payload = [
        _text_draft_external_action_payload(action)
        for action in actions
        if _text_draft_external_action_has_payload(action)
    ]
    data.pop(CONFIG_KEY_TEXT_DRAFT_EXTERNAL_ACTION, None)
    data.pop(CONFIG_KEY_TEXT_DRAFT_CODEX_PROJECT_ACTION, None)
    data.pop(CONFIG_KEY_TEXT_DRAFT_CODEX_SESSION_ACTION, None)
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


def _is_text_draft_codex_external_action(action: TextDraftExternalAction) -> bool:
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


@dataclass(frozen=True)
class TextDraftCaseSuggestion:
    citation: str
    display_name: str
    search_terms: tuple[str, ...]


class ProseApp(Adw.Application):
    def __init__(self) -> None:
        super().__init__(application_id=APP_ID, flags=Gio.ApplicationFlags.HANDLES_OPEN)
        self._window: ProseWindow | None = None

    def _ensure_window(self) -> ProseWindow:
        if not self._window:
            self._window = ProseWindow(self)
        return self._window

    def do_activate(self) -> None:  # noqa: D401
        self._ensure_window().present()

    def do_open(self, files: list[Gio.File], n_files: int, _hint: str) -> None:  # noqa: D401
        window = self._ensure_window()
        window.present()
        if n_files != 1:
            window.show_open_error("Open one .odt file at a time.")
            return
        file_path = files[0].get_path()
        if not file_path:
            window.show_open_error("Choose a local .odt file.")
            return
        window.open_odt_path(Path(file_path))


class ProseWindow(Adw.ApplicationWindow):
    def __init__(self, app: ProseApp) -> None:
        super().__init__(application=app, title=APP_NAME, default_width=920, default_height=680)
        _migrate_shared_style_rules_prompts()
        self._libreoffice_python_path = load_libreoffice_python_path()
        _import_uno_from_candidates(self._libreoffice_python_path)
        self._shared_style_rules = load_shared_style_rules()
        self._proof_settings = load_proofread_settings()
        self._spelling_settings = load_spellingstyle_settings()
        self._text_draft_spelling_prompt = load_text_draft_spellingstyle_prompt()
        self._citation_validator_settings = load_citation_validator_settings()
        self._improve1_settings = load_improve1_settings()
        self._improve2_settings = load_improve2_settings()
        self._combine_cites_settings = load_combine_cites_settings()
        self._thesaurus_settings = load_thesaurus_settings()
        self._reference_settings = load_reference_settings()
        self._ask_settings = load_ask_settings()
        self._shorten_settings = load_shorten_settings()
        self._introduction_settings = load_introduction_settings()
        self._introduction_reply_settings = load_introduction_reply_settings()
        self._conclusion_settings = load_conclusion_settings()
        self._concl_no_issues_settings = load_concl_no_issues_settings()
        self._topic_sentence_settings = load_topic_sentence_settings()
        self._concl_section_settings = load_concl_section_settings()
        self._translate_settings = load_translate_settings()
        self._model_profiles = load_model_profiles()
        self._editor_action_profile_defaults = load_editor_action_profile_defaults(
            self._model_profiles,
            {
                "improve-generated": self._improve1_settings,
                "rephrase-generated": self._improve2_settings,
                "spelling": self._spelling_settings,
                "combine": self._combine_cites_settings,
                "thesaurus": self._thesaurus_settings,
                "shorten": self._shorten_settings,
                "intro": self._introduction_settings,
                "intro-reply": self._introduction_reply_settings,
                "conclusion": self._conclusion_settings,
                "concl-no-issues": self._concl_no_issues_settings,
                "topic-sentence": self._topic_sentence_settings,
                "concl-section": self._concl_section_settings,
                "translate": self._translate_settings,
            },
        )
        self._prefix_settings = load_prefix_settings()
        self._editor_source_file = load_editor_source_file()
        self._text_draft_template_dir = load_text_draft_template_dir()
        self._text_draft_external_actions = load_text_draft_external_actions()
        self._last_odt_path = load_last_odt_file()
        self._concordance_file_path = load_concordance_file_path()
        self._editor_pinned_action_ids = load_editor_pinned_actions()
        self._text_draft_pinned_action_ids = load_text_draft_pinned_actions()
        self._ctx = None
        self._desktop = None
        self._active_doc = None
        self._listener_proc: subprocess.Popen[str] | None = None
        self._suggestions: list[Suggestion] = []
        self._suggestion_rows: list[Gtk.Box] = []
        self._auto_suggestions_json_attempt_signature: tuple[str, int, int] | None = None
        self._current_doc_path: Path | None = None
        self._busy = False
        self._last_model_response_raw: str | None = None
        self._last_model_response_title: str | None = None
        self._last_model_response_api_url: str | None = None
        self._last_model_response_timestamp: str | None = None
        self._last_model_response_diagnostic: str | None = None
        self._model_response_window: Adw.ApplicationWindow | None = None
        self._last_editor_request_raw: str | None = None
        self._last_editor_request_title: str | None = None
        self._last_editor_request_api_url: str | None = None
        self._last_editor_request_timestamp: str | None = None
        self._editor_prompt_window: Adw.ApplicationWindow | None = None
        self._editor_insert_cursor = None
        self._editor_insert_doc = None
        self._editor_insert_end = None
        self._improve_insert_cursor = None
        self._improve_insert_doc = None
        self._improve_started = False
        self._last_insert_len = 0
        self._editor_pending_newlines = 0
        self._improve_pending_newlines = 0
        self._pending_regenerate_context: RegenerateContext | None = None
        self._last_regenerate_context: RegenerateContext | None = None
        self._spelling_output_buffer: Gtk.TextBuffer | None = None
        self._editor_output_surface: Gtk.Stack | None = None
        self._multi_draft_window: Adw.ApplicationWindow | None = None
        self._multi_draft_grid: Gtk.Grid | None = None
        self._multi_draft_choices: dict[str, MultiDraftChoice] = {}
        self._multi_draft_lock = threading.Lock()
        self._multi_draft_run_id = 0
        self._multi_draft_remaining = 0
        self._multi_draft_cancel_event: threading.Event | None = None
        self._multi_draft_insert_buttons: list[Gtk.Button] = []
        self._multi_draft_insert_mode = "editor"
        self._multi_draft_replace_doc: XTextDocument | None = None  # type: ignore[type-arg]
        self._multi_draft_replace_start = None
        self._multi_draft_replace_end = None
        self._text_draft_original_output_buffer: Gtk.TextBuffer | None = None
        self._text_draft_buffer: Gtk.TextBuffer | None = None
        self._text_draft_view: Gtk.TextView | None = None
        self._text_draft_draft_surface: Gtk.Stack | None = None
        self._text_draft_clear_button: Gtk.Button | None = None
        self._text_draft_terminal = None
        self._text_draft_terminal_close_button: Gtk.Button | None = None
        self._text_draft_terminal_buttons: list[Gtk.Widget] = []
        self._text_draft_terminal_active = False
        self._text_draft_terminal_pid: int | None = None
        self._text_draft_terminal_mode: str | None = None
        self._text_draft_terminal_cwd: str | None = None
        self._text_draft_terminal_closing = False
        self._text_draft_terminal_generated_output = ""
        self._text_draft_terminal_replace_text = ""
        self._text_draft_clear_draft_on_terminal_return = False
        self._text_draft_insert_start_mark: Gtk.TextMark | None = None
        self._text_draft_insert_end_mark: Gtk.TextMark | None = None
        self._text_draft_pending_regenerate_context: RegenerateContext | None = None
        self._text_draft_last_regenerate_context: RegenerateContext | None = None
        self._text_draft_regenerate_menu_button: Gtk.MenuButton | None = None
        self._text_draft_regenerate_profile_buttons: list[Gtk.Button] = []
        self._text_draft_action_buttons: list[Gtk.Widget] = []
        self._text_draft_actions_wrap: Gtk.FlowBox | None = None
        self._text_draft_template_button: Gtk.MenuButton | None = None
        self._text_draft_templates_popover: Gtk.Popover | None = None
        self._text_draft_templates_grid: Gtk.FlowBox | None = None
        self._text_draft_templates_title: Gtk.Label | None = None
        self._text_draft_templates_back_button: Gtk.Button | None = None
        self._text_draft_template_categories: list[TextDraftTemplateCategory] = []
        self._text_draft_template_current_category_key: str | None = None
        self._text_draft_template_menu_buttons: list[Gtk.Widget] = []
        self._text_draft_template_email_box: Gtk.Box | None = None
        self._text_draft_template_email_buttons: list[Gtk.Button] = []
        self._text_draft_template_emails: list[tuple[str, str]] = []
        self._text_draft_external_action_box: Gtk.Box | None = None
        self._text_draft_external_action_buttons: list[Gtk.Widget] = []
        self._text_draft_case_search_entry: Gtk.Entry | None = None
        self._text_draft_case_results_scroller: Gtk.ScrolledWindow | None = None
        self._text_draft_case_list_box: Gtk.ListBox | None = None
        self._text_draft_case_attachments_box: Gtk.Box | None = None
        self._text_draft_case_suggestions: list[TextDraftCaseSuggestion] = []
        self._text_draft_case_selected_index = 0
        self._text_draft_case_prefix = ""
        self._text_draft_case_prefix_bounds: tuple[int, int] | None = None
        self._text_draft_case_attachments: list[str] = []
        self._text_draft_case_index: list[TextDraftCaseSuggestion] = []
        self._text_draft_case_index_path: Path | None = None
        self._text_draft_case_index_mtime_ns: int | None = None
        self._text_draft_temp_path: Path | None = None
        self._text_draft_temp_flush_source_id = 0
        self._text_draft_temp_dirty = False
        self._text_draft_pending_newlines = 0
        self._text_draft_original_pending_newlines = 0
        self._css_provider: Gtk.CssProvider | None = None
        self._thesaurus_words: list[str] = []
        self._thesaurus_rows: list[Gtk.Widget] = []
        self._reference_output_text = ""
        self._reference_placeholder_active = True
        self._reference_output_mode = "thesaurus"
        self._combine_cites_doc = None
        self._combine_cites_cursor = None
        self._settings_window: SettingsWindow | None = None
        self._editor_commands_window: EditorCommandsWindow | None = None
        self._shortcuts_window: Gtk.ShortcutsWindow | None = None
        self._rt_prefix_entry: Gtk.Entry | None = None
        self._ct_prefix_entry: Gtk.Entry | None = None
        self._substitution_rows: list[tuple[Gtk.Entry, Gtk.Entry]] = []
        self._regenerate_menu_button: Gtk.MenuButton | None = None
        self._regenerate_profile_buttons: list[Gtk.Button] = []
        self._transform_action_buttons: list[Gtk.Widget] = []
        self._transform_actions_wrap: Gtk.FlowBox | None = None
        self.connect("close-request", self._on_window_close_request)
        self._build_ui()
        self._ensure_menu()
        self._register_actions()
        self._update_uno_status()

    # UI -----------------------------------------------------------------
    def _build_ui(self) -> None:
        self._ensure_css()
        view = Adw.ToolbarView()
        overlay = Adw.ToastOverlay()
        overlay.set_child(view)
        self.set_content(overlay)
        self._overlay = overlay

        header = Adw.HeaderBar()
        header.add_css_class("flat")
        view.add_top_bar(header)

        menu_button = Gtk.MenuButton(icon_name="open-menu-symbolic")
        header.pack_end(menu_button)
        self._menu_button = menu_button

        launch_btn = Gtk.Button(label="Launch Writer", icon_name="document-open-symbolic")
        launch_btn.add_css_class("flat")
        launch_btn.set_action_name("app.launch-writer")
        header.pack_start(launch_btn)
        self._launch_btn = launch_btn

        open_last_btn = Gtk.Button(icon_name="document-open-recent-symbolic")
        open_last_btn.add_css_class("flat")
        open_last_btn.set_action_name("app.open-last-odt")
        open_last_btn.set_tooltip_text("Open last Writer document")
        open_last_btn.set_sensitive(bool(self._last_odt_path))
        header.pack_start(open_last_btn)
        self._open_last_btn = open_last_btn

        status_label = Gtk.Label(label="LibreOffice status unknown", halign=Gtk.Align.START)
        status_label.add_css_class("dim-label")
        status_label.set_wrap(True)
        status_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        status_label.set_xalign(0.0)
        status_label.set_hexpand(True)
        self._status_label = status_label

        status_spinner = Gtk.Spinner()
        status_spinner.set_spinning(False)
        status_spinner.set_visible(False)
        self._status_spinner = status_spinner

        status_row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        status_row.set_halign(Gtk.Align.FILL)
        status_row.set_hexpand(True)
        status_row.append(status_spinner)
        status_row.append(status_label)

        mode_switcher = Adw.ViewSwitcher()
        mode_switcher.set_policy(Adw.ViewSwitcherPolicy.WIDE)

        stack = Adw.ViewStack()
        stack.set_hexpand(True)
        stack.set_vexpand(True)
        stack.connect("notify::visible-child-name", self._on_main_stack_visible_child_name_changed)
        self._main_stack = stack
        mode_switcher.set_stack(stack)

        editor_panel = self._build_editor_panel()
        text_draft_panel = self._build_text_draft_panel()
        proof_panel = self._build_proof_panel()
        prefixes_panel = self._build_prefixes_panel()
        editor_page = stack.add_titled(editor_panel, "editor", "Editor")
        editor_page.set_icon_name("document-edit-symbolic")
        text_draft_page = stack.add_titled(text_draft_panel, "text-draft", "Text Draft")
        text_draft_page.set_icon_name("text-x-generic-symbolic")
        proof_page = stack.add_titled(proof_panel, "proof", "Proof Reader")
        proof_page.set_icon_name("tools-check-spelling-symbolic")
        prefixes_page = stack.add_titled(prefixes_panel, "prefixes", "Prefixes")
        prefixes_page.set_icon_name("format-indent-more-symbolic")
        stack.set_visible_child_name("editor")

        top_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        top_box.set_margin_top(18)
        top_box.set_margin_bottom(6)
        top_box.set_margin_start(18)
        top_box.set_margin_end(18)
        top_box.append(status_row)
        top_box.append(mode_switcher)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        content.set_vexpand(True)
        content.append(top_box)
        content.append(stack)
        view.set_content(content)

    def _build_editor_panel(self) -> Gtk.Box:
        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        panel.set_margin_top(6)
        panel.set_margin_bottom(12)
        panel.set_margin_start(18)
        panel.set_margin_end(18)

        action_wrap = Gtk.FlowBox()
        action_wrap.set_selection_mode(Gtk.SelectionMode.NONE)
        action_wrap.set_column_spacing(6)
        action_wrap.set_row_spacing(6)
        action_wrap.set_max_children_per_line(6)
        action_wrap.set_hexpand(True)
        panel.append(action_wrap)
        self._transform_actions_wrap = action_wrap
        self._rebuild_transform_action_buttons()

        split_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        split_box.set_hexpand(True)
        split_box.set_vexpand(True)

        output_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        output_section.set_hexpand(True)
        output_section.set_vexpand(True)
        output_header = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        output_label = Gtk.Label(label="Original output", xalign=0)
        output_label.add_css_class("dim-label")
        output_label.set_hexpand(True)
        output_label.set_halign(Gtk.Align.START)
        output_header.append(output_label)
        regenerate_menu_button = Gtk.MenuButton(label="Retry")
        regenerate_menu_button.set_tooltip_text("Try the last generated output again with another model profile.")
        self._apply_quick_action_button_classes(regenerate_menu_button)
        output_header.append(regenerate_menu_button)
        self._regenerate_menu_button = regenerate_menu_button
        self._rebuild_regenerate_profile_chips()

        output_section.append(output_header)

        output_surface = Gtk.Stack()
        output_surface.set_hexpand(True)
        output_surface.set_vexpand(True)
        output_surface.add_css_class("editor-lookup-surface")
        output_scroller = Gtk.ScrolledWindow()
        output_scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        output_scroller.set_hexpand(True)
        output_scroller.set_vexpand(True)
        output_scroller.set_min_content_height(160)
        output_scroller.add_css_class("spelling-output-scroller")
        output_scroller.add_css_class("editable-text-scroller")
        output_scroller.add_css_class("editor-lookup-page")
        output_buffer = Gtk.TextBuffer()
        output_view = Gtk.TextView.new_with_buffer(output_buffer)
        output_view.set_monospace(False)
        output_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        output_view.set_vexpand(True)
        output_view.set_hexpand(True)
        output_view.set_size_request(-1, 160)
        output_view.set_left_margin(SPELLING_OUTPUT_PADDING_PX)
        output_view.set_right_margin(SPELLING_OUTPUT_PADDING_PX)
        output_view.set_top_margin(SPELLING_OUTPUT_PADDING_PX)
        output_view.set_bottom_margin(SPELLING_OUTPUT_PADDING_PX)
        output_view.add_css_class("spelling-output-view")
        output_view.add_css_class("editable-text-view")
        output_scroller.set_child(output_view)
        output_surface.add_named(output_scroller, "output")
        output_surface.set_visible_child_name("output")
        output_section.append(output_surface)
        self._spelling_output_buffer = output_buffer
        self._editor_output_surface = output_surface

        thesaurus_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        thesaurus_section.set_hexpand(True)
        thesaurus_section.set_vexpand(True)
        thesaurus_header = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        thesaurus_header.set_margin_bottom(6)
        thesaurus_btn = Gtk.Button(label="Thesaurus")
        thesaurus_btn.add_css_class("flat")
        thesaurus_btn.add_css_class("reference-toggle")
        thesaurus_btn.add_css_class("transform-pill")
        thesaurus_btn.connect("clicked", self._on_thesaurus_button_clicked)
        thesaurus_header.append(thesaurus_btn)
        reference_btn = Gtk.Button(label="Look Up")
        reference_btn.add_css_class("flat")
        reference_btn.add_css_class("reference-toggle")
        reference_btn.add_css_class("transform-pill")
        reference_btn.connect("clicked", self._on_reference_button_clicked)
        thesaurus_header.append(reference_btn)
        reference_query_entry = Gtk.Entry()
        reference_query_entry.set_placeholder_text("Ask")
        reference_query_entry.set_hexpand(True)
        reference_query_entry.add_css_class("reference-query-entry")
        reference_query_entry.connect("activate", self._on_reference_question_activated)
        thesaurus_header.append(reference_query_entry)
        thesaurus_section.append(thesaurus_header)
        self._thesaurus_btn = thesaurus_btn
        self._reference_btn = reference_btn
        self._reference_query_entry = reference_query_entry

        thesaurus_stack = Gtk.Stack()
        thesaurus_stack.set_hexpand(True)
        thesaurus_stack.set_vexpand(True)
        thesaurus_stack.add_css_class("editor-lookup-surface")

        thesaurus_scroller = Gtk.ScrolledWindow()
        thesaurus_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        thesaurus_scroller.set_hexpand(True)
        thesaurus_scroller.set_vexpand(True)
        thesaurus_scroller.set_min_content_height(160)
        thesaurus_scroller.add_css_class("spelling-output-scroller")
        thesaurus_scroller.add_css_class("editor-lookup-page")
        thesaurus_list = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=0)
        thesaurus_list.set_hexpand(True)
        thesaurus_list.set_vexpand(True)
        thesaurus_list.add_css_class("thesaurus-results")
        thesaurus_placeholder = Gtk.Label(label="")
        thesaurus_placeholder.add_css_class("dim-label")
        thesaurus_placeholder.set_wrap(True)
        thesaurus_placeholder.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        thesaurus_placeholder.set_margin_top(SPELLING_OUTPUT_PADDING_PX)
        thesaurus_placeholder.set_margin_bottom(SPELLING_OUTPUT_PADDING_PX)
        thesaurus_placeholder.set_margin_start(SPELLING_OUTPUT_PADDING_PX)
        thesaurus_placeholder.set_margin_end(SPELLING_OUTPUT_PADDING_PX)
        thesaurus_placeholder.set_xalign(0)
        thesaurus_list.append(thesaurus_placeholder)
        thesaurus_scroller.set_child(thesaurus_list)
        thesaurus_stack.add_named(thesaurus_scroller, "thesaurus")

        reference_scroller = Gtk.ScrolledWindow()
        reference_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        reference_scroller.set_hexpand(True)
        reference_scroller.set_vexpand(True)
        reference_scroller.set_min_content_height(160)
        reference_scroller.add_css_class("spelling-output-scroller")
        reference_scroller.add_css_class("editor-lookup-page")
        reference_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        reference_box.set_hexpand(True)
        reference_box.set_vexpand(True)
        reference_box.add_css_class("reference-output-box")
        reference_box.set_margin_top(SPELLING_OUTPUT_PADDING_PX)
        reference_box.set_margin_bottom(SPELLING_OUTPUT_PADDING_PX)
        reference_box.set_margin_start(SPELLING_OUTPUT_PADDING_PX)
        reference_box.set_margin_end(SPELLING_OUTPUT_PADDING_PX)
        reference_label = Gtk.Label(label="No reference yet.", xalign=0)
        reference_label.set_wrap(True)
        reference_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        reference_label.set_selectable(True)
        reference_label.set_use_markup(True)
        reference_label.add_css_class("reference-output")
        reference_label.connect("activate-link", self._on_reference_link_clicked)
        reference_box.append(reference_label)
        reference_scroller.set_child(reference_box)
        thesaurus_stack.add_named(reference_scroller, "reference")

        thesaurus_stack.set_visible_child_name("thesaurus")
        thesaurus_section.append(thesaurus_stack)
        self._thesaurus_list = thesaurus_list
        self._thesaurus_placeholder = thesaurus_placeholder
        self._thesaurus_stack = thesaurus_stack
        self._reference_label = reference_label

        split_box.append(output_section)
        split_box.append(thesaurus_section)
        panel.append(split_box)

        return panel

    def _build_text_draft_panel(self) -> Gtk.Box:
        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        panel.set_margin_top(6)
        panel.set_margin_bottom(12)
        panel.set_margin_start(18)
        panel.set_margin_end(18)

        action_wrap = Gtk.FlowBox()
        action_wrap.set_selection_mode(Gtk.SelectionMode.NONE)
        action_wrap.set_column_spacing(6)
        action_wrap.set_row_spacing(6)
        action_wrap.set_max_children_per_line(6)
        action_wrap.set_hexpand(True)
        panel.append(action_wrap)
        self._text_draft_actions_wrap = action_wrap
        self._rebuild_text_draft_action_buttons()

        original_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        original_header_row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        original_header = Gtk.Label(label="Original Output", xalign=0)
        original_header.add_css_class("dim-label")
        original_header.set_hexpand(True)
        original_header.set_halign(Gtk.Align.START)
        original_header_row.append(original_header)
        regenerate_menu_button = Gtk.MenuButton(label="Retry")
        regenerate_menu_button.set_tooltip_text("Try the last Text Draft output again with another model profile.")
        self._apply_quick_action_button_classes(regenerate_menu_button)
        original_header_row.append(regenerate_menu_button)
        self._text_draft_regenerate_menu_button = regenerate_menu_button
        self._rebuild_text_draft_regenerate_profile_chips()
        original_section.append(original_header_row)

        original_surface = Gtk.Stack()
        original_surface.set_hexpand(True)
        original_surface.set_vexpand(True)
        original_surface.add_css_class("editor-lookup-surface")
        original_scroller = Gtk.ScrolledWindow()
        original_scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        original_scroller.set_hexpand(True)
        original_scroller.set_vexpand(True)
        original_scroller.set_min_content_height(150)
        original_scroller.add_css_class("spelling-output-scroller")
        original_scroller.add_css_class("editable-text-scroller")
        original_scroller.add_css_class("editor-lookup-page")
        original_buffer = Gtk.TextBuffer()
        original_view = Gtk.TextView.new_with_buffer(original_buffer)
        original_view.set_monospace(False)
        original_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        original_view.set_editable(True)
        original_view.set_cursor_visible(True)
        original_view.set_vexpand(True)
        original_view.set_hexpand(True)
        original_view.set_size_request(-1, 150)
        original_view.set_left_margin(SPELLING_OUTPUT_PADDING_PX)
        original_view.set_right_margin(SPELLING_OUTPUT_PADDING_PX)
        original_view.set_top_margin(SPELLING_OUTPUT_PADDING_PX)
        original_view.set_bottom_margin(SPELLING_OUTPUT_PADDING_PX)
        original_view.add_css_class("spelling-output-view")
        original_view.add_css_class("editable-text-view")
        original_scroller.set_child(original_view)
        original_surface.add_named(original_scroller, "original")
        original_surface.set_visible_child_name("original")
        original_section.append(original_surface)
        self._text_draft_original_output_buffer = original_buffer

        panel.append(original_section)

        draft_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        draft_section.set_hexpand(True)
        draft_section.set_vexpand(True)
        draft_header_row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        draft_header_row.set_margin_bottom(6)

        templates_button = self._build_text_draft_templates_button()
        draft_header_row.append(templates_button)
        self._text_draft_template_button = templates_button

        case_search_entry = Gtk.Entry()
        case_search_entry.set_placeholder_text("Add case")
        case_search_entry.set_tooltip_text("Search concordance cases to append to this Draft.")
        case_search_entry.set_hexpand(True)
        case_search_entry.set_halign(Gtk.Align.FILL)
        case_search_entry.set_valign(Gtk.Align.CENTER)
        case_search_entry.set_size_request(0, -1)
        case_search_entry.add_css_class("text-draft-case-search")
        case_search_entry.connect("changed", self._on_text_draft_case_search_changed)
        case_search_key_controller = Gtk.EventControllerKey()
        case_search_key_controller.set_propagation_phase(Gtk.PropagationPhase.CAPTURE)
        case_search_key_controller.connect("key-pressed", self._on_text_draft_case_search_key_pressed)
        case_search_entry.add_controller(case_search_key_controller)
        draft_header_row.append(case_search_entry)
        self._text_draft_case_search_entry = case_search_entry

        template_email_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        template_email_box.set_visible(False)
        draft_header_row.append(template_email_box)
        self._text_draft_template_email_box = template_email_box

        terminal_close_btn = Gtk.Button(icon_name="go-previous-symbolic")
        terminal_close_btn.add_css_class("flat")
        terminal_close_btn.add_css_class("transform-pill")
        terminal_close_btn.add_css_class("transform-pill-compact")
        terminal_close_btn.set_tooltip_text("Return to the Draft text box.")
        terminal_close_btn.set_visible(False)
        terminal_close_btn.connect("clicked", self._on_text_draft_terminal_close_clicked)
        draft_header_row.append(terminal_close_btn)
        self._text_draft_terminal_close_button = terminal_close_btn

        external_action_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=0)
        draft_header_row.append(external_action_box)
        self._text_draft_external_action_box = external_action_box
        self._rebuild_text_draft_external_action_buttons()

        copy_draft_btn = Gtk.Button(label="Copy", icon_name="edit-copy-symbolic")
        copy_draft_btn.add_css_class("flat")
        copy_draft_btn.add_css_class("transform-pill")
        copy_draft_btn.add_css_class("transform-pill-compact")
        copy_draft_btn.set_tooltip_text("Copy the full Draft text to the clipboard.")
        copy_draft_btn.connect("clicked", self._on_copy_text_draft_clicked)
        draft_header_row.append(copy_draft_btn)

        draft_section.append(draft_header_row)

        case_results_list = Gtk.ListBox()
        case_results_list.add_css_class("text-draft-case-suggestion-list")
        case_results_list.set_selection_mode(Gtk.SelectionMode.SINGLE)
        case_results_list.connect("row-activated", self._on_text_draft_case_row_activated)
        case_results_scroller = Gtk.ScrolledWindow()
        case_results_scroller.add_css_class("text-draft-case-results")
        case_results_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        case_results_scroller.set_max_content_height(TEXT_DRAFT_CASE_SUGGESTION_MAX_HEIGHT)
        case_results_scroller.set_propagate_natural_height(True)
        case_results_scroller.set_visible(False)
        case_results_scroller.set_child(case_results_list)
        draft_section.append(case_results_scroller)
        self._text_draft_case_results_scroller = case_results_scroller
        self._text_draft_case_list_box = case_results_list

        case_attachments_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        case_attachments_box.set_hexpand(True)
        case_attachments_box.set_visible(False)
        case_attachments_box.add_css_class("text-draft-case-attachments")
        self._text_draft_case_attachments_box = case_attachments_box

        draft_surface = Gtk.Stack()
        draft_surface.set_hexpand(True)
        draft_surface.set_vexpand(True)
        draft_surface.add_css_class("editor-lookup-surface")
        draft_surface.connect(
            "notify::visible-child-name",
            self._on_text_draft_surface_visible_child_name_changed,
        )
        draft_scroller = Gtk.ScrolledWindow()
        draft_scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        draft_scroller.set_hexpand(True)
        draft_scroller.set_vexpand(True)
        draft_scroller.set_min_content_height(260)
        draft_scroller.add_css_class("spelling-output-scroller")
        draft_scroller.add_css_class("editable-text-scroller")
        draft_scroller.add_css_class("editor-lookup-page")
        draft_buffer = Gtk.TextBuffer()
        draft_buffer.connect("changed", self._on_text_draft_buffer_changed)
        draft_view = Gtk.TextView.new_with_buffer(draft_buffer)
        draft_view.set_monospace(False)
        draft_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        draft_view.set_vexpand(True)
        draft_view.set_hexpand(True)
        draft_view.set_size_request(-1, 260)
        draft_view.set_left_margin(SPELLING_OUTPUT_PADDING_PX)
        draft_view.set_right_margin(SPELLING_OUTPUT_PADDING_PX)
        draft_view.set_top_margin(SPELLING_OUTPUT_PADDING_PX)
        draft_view.set_bottom_margin(SPELLING_OUTPUT_PADDING_PX)
        draft_view.add_css_class("spelling-output-view")
        draft_view.add_css_class("editable-text-view")
        draft_key_controller = Gtk.EventControllerKey()
        draft_key_controller.set_propagation_phase(Gtk.PropagationPhase.CAPTURE)
        draft_key_controller.connect("key-pressed", self._on_text_draft_key_pressed)
        draft_view.add_controller(draft_key_controller)
        draft_scroller.set_child(draft_view)
        draft_surface.add_named(draft_scroller, "draft")

        if Vte is not None:
            terminal_frame = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
            terminal_frame.set_hexpand(True)
            terminal_frame.set_vexpand(True)
            terminal_frame.set_size_request(-1, 260)
            terminal_frame.add_css_class("text-draft-terminal-frame")
            terminal_frame.set_overflow(Gtk.Overflow.HIDDEN)
            terminal = Vte.Terminal()
            terminal.set_hexpand(True)
            terminal.set_vexpand(True)
            terminal.add_css_class("text-draft-terminal")
            _apply_prose_terminal_theme(terminal)
            terminal_key_controller = Gtk.EventControllerKey()
            terminal_key_controller.set_propagation_phase(Gtk.PropagationPhase.CAPTURE)
            terminal_key_controller.connect("key-pressed", self._on_text_draft_terminal_key_pressed)
            terminal.add_controller(terminal_key_controller)
            Adw.StyleManager.get_default().connect(
                "notify::dark",
                self._on_text_draft_terminal_style_changed,
            )
            terminal.connect("child-exited", self._on_text_draft_terminal_child_exited)
            terminal_frame.append(terminal)
            draft_surface.add_named(terminal_frame, "terminal")
            self._text_draft_terminal = terminal
        else:
            terminal_missing = Gtk.Label(
                label=(
                    "Embedded terminal support requires GTK4 VTE "
                    "(gir1.2-vte-3.91 and libvte-2.91-gtk4-0)."
                ),
                xalign=0,
            )
            terminal_missing.set_wrap(True)
            terminal_missing.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
            terminal_missing.set_margin_top(SPELLING_OUTPUT_PADDING_PX)
            terminal_missing.set_margin_bottom(SPELLING_OUTPUT_PADDING_PX)
            terminal_missing.set_margin_start(SPELLING_OUTPUT_PADDING_PX)
            terminal_missing.set_margin_end(SPELLING_OUTPUT_PADDING_PX)
            terminal_missing.add_css_class("dim-label")
            draft_surface.add_named(terminal_missing, "terminal-missing")

        draft_surface.set_visible_child_name("draft")
        draft_overlay = Gtk.Overlay()
        draft_overlay.set_hexpand(True)
        draft_overlay.set_vexpand(True)
        draft_overlay.set_child(draft_surface)
        clear_draft_btn = Gtk.Button(icon_name="user-trash-symbolic")
        clear_draft_btn.add_css_class("flat")
        clear_draft_btn.add_css_class("text-draft-clear-draft")
        clear_draft_btn.set_halign(Gtk.Align.END)
        clear_draft_btn.set_valign(Gtk.Align.END)
        clear_draft_btn.set_margin_bottom(8)
        clear_draft_btn.set_margin_end(8)
        clear_draft_btn.set_tooltip_text("Clear Draft.")
        clear_draft_btn.set_visible(False)
        clear_draft_btn.connect("clicked", self._on_clear_text_draft_clicked)
        draft_overlay.add_overlay(clear_draft_btn)
        draft_section.append(draft_overlay)
        draft_section.append(case_attachments_box)
        self._text_draft_buffer = draft_buffer
        self._text_draft_view = draft_view
        self._text_draft_draft_surface = draft_surface
        self._text_draft_clear_button = clear_draft_btn

        self._refresh_text_draft_templates()
        panel.append(draft_section)
        return panel

    def _clear_flow_box(self, flow_box: Gtk.FlowBox) -> None:
        child = flow_box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            flow_box.remove(child)
            child = next_child

    def _clear_list_box(self, list_box: Gtk.ListBox) -> None:
        child = list_box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            list_box.remove(child)
            child = next_child

    def _clear_box(self, box: Gtk.Box) -> None:
        child = box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            box.remove(child)
            child = next_child

    def _model_profile_by_key(self, profile_key: str) -> ModelProfile | None:
        for profile in self._model_profiles:
            if profile.key == profile_key:
                return profile
        return None

    def _model_profile_by_nickname(self, nickname: str) -> ModelProfile | None:
        requested = nickname.strip()
        if not requested:
            return None
        for profile in self._model_profiles:
            if profile.display_name() == requested:
                return profile
        lowered = requested.lower()
        for profile in self._model_profiles:
            if profile.display_name().lower() == lowered:
                return profile
        return None

    def _command_title(self, action_key: str) -> str:
        return PROFILE_BACKED_COMMAND_TITLES.get(action_key, action_key.replace("-", " ").title())

    def _default_profile_key_for_action(self, action_key: str) -> str | None:
        default_key = self._editor_action_profile_defaults.get(action_key)
        return default_key if default_key in MODEL_PROFILE_IDS else None

    def _default_profile_for_action(self, action_key: str) -> ModelProfile | None:
        default_key = self._default_profile_key_for_action(action_key)
        if default_key is None:
            return None
        return self._model_profile_by_key(default_key)

    def _proofread_runtime_profile(self) -> ModelProfile:
        nickname = self._proof_settings.model_id.strip() or "Proof Reading"
        return ModelProfile(
            key="proof",
            nickname=nickname,
            abbreviation="",
            api_url=self._proof_settings.api_url,
            model_id=self._proof_settings.model_id,
            api_key=self._proof_settings.api_key,
            disable_reasoning=self._proof_settings.disable_reasoning,
        )

    def _profile_slot_label(self, profile: ModelProfile) -> str:
        abbreviation = profile.abbreviation.strip()
        if abbreviation:
            return abbreviation
        try:
            return str(MODEL_PROFILE_IDS.index(profile.key) + 1)
        except ValueError:
            return profile.display_name()[:1] or "?"

    def _active_regenerate_context(self) -> RegenerateContext | None:
        return self._pending_regenerate_context or self._last_regenerate_context

    def _configured_regenerate_profiles(self) -> list[ModelProfile]:
        return [profile for profile in self._model_profiles if profile.is_configured()]

    def _regenerate_profile_chip_tooltip(
        self, profile: ModelProfile, context: RegenerateContext | None
    ) -> str:
        if context is None:
            tooltip = "Run a supported text-generation command first."
        else:
            tooltip = f"Regenerate {context.command_title} with {profile.display_name()}."
            if profile.key == self._default_profile_key_for_action(context.action_key):
                tooltip = f"{tooltip}\nDefault {context.command_title} profile."
        if profile.model_id.strip():
            tooltip = f"{tooltip}\nModel: {profile.model_id.strip()}"
        return tooltip

    def _regenerate_profile_chip_sensitive(self) -> bool:
        return not self._busy and self._active_regenerate_context() is not None

    def _build_regenerate_profile_popover(
        self,
        profiles: list[ModelProfile],
        context: RegenerateContext | None,
        tooltip_builder: Callable[[ModelProfile, RegenerateContext | None], str],
        activate_handler: Callable[[Gtk.Button | None, str | None], None],
    ) -> tuple[Gtk.Popover, list[Gtk.Button]]:
        popover = Gtk.Popover()
        popover.set_autohide(True)
        popover.set_cascade_popdown(True)
        popover.set_position(Gtk.PositionType.BOTTOM)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(10)
        content.set_margin_end(10)

        title = Gtk.Label(label="Try again with", xalign=0)
        title.add_css_class("caption")
        title.add_css_class("dim-label")
        content.append(title)

        profile_grid = Gtk.FlowBox()
        profile_grid.set_selection_mode(Gtk.SelectionMode.NONE)
        profile_grid.set_column_spacing(6)
        profile_grid.set_row_spacing(6)
        profile_grid.set_max_children_per_line(2)

        buttons: list[Gtk.Button] = []
        for profile in profiles:
            button = Gtk.Button(label=self._profile_slot_label(profile))
            button.add_css_class("flat")
            button.add_css_class("improve-profile-chip")
            button.set_tooltip_text(tooltip_builder(profile, context))
            button.connect(
                "clicked",
                self._on_regenerate_profile_menu_clicked,
                popover,
                activate_handler,
                profile.display_name(),
            )
            profile_grid.append(button)
            buttons.append(button)

        content.append(profile_grid)
        popover.set_child(content)
        return popover, buttons

    def _on_regenerate_profile_menu_clicked(
        self,
        _button: Gtk.Button,
        popover: Gtk.Popover,
        activate_handler: Callable[[Gtk.Button | None, str | None], None],
        profile_nickname: str,
    ) -> None:
        popover.popdown()
        activate_handler(None, profile_nickname)

    def _rebuild_regenerate_profile_chips(self) -> None:
        menu_button = self._regenerate_menu_button
        if menu_button is None:
            return

        self._regenerate_profile_buttons = []
        profiles = self._configured_regenerate_profiles()
        context = self._active_regenerate_context()
        menu_button.set_visible(bool(profiles))
        menu_button.set_sensitive(self._regenerate_profile_chip_sensitive())
        if not profiles:
            return

        popover, buttons = self._build_regenerate_profile_popover(
            profiles,
            context,
            self._regenerate_profile_chip_tooltip,
            self._on_regenerate_clicked,
        )
        for button in buttons:
            button.set_sensitive(self._regenerate_profile_chip_sensitive())
        menu_button.set_popover(popover)
        self._regenerate_profile_buttons = buttons

    def _make_regenerate_context(self, action_key: str, source_text: str) -> RegenerateContext | None:
        insert_mode = REGENERATE_INSERT_MODE_BY_ACTION.get(action_key)
        if insert_mode is None:
            return None
        return RegenerateContext(
            action_key=action_key,
            command_title=self._command_title(action_key),
            source_text=source_text,
            insert_mode=insert_mode,
        )

    def _set_pending_regenerate_context(self, action_key: str, source_text: str) -> None:
        context = self._make_regenerate_context(action_key, source_text)
        self._pending_regenerate_context = context
        self._rebuild_regenerate_profile_chips()

    def _clear_pending_regenerate_context(self) -> None:
        if self._pending_regenerate_context is None:
            return
        self._pending_regenerate_context = None
        self._rebuild_regenerate_profile_chips()

    def _commit_pending_regenerate_context(self) -> None:
        if self._pending_regenerate_context is None:
            return
        self._last_regenerate_context = self._pending_regenerate_context
        self._pending_regenerate_context = None
        self._rebuild_regenerate_profile_chips()

    def _resolve_profile_for_action(self, action_key: str, profile_nickname: str | None) -> ModelProfile | None:
        if profile_nickname:
            profile = self._model_profile_by_nickname(profile_nickname)
            if profile is None:
                self._show_toast(f'Unknown model profile "{profile_nickname}".')
            return profile
        profile = self._default_profile_for_action(action_key)
        if profile is None:
            self._show_toast(f'Choose a default model profile for "{self._command_title(action_key)}" in Settings.')
        return profile

    def _active_text_draft_regenerate_context(self) -> RegenerateContext | None:
        return self._text_draft_pending_regenerate_context or self._text_draft_last_regenerate_context

    def _text_draft_regenerate_profile_chip_sensitive(self) -> bool:
        return not self._busy and self._active_text_draft_regenerate_context() is not None

    def _text_draft_regenerate_profile_chip_tooltip(
        self, profile: ModelProfile, context: RegenerateContext | None
    ) -> str:
        if context is None:
            tooltip = "Run a supported Text Draft command first."
        else:
            tooltip = f"Regenerate {context.command_title} in Text Draft with {profile.display_name()}."
            if profile.key == self._default_profile_key_for_action(context.action_key):
                tooltip = f"{tooltip}\nDefault {context.command_title} profile."
        if profile.model_id.strip():
            tooltip = f"{tooltip}\nModel: {profile.model_id.strip()}"
        return tooltip

    def _rebuild_text_draft_regenerate_profile_chips(self) -> None:
        menu_button = self._text_draft_regenerate_menu_button
        if menu_button is None:
            return

        self._text_draft_regenerate_profile_buttons = []
        profiles = self._configured_regenerate_profiles()
        context = self._active_text_draft_regenerate_context()
        menu_button.set_visible(bool(profiles))
        menu_button.set_sensitive(self._text_draft_regenerate_profile_chip_sensitive())
        if not profiles:
            return

        popover, buttons = self._build_regenerate_profile_popover(
            profiles,
            context,
            self._text_draft_regenerate_profile_chip_tooltip,
            self._on_text_draft_regenerate_clicked,
        )
        for button in buttons:
            button.set_sensitive(self._text_draft_regenerate_profile_chip_sensitive())
        menu_button.set_popover(popover)
        self._text_draft_regenerate_profile_buttons = buttons

    def _make_text_draft_regenerate_context(self, action_key: str, source_text: str) -> RegenerateContext | None:
        if action_key not in {"improve-generated", "rephrase-generated", "improve-selected"}:
            return None
        return RegenerateContext(
            action_key=action_key,
            command_title=self._command_title(action_key),
            source_text=source_text,
            insert_mode="text-draft",
        )

    def _set_pending_text_draft_regenerate_context(self, action_key: str, source_text: str) -> None:
        self._text_draft_pending_regenerate_context = self._make_text_draft_regenerate_context(action_key, source_text)
        self._rebuild_text_draft_regenerate_profile_chips()

    def _clear_pending_text_draft_regenerate_context(self) -> None:
        if self._text_draft_pending_regenerate_context is None:
            return
        self._text_draft_pending_regenerate_context = None
        self._rebuild_text_draft_regenerate_profile_chips()

    def _commit_pending_text_draft_regenerate_context(self) -> None:
        if self._text_draft_pending_regenerate_context is None:
            return
        self._text_draft_last_regenerate_context = self._text_draft_pending_regenerate_context
        self._text_draft_pending_regenerate_context = None
        self._rebuild_text_draft_regenerate_profile_chips()

    def _build_quick_action_button(
        self,
        definition: QuickActionDefinition,
        *,
        label: str | None = None,
        on_clicked: Callable[[Gtk.Button], None] | None = None,
    ) -> Gtk.Button:
        button = Gtk.Button(label=label or definition.label)
        button.set_action_name(f"app.{definition.action_name}")
        if definition.supports_profiles:
            default_profile = self._default_profile_for_action(definition.key)
            default_target = default_profile.display_name() if default_profile is not None else ""
            button.set_action_target_value(GLib.Variant("s", default_target))
        button.set_tooltip_text(definition.description)
        self._apply_quick_action_button_classes(button)
        if on_clicked is not None:
            button.connect("clicked", on_clicked)
        return button

    def _apply_quick_action_button_classes(
        self,
        widget: Gtk.Widget,
    ) -> None:
        if isinstance(widget, Gtk.Button | Gtk.MenuButton):
            widget.set_has_frame(False)
        widget.add_css_class("flat")
        widget.add_css_class("transform-pill")
        widget.add_css_class("transform-pill-compact")

    def _build_quick_action_widget(
        self,
        definition: QuickActionDefinition,
        *,
        label: str | None = None,
        close_popover: Gtk.Popover | None = None,
    ) -> Gtk.Widget:
        return self._build_quick_action_button(
            definition,
            label=label,
            on_clicked=(lambda _button, current=close_popover: current.popdown()) if close_popover else None,
        )

    def _build_more_actions_button(
        self, definitions: list[QuickActionDefinition]
    ) -> tuple[Gtk.MenuButton, list[Gtk.Widget]]:
        popover = Gtk.Popover()
        popover.set_autohide(True)
        popover.set_cascade_popdown(True)
        popover.set_position(Gtk.PositionType.BOTTOM)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(10)
        content.set_margin_end(10)

        title = Gtk.Label(label="More Actions", xalign=0)
        title.add_css_class("caption")
        title.add_css_class("dim-label")
        content.append(title)

        action_grid = Gtk.FlowBox()
        action_grid.set_selection_mode(Gtk.SelectionMode.NONE)
        action_grid.set_column_spacing(6)
        action_grid.set_row_spacing(6)
        action_grid.set_max_children_per_line(2)

        action_buttons: list[Gtk.Widget] = []
        for definition in definitions:
            widget = self._build_quick_action_widget(definition, label=definition.label, close_popover=popover)
            action_grid.append(widget)
            action_buttons.append(widget)
        content.append(action_grid)

        popover.set_child(content)

        more_button = Gtk.MenuButton(label="More")
        more_button.set_tooltip_text("Show more editor actions")
        more_button.set_popover(popover)
        self._apply_quick_action_button_classes(more_button)
        return more_button, action_buttons

    def _build_draft_choices_button(self) -> tuple[Gtk.MenuButton, list[Gtk.Widget]]:
        definitions = [
            EDITOR_QUICK_ACTION_BY_KEY[key]
            for key in EDITOR_DRAFT_CHOICE_ACTION_KEYS
            if key in EDITOR_QUICK_ACTION_BY_KEY
        ]
        popover = Gtk.Popover()
        popover.set_autohide(True)
        popover.set_cascade_popdown(True)
        popover.set_position(Gtk.PositionType.BOTTOM)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(10)
        content.set_margin_end(10)

        title = Gtk.Label(label="Choices", xalign=0)
        title.add_css_class("caption")
        title.add_css_class("dim-label")
        content.append(title)

        action_grid = Gtk.FlowBox()
        action_grid.set_selection_mode(Gtk.SelectionMode.NONE)
        action_grid.set_column_spacing(6)
        action_grid.set_row_spacing(6)
        action_grid.set_max_children_per_line(2)

        action_buttons: list[Gtk.Widget] = []
        for definition in definitions:
            widget = self._build_quick_action_widget(definition, label=definition.label, close_popover=popover)
            action_grid.append(widget)
            action_buttons.append(widget)
        content.append(action_grid)

        popover.set_child(content)

        draft_choices_button = Gtk.MenuButton(label="Choices")
        draft_choices_button.set_tooltip_text("Show choice generators")
        draft_choices_button.set_popover(popover)
        self._apply_quick_action_button_classes(draft_choices_button)
        return draft_choices_button, action_buttons

    def _rebuild_transform_action_buttons(self) -> None:
        flow_box = self._transform_actions_wrap
        if flow_box is None:
            return

        self._clear_flow_box(flow_box)
        action_buttons: list[Gtk.Widget] = []

        pinned_keys = self._editor_pinned_action_ids
        pinned_set = set(pinned_keys)
        ordered_keys = _ordered_editor_quick_action_keys(pinned_keys)
        remaining_actions: list[QuickActionDefinition] = []

        for key in ordered_keys:
            definition = EDITOR_QUICK_ACTION_BY_KEY[key]
            if key in pinned_set:
                widget = self._build_quick_action_widget(definition)
                flow_box.append(widget)
                action_buttons.append(widget)
            else:
                remaining_actions.append(definition)

        draft_choices_button, draft_choice_buttons = self._build_draft_choices_button()
        flow_box.append(draft_choices_button)
        action_buttons.append(draft_choices_button)
        action_buttons.extend(draft_choice_buttons)

        if remaining_actions:
            more_button, more_action_buttons = self._build_more_actions_button(remaining_actions)
            flow_box.append(more_button)
            action_buttons.append(more_button)
            action_buttons.extend(more_action_buttons)

        self._transform_action_buttons = action_buttons
        for widget in self._transform_action_buttons:
            widget.set_sensitive(not self._busy)

    def _rebuild_text_draft_action_buttons(self) -> None:
        flow_box = self._text_draft_actions_wrap
        if flow_box is None:
            return

        self._clear_flow_box(flow_box)
        action_buttons: list[Gtk.Widget] = []

        pinned_keys = self._text_draft_pinned_action_ids
        pinned_set = set(pinned_keys)
        ordered_keys = _ordered_text_draft_quick_action_keys(pinned_keys)
        remaining_actions: list[QuickActionDefinition] = []

        for key in ordered_keys:
            definition = TEXT_DRAFT_QUICK_ACTION_BY_KEY[key]
            if key in pinned_set:
                widget = self._build_quick_action_widget(definition)
                flow_box.append(widget)
                action_buttons.append(widget)
            else:
                remaining_actions.append(definition)

        if remaining_actions:
            more_button, more_action_buttons = self._build_more_actions_button(remaining_actions)
            flow_box.append(more_button)
            action_buttons.append(more_button)
            action_buttons.extend(more_action_buttons)

        self._text_draft_action_buttons = action_buttons
        for widget in self._text_draft_action_buttons:
            widget.set_sensitive(not self._busy)

    def _rebuild_text_draft_external_action_buttons(self) -> None:
        box = self._text_draft_external_action_box
        if box is None:
            return
        self._clear_box(box)
        self._text_draft_external_action_buttons = []
        for external_action in self._text_draft_external_actions:
            if not external_action.is_configured():
                continue
            display_label = self._text_draft_external_action_display_label(external_action)
            icon_name = self._text_draft_external_action_display_icon(external_action)
            button = Gtk.Button(label=display_label, icon_name=icon_name)
            button.add_css_class("flat")
            button.add_css_class("transform-pill")
            button.add_css_class("transform-pill-compact")
            tooltip = external_action.tooltip.strip() or f"Run {display_label} with the current Draft text."
            button.set_tooltip_text(tooltip)
            button.set_sensitive(not self._busy)
            button.connect(
                "clicked",
                lambda clicked_button, action=external_action: self._on_text_draft_external_action_clicked(
                    clicked_button,
                    action,
                ),
            )
            box.append(button)
            self._text_draft_external_action_buttons.append(button)

    def _ensure_css(self) -> None:
        if self._css_provider is not None:
            return
        css = f"""
.spelling-output-scroller {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: alpha(@window_fg_color, 0.08);
  border: none;
  box-shadow: none;
}}
.spelling-output-scroller > viewport {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: transparent;
  box-shadow: none;
}}
.spelling-output-view {{
  font-size: {SPELLING_OUTPUT_FONT_SIZE_PX}px;
  background-color: transparent;
  background-image: none;
  box-shadow: none;
  color: @window_fg_color;
  caret-color: @window_fg_color;
}}
.spelling-output-view.view,
textview.spelling-output-view,
textview.spelling-output-view.view {{
  background-color: transparent;
  background-image: none;
  box-shadow: none;
  color: @window_fg_color;
  caret-color: @window_fg_color;
}}
.spelling-output-view text {{
  background-color: transparent;
}}
.spelling-output-view.view text,
textview.spelling-output-view text,
textview.spelling-output-view.view text,
.spelling-output-view.view border,
textview.spelling-output-view border,
textview.spelling-output-view.view border {{
  background-color: transparent;
  background-image: none;
  box-shadow: none;
}}
.editor-lookup-surface textview.spelling-output-view > border.top,
.editor-lookup-surface textview.spelling-output-view > border.right,
.editor-lookup-surface textview.spelling-output-view > border.bottom,
.editor-lookup-surface textview.spelling-output-view > border.left,
.editor-lookup-surface textview.spelling-output-view.view > border.top,
.editor-lookup-surface textview.spelling-output-view.view > border.right,
.editor-lookup-surface textview.spelling-output-view.view > border.bottom,
.editor-lookup-surface textview.spelling-output-view.view > border.left,
.editor-lookup-surface textview.spelling-output-view > text,
.editor-lookup-surface textview.spelling-output-view.view > text {{
  background-color: transparent;
  background-image: none;
  box-shadow: none;
}}
.spelling-output-view border,
.spelling-output-view text {{
  background-image: none;
}}
.proof-suggestion-replacement {{
  background-color: transparent;
  background-image: none;
  box-shadow: none;
  color: alpha(@window_fg_color, 0.82);
  caret-color: @window_fg_color;
}}
.proof-suggestion-replacement text,
.proof-suggestion-replacement border {{
  background-color: transparent;
  background-image: none;
  box-shadow: none;
}}
.proof-suggestion-diff-row {{
  background-color: transparent;
}}
.proof-suggestion-diff-marker {{
  border-radius: 7px;
  font-weight: 700;
  min-width: 20px;
  padding: 1px 4px;
}}
.proof-suggestion-diff-marker-minus {{
  background-color: alpha(@error_color, 0.12);
  color: @error_color;
}}
.proof-suggestion-diff-marker-plus {{
  background-color: alpha(@success_color, 0.14);
  color: @success_color;
}}
.proof-suggestion-original-text {{
  color: alpha(@window_fg_color, 0.72);
}}
.thesaurus-results,
.reference-output-box {{
  background-color: transparent;
}}
.thesaurus-row,
.thesaurus-row label {{
  background-color: transparent;
  box-shadow: none;
}}
.thesaurus-row {{
  border-bottom: 1px solid alpha(@borders, 0.55);
}}
.thesaurus-row:last-child {{
  border-bottom: none;
}}
.editable-text-scroller {{
  box-shadow: none;
}}
.editable-text-scroller:focus-within {{
  border: none;
  box-shadow: none;
}}
.editable-text-view,
.editable-text-view text,
.editable-text-view:focus,
.editable-text-view text:focus {{
  box-shadow: none;
  outline: none;
}}
.editor-lookup-surface {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: alpha(@window_fg_color, 0.08);
  border: none;
  box-shadow: none;
}}
.text-draft-terminal-frame {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: @window_bg_color;
  background-image: none;
  border: none;
  box-shadow: none;
}}
.text-draft-terminal {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  padding: 8px;
  background-color: @window_bg_color;
  background-image: none;
  color: @window_fg_color;
}}
.editor-lookup-surface > stackpage {{
  background-color: transparent;
}}
.editor-lookup-page {{
  background-color: transparent;
  border: none;
  box-shadow: none;
}}
.editor-lookup-page > viewport {{
  background-color: transparent;
}}
.editor-lookup-surface .thesaurus-results,
.editor-lookup-surface .reference-output-box {{
  background-color: transparent;
}}
.multi-draft-results {{
  background-color: transparent;
}}
.multi-draft-choice {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: alpha(@window_fg_color, 0.05);
  padding: 10px;
}}
.multi-draft-choice-improve {{
  background-color: rgba(53, 132, 228, 0.08);
}}
.multi-draft-choice-rephrase {{
  background-color: rgba(200, 136, 0, 0.08);
}}
.multi-draft-choice-rephrase-1 {{
  background-color: rgba(200, 136, 0, 0.08);
}}
.multi-draft-choice-rephrase-2 {{
  background-color: rgba(46, 160, 67, 0.08);
}}
.multi-draft-column-header {{
  font-weight: 700;
  color: @window_fg_color;
}}
.multi-draft-variant-badge {{
  border-radius: 7px;
  padding: 2px 7px;
  font-size: 0.78rem;
  font-weight: 700;
}}
.multi-draft-variant-badge-improve {{
  background-color: rgba(53, 132, 228, 0.18);
  color: @window_fg_color;
}}
.multi-draft-variant-badge-rephrase {{
  background-color: rgba(200, 136, 0, 0.18);
  color: @window_fg_color;
}}
.multi-draft-variant-badge-rephrase-1 {{
  background-color: rgba(200, 136, 0, 0.18);
  color: @window_fg_color;
}}
.multi-draft-variant-badge-rephrase-2 {{
  background-color: rgba(46, 160, 67, 0.18);
  color: @window_fg_color;
}}
.multi-draft-choice-output {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  background-color: alpha(@window_fg_color, 0.06);
  box-shadow: none;
}}
.multi-draft-choice-output > viewport {{
  background-color: transparent;
}}
.thesaurus-row label {{
  font-size: {SPELLING_OUTPUT_FONT_SIZE_PX}px;
  color: @window_fg_color;
}}
.reference-toggle-active {{
  background-color: alpha(@window_fg_color, 0.08);
  color: @window_fg_color;
}}
button.transform-pill,
menubutton.transform-pill > button {{
  border-radius: {SPELLING_OUTPUT_CORNER_RADIUS_PX}px;
  padding-left: 14px;
  padding-right: 14px;
}}
button.transform-pill-compact,
menubutton.transform-pill-compact > button {{
  padding-left: 10px;
  padding-right: 10px;
  min-height: 28px;
  font-size: 0.85rem;
}}
button.improve-profile-chip {{
  min-height: 24px;
  min-width: 24px;
  padding-left: 6px;
  padding-right: 6px;
  border-radius: 8px;
  font-size: 0.8rem;
}}
.reference-output {{
  font-size: {REFERENCE_OUTPUT_FONT_SIZE_PX}px;
  color: @window_fg_color;
}}
.reference-query-entry {{
  font-size: {SPELLING_OUTPUT_FONT_SIZE_PX}px;
}}
.text-draft-case-suggestion-list {{
  min-width: 260px;
  padding: 2px;
  background-color: transparent;
}}
.text-draft-case-results {{
  background-color: transparent;
  border: none;
  box-shadow: none;
}}
.text-draft-case-results > viewport {{
  background-color: transparent;
}}
.text-draft-case-suggestion-row {{
  border-radius: 4px;
  padding: 3px 6px;
}}
.text-draft-case-suggestion-row label {{
  font-size: 0.92rem;
}}
.text-draft-case-search {{
  font-size: {SPELLING_OUTPUT_FONT_SIZE_PX}px;
}}
button.text-draft-clear-draft {{
  min-height: 28px;
  min-width: 28px;
  padding: 4px;
  border-radius: 999px;
  opacity: 0.72;
  background-color: alpha(@window_bg_color, 0.86);
}}
button.text-draft-clear-draft:hover {{
  opacity: 1;
  background-color: alpha(@window_fg_color, 0.12);
}}
.text-draft-case-attachments {{
  padding: 0;
}}
.text-draft-case-attachment {{
  border-radius: 8px;
  background-color: alpha(@window_fg_color, 0.08);
  padding: 4px 6px 4px 10px;
}}
button.text-draft-case-remove {{
  min-height: 20px;
  min-width: 20px;
  padding: 0;
  border-radius: 6px;
}}
"""
        provider = Gtk.CssProvider()
        provider.load_from_data(css.encode("utf-8"))
        display = Gdk.Display.get_default()
        if display:
            Gtk.StyleContext.add_provider_for_display(
                display,
                provider,
                Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION,
            )
            self._css_provider = provider

    def _build_proof_panel(self) -> Gtk.Box:
        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        panel.set_margin_top(6)
        panel.set_margin_bottom(12)
        panel.set_margin_start(18)
        panel.set_margin_end(18)

        page_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        page_box.set_hexpand(True)
        self._start_spin = Gtk.SpinButton.new_with_range(1, 9999, 1)
        self._start_spin.set_value(1)
        self._end_spin = Gtk.SpinButton.new_with_range(1, 9999, 1)
        self._end_spin.set_value(4)
        page_box.append(Gtk.Label(label="Page range:"))
        page_box.append(self._start_spin)
        page_box.append(Gtk.Label(label="to"))
        page_box.append(self._end_spin)
        panel.append(page_box)

        document_scope_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        self._entire_document_check = Gtk.CheckButton(label="Entire document")
        self._entire_document_check.connect("toggled", self._on_proof_entire_document_toggled)
        document_scope_box.append(self._entire_document_check)

        run_btn = Gtk.Button(icon_name="media-playback-start-symbolic")
        run_btn.add_css_class("suggested-action")
        run_btn.add_css_class("flat")
        run_btn.set_action_name("app.request-changes")
        run_btn.set_tooltip_text("Request changes")
        self._run_btn = run_btn
        document_scope_box.append(run_btn)

        load_json_btn = Gtk.Button(icon_name="document-open-symbolic")
        load_json_btn.add_css_class("flat")
        load_json_btn.set_action_name("app.load-suggestions-json")
        load_json_btn.set_tooltip_text("Load suggestions JSON")
        self._load_json_btn = load_json_btn
        document_scope_box.append(load_json_btn)
        panel.append(document_scope_box)

        sep = Gtk.Separator(orientation=Gtk.Orientation.HORIZONTAL)
        panel.append(sep)

        list_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        scroller.set_hexpand(True)
        scroller.set_vexpand(True)  # fill remaining vertical space for suggestions
        scroller.set_child(list_box)
        self._list_box = list_box

        info_label = Gtk.Label(label="Proof reading suggestions will appear here after Prose finishes batching the document.")
        info_label.set_wrap(True)
        info_label.add_css_class("dim-label")
        self._empty_label = info_label
        list_box.append(info_label)

        panel.append(scroller)
        return panel

    def _build_prefixes_panel(self) -> Gtk.Widget:
        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        panel.set_margin_top(6)
        panel.set_margin_bottom(12)
        panel.set_margin_start(18)
        panel.set_margin_end(18)

        intro = Gtk.Label(
            label=(
                "Manage the prefixes used by Input RT/Input CT and the temporary word substitutions "
                "applied before direct input or SpellingStyle sends text to the first model."
            ),
            xalign=0,
        )
        intro.add_css_class("dim-label")
        intro.set_wrap(True)
        intro.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        panel.append(intro)

        prefixes_group = Adw.PreferencesGroup(title="Citation Prefixes")
        prefixes_group.add_css_class("list-stack")
        panel.append(prefixes_group)

        rt_row = Adw.PreferencesRow()
        rt_row.set_selectable(False)
        rt_row.set_activatable(False)
        rt_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        rt_box.set_margin_top(10)
        rt_box.set_margin_bottom(10)
        rt_box.set_margin_start(12)
        rt_box.set_margin_end(12)
        rt_label = Gtk.Label(label="RT prefix", xalign=0)
        rt_label.set_size_request(90, -1)
        rt_box.append(rt_label)
        rt_entry = Gtk.Entry()
        rt_entry.set_hexpand(True)
        rt_entry.set_text(self._prefix_settings.rt_prefix)
        rt_box.append(rt_entry)
        rt_row.set_child(rt_box)
        prefixes_group.add(rt_row)
        self._rt_prefix_entry = rt_entry

        ct_row = Adw.PreferencesRow()
        ct_row.set_selectable(False)
        ct_row.set_activatable(False)
        ct_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        ct_box.set_margin_top(10)
        ct_box.set_margin_bottom(10)
        ct_box.set_margin_start(12)
        ct_box.set_margin_end(12)
        ct_label = Gtk.Label(label="CT prefix", xalign=0)
        ct_label.set_size_request(90, -1)
        ct_box.append(ct_label)
        ct_entry = Gtk.Entry()
        ct_entry.set_hexpand(True)
        ct_entry.set_text(self._prefix_settings.ct_prefix)
        ct_box.append(ct_entry)
        ct_row.set_child(ct_box)
        prefixes_group.add(ct_row)
        self._ct_prefix_entry = ct_entry

        substitutions_group = Adw.PreferencesGroup(
            title="Word Substitutions",
            description="Any non-empty pair replaces every match before direct input or SpellingStyle runs.",
        )
        substitutions_group.add_css_class("list-stack")
        panel.append(substitutions_group)

        self._substitution_rows = []
        for index in range(MAX_WORD_SUBSTITUTIONS):
            current = self._prefix_settings.substitutions[index]
            row = Adw.PreferencesRow()
            row.set_selectable(False)
            row.set_activatable(False)

            row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
            row_box.set_margin_top(10)
            row_box.set_margin_bottom(10)
            row_box.set_margin_start(12)
            row_box.set_margin_end(12)

            label = Gtk.Label(label=f"Pair {index + 1}", xalign=0)
            label.set_size_request(56, -1)
            row_box.append(label)

            original_entry = Gtk.Entry()
            original_entry.set_hexpand(True)
            original_entry.set_placeholder_text("Original")
            original_entry.set_text(current.original)
            row_box.append(original_entry)

            arrow = Gtk.Label(label="->", xalign=0.5)
            arrow.add_css_class("dim-label")
            row_box.append(arrow)

            replacement_entry = Gtk.Entry()
            replacement_entry.set_hexpand(True)
            replacement_entry.set_placeholder_text("Replacement")
            replacement_entry.set_text(current.replacement)
            row_box.append(replacement_entry)

            row.set_child(row_box)
            substitutions_group.add(row)
            self._substitution_rows.append((original_entry, replacement_entry))

        buttons = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        buttons.set_halign(Gtk.Align.END)
        save_btn = Gtk.Button(label="Save Prefixes")
        save_btn.add_css_class("suggested-action")
        save_btn.add_css_class("flat")
        save_btn.set_action_name("app.save-prefixes")
        buttons.append(save_btn)
        panel.append(buttons)

        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        scroller.set_hexpand(True)
        scroller.set_vexpand(True)
        scroller.set_child(panel)
        return scroller

    def _ensure_menu(self) -> None:
        menu = Gio.Menu()
        menu.append("Prompt Audit", "app.view-last-editor-prompt")
        menu.append("Response Audit", "app.view-last-model-response")
        menu.append("Settings", "app.settings")
        menu.append("Keyboard Shortcuts", "app.show-shortcuts")
        self._menu_button.set_menu_model(menu)

        action_settings = Gio.SimpleAction.new("settings", None)
        action_settings.connect("activate", self._on_open_settings)
        self.get_application().add_action(action_settings)

    def _register_actions(self) -> None:
        app = self.get_application()

        def _add_action(name: str, handler: Callable[[], None]) -> None:
            action = Gio.SimpleAction.new(name, None)
            def _activate(_action: Gio.SimpleAction, _param: GLib.Variant | None) -> None:
                self._show_action_view(name)
                handler()

            action.connect("activate", _activate)
            app.add_action(action)

        def _add_string_action(name: str, handler: Callable[[str], None]) -> None:
            action = Gio.SimpleAction.new(name, GLib.VariantType.new("s"))
            def _activate(_action: Gio.SimpleAction, param: GLib.Variant | None) -> None:
                self._show_action_view(name)
                handler(param.get_string() if param else "")

            action.connect("activate", _activate)
            app.add_action(action)

        _add_action("launch-writer", lambda: self._on_launch_clicked(None))
        _add_action("open-last-odt", lambda: self._on_open_last_odt_clicked(None))
        _add_action("choose-source-file", lambda: self._on_choose_source_file(None))
        _add_action("editor-commands", lambda: self._on_open_editor_commands())
        _add_action("direct-input", lambda: self._on_direct_input_clicked(None))
        _add_string_action("paste-clean-italics", self._on_paste_clean_italics_action)
        _add_action("input-rt", lambda: self._on_input_rt_clicked(None))
        _add_action("input-ct", lambda: self._on_input_ct_clicked(None))
        _add_action("speech-find", lambda: self._on_speech_find_clicked(None))
        _add_action(
            "direct-input-no-trailing-space",
            lambda: self._on_direct_input_no_trailing_space_clicked(None),
        )
        _add_action("save-prefixes", lambda: self._on_save_prefixes_clicked(None))
        _add_action("combine-cites", lambda: self._on_combine_cites_clicked(None))
        _add_action("spellingstyle", lambda: self._on_spellingstyle_clicked(None))
        _add_action("text-draft-spellingstyle", lambda: self._on_text_draft_spellingstyle_clicked(None))
        _add_string_action("improve-generated", lambda nickname: self._on_improve_clicked(None, nickname))
        _add_string_action("rephrase-generated", lambda nickname: self._on_rephrase_generated_clicked(None, nickname))
        _add_string_action("text-draft-improve-generated", lambda nickname: self._on_text_draft_improve_clicked(None, nickname))
        _add_action("text-draft-generated-choices", lambda: self._on_text_draft_generated_choices_clicked(None))
        _add_string_action(
            "text-draft-rephrase-generated",
            lambda nickname: self._on_text_draft_rephrase_generated_clicked(None, nickname),
        )
        _add_string_action("improve", lambda nickname: self._on_improve_clicked(None, nickname))
        _add_action(
            "improve1",
            lambda: self._on_improve_clicked(
                None,
                (self._model_profile_by_key("profile1").display_name() if self._model_profile_by_key("profile1") else None),
            ),
        )
        _add_action(
            "improve2",
            lambda: self._on_improve_clicked(
                None,
                (self._model_profile_by_key("profile2").display_name() if self._model_profile_by_key("profile2") else None),
            ),
        )
        _add_string_action("improve-selected", lambda nickname: self._on_improve_selected_clicked(None, nickname))
        _add_string_action(
            "text-draft-improve-selected",
            lambda nickname: self._on_text_draft_improve_selected_clicked(None, nickname),
        )
        _add_action("text-draft-selected-choices", lambda: self._on_text_draft_selected_choices_clicked(None))
        _add_action("keep-original", lambda: self._on_keep_original_clicked(None))
        _add_action("text-draft-keep-original", lambda: self._on_text_draft_keep_original_clicked(None))
        _add_action("text-draft-wrap-quotes", lambda: self._on_text_draft_wrap_quotes_clicked(None))
        _add_action("text-draft-copy", lambda: self._on_copy_text_draft_clicked(None))
        _add_action("text-draft-external-action", lambda: self._on_text_draft_external_action_clicked(None))
        _add_action("reference-lookup", lambda: self._on_reference_clicked(None))
        _add_string_action("transform-shorten", lambda nickname: self._on_shorten_clicked(None, nickname))
        _add_action("transform-wrap-quotes", lambda: self._on_wrap_quotes_clicked(None))
        _add_action("transform-translate", lambda: self._on_translate_clicked(None))
        _add_string_action("transform-topic-sentence", lambda nickname: self._on_topic_sentence_clicked(None, nickname))
        _add_string_action("transform-introduction", lambda nickname: self._on_introduction_clicked(None, nickname))
        _add_action("transform-introduction-choices", lambda: self._on_introduction_choices_clicked(None))
        _add_action("generated-choices", lambda: self._on_generated_choices_clicked(None))
        _add_action("selected-choices", lambda: self._on_selected_choices_clicked(None))
        _add_string_action(
            "transform-introduction-reply",
            lambda nickname: self._on_introduction_reply_clicked(None, nickname),
        )
        _add_action("transform-introduction-reply-choices", lambda: self._on_introduction_reply_choices_clicked(None))
        _add_string_action("transform-conclusion", lambda nickname: self._on_conclusion_clicked(None, nickname))
        _add_action("transform-conclusion-choices", lambda: self._on_conclusion_choices_clicked(None))
        _add_string_action(
            "transform-concl-no-issues",
            lambda nickname: self._on_concl_no_issues_clicked(None, nickname),
        )
        _add_action("transform-concl-no-issues-choices", lambda: self._on_concl_no_issues_choices_clicked(None))
        _add_string_action("transform-concl-section", lambda nickname: self._on_concl_section_clicked(None, nickname))
        _add_action("transform-concl-section-choices", lambda: self._on_concl_section_choices_clicked(None))
        _add_action("add-case", lambda: self._on_add_case_clicked(None))
        _add_action("import-socf", lambda: self._on_import_socf_clicked(None))
        _add_action("request-changes", lambda: self._on_request_clicked(None))
        _add_action("load-suggestions-json", lambda: self._on_load_suggestions_json_clicked(None))
        _add_action("view-last-editor-prompt", lambda: self._on_view_editor_prompt_clicked(None))
        _add_action("view-last-model-response", lambda: self._on_view_model_response_clicked(None))
        prompt_audit_action = app.lookup_action("view-last-editor-prompt")
        if prompt_audit_action is not None:
            prompt_audit_action.set_enabled(False)
        response_audit_action = app.lookup_action("view-last-model-response")
        if response_audit_action is not None:
            response_audit_action.set_enabled(False)
        _add_action("focus-ask", self._on_focus_ask)
        _add_action("show-shortcuts", self._on_show_shortcuts)
        _add_action("save-settings", self._on_action_save_settings)

        goto_action = Gio.SimpleAction.new("suggestion-goto", GLib.VariantType.new("i"))
        goto_action.connect("activate", self._on_action_suggestion_goto)
        app.add_action(goto_action)

        accept_action = Gio.SimpleAction.new("suggestion-accept", GLib.VariantType.new("i"))
        accept_action.connect("activate", self._on_action_suggestion_accept)
        app.add_action(accept_action)

        reject_action = Gio.SimpleAction.new("suggestion-reject", GLib.VariantType.new("i"))
        reject_action.connect("activate", self._on_action_suggestion_reject)
        app.add_action(reject_action)

        app.set_accels_for_action("app.focus-ask", ["<Ctrl><Shift>Q"])
        app.set_accels_for_action("app.show-shortcuts", ["F1"])

    def _build_shortcuts_window(self) -> Gtk.ShortcutsWindow:
        if self._shortcuts_window is not None:
            return self._shortcuts_window

        window = Gtk.ShortcutsWindow(
            transient_for=self,
            modal=False,
            hide_on_close=True,
            title=f"{APP_NAME} Keyboard Shortcuts",
        )
        window.set_default_size(760, 420)

        section = Gtk.ShortcutsSection(title="Keyboard Shortcuts")

        editor_group = Gtk.ShortcutsGroup(title="Editor")
        editor_group.append(
            Gtk.ShortcutsShortcut(title="Focus Ask field", accelerator="<Primary><Shift>Q")
        )
        section.append(editor_group)

        terminal_group = Gtk.ShortcutsGroup(title="Text Draft Terminal")
        terminal_group.append(
            Gtk.ShortcutsShortcut(title="Copy terminal selection", accelerator="<Primary><Shift>C")
        )
        terminal_group.append(
            Gtk.ShortcutsShortcut(title="Paste into terminal", accelerator="<Primary><Shift>V")
        )
        section.append(terminal_group)

        help_group = Gtk.ShortcutsGroup(title="Reference")
        help_group.append(Gtk.ShortcutsShortcut(title="Show keyboard shortcuts", accelerator="F1"))
        section.append(help_group)

        window.add_section(section)
        self._shortcuts_window = window
        return window

    def _on_show_shortcuts(self) -> None:
        window = self._build_shortcuts_window()
        window.set_transient_for(self)
        window.present()

    def _show_action_view(self, action_name: str) -> None:
        view_name = _main_view_name_for_action(action_name)
        if view_name is None:
            return
        stack = getattr(self, "_main_stack", None)
        if stack is None:
            return
        if stack.get_visible_child_name() != view_name:
            stack.set_visible_child_name(view_name)

    def _on_main_stack_visible_child_name_changed(self, stack: Adw.ViewStack, *_args: object) -> None:
        if stack.get_visible_child_name() == "proof":
            self._try_auto_load_suggestions_json()

    def _try_auto_load_suggestions_json(self) -> None:
        if self._busy:
            return
        path = self._proof_settings.suggestions_json_path
        if path is None:
            return
        path = path.expanduser().resolve(strict=False)
        if path.suffix.lower() != ".json":
            return
        try:
            stat = path.stat()
        except FileNotFoundError:
            return
        except OSError:
            return
        signature = (str(path), int(stat.st_mtime_ns), int(stat.st_size))
        if signature == self._auto_suggestions_json_attempt_signature:
            return
        self._auto_suggestions_json_attempt_signature = signature
        self._load_suggestions_json_path(path, auto=True)

    def _load_suggestions_json_path(self, path: Path, *, auto: bool = False) -> bool:
        if path.suffix.lower() != ".json":
            self._show_toast("Choose a .json file.")
            return False
        if not path.exists():
            if not auto:
                self._show_toast("Selected JSON file no longer exists.")
            return False
        try:
            raw = path.read_text(encoding="utf-8")
        except OSError as exc:
            self._status_label.set_label(f"Unable to read {path.name}: {exc}")
            self._show_toast("Unable to read suggestions JSON.")
            return False
        try:
            suggestions = self._dedupe_suggestions(self._parse_suggestions(raw, allow_empty=True))
        except Exception as exc:  # noqa: BLE001
            self._status_label.set_label(f"Unable to load {path.name}: {exc}")
            self._show_toast("Invalid suggestions JSON.")
            return False
        self._suggestions = self._sort_suggestions(suggestions)
        if not self._suggestions:
            self._empty_label.set_label("No usable suggestions were found in the selected JSON file.")
        else:
            self._empty_label.set_label("Proof reading suggestions will appear here after Prose finishes batching the document.")
        self._render_suggestions()
        count = len(self._suggestions)
        label = "suggestion" if count == 1 else "suggestions"
        verb = "Auto-loaded" if auto else "Loaded"
        self._status_label.set_label(f"{verb} {count} {label} from {path.name}.")
        self._show_toast(f"{verb} {count} {label}.")
        return True


    # Actions ------------------------------------------------------------
    def _on_open_settings(self, *_args: object) -> None:
        if self._settings_window:
            self._settings_window.present()
            return
        win = SettingsWindow(
            self,
            self._model_profiles,
            self._proof_settings,
            self._spelling_settings,
            self._text_draft_spelling_prompt,
            self._citation_validator_settings,
            self._improve1_settings,
            self._improve2_settings,
            self._combine_cites_settings,
            self._thesaurus_settings,
            self._reference_settings,
            self._ask_settings,
            self._shorten_settings,
            self._introduction_settings,
            self._introduction_reply_settings,
            self._conclusion_settings,
            self._concl_no_issues_settings,
            self._topic_sentence_settings,
            self._concl_section_settings,
            self._translate_settings,
            self._shared_style_rules,
            self._editor_action_profile_defaults,
            self._editor_pinned_action_ids,
            self._text_draft_pinned_action_ids,
            self._text_draft_template_dir,
            self._text_draft_external_actions,
            self._libreoffice_python_path,
            self._concordance_file_path,
            self._editor_source_file,
            self._on_editor_source_file_updated,
            self._copy_normal_profile_to_prose,
            self._on_settings_saved,
        )
        win.connect("close-request", self._on_settings_closed)
        self._settings_window = win
        win.present()

    def _on_open_editor_commands(self) -> None:
        if not self._editor_commands_window:
            win = EditorCommandsWindow(self)
            win.connect("close-request", self._on_editor_commands_closed)
            self._editor_commands_window = win
        self._editor_commands_window.present()

    def _on_settings_saved(
        self,
        model_profiles: list[ModelProfile],
        proof_settings: ProofreadSettings,
        spelling_settings: SpellingStyleSettings,
        text_draft_spelling_prompt: str,
        citation_validator_settings: CitationValidatorSettings,
        improve1_settings: Improve1Settings,
        improve2_settings: Improve2Settings,
        combine_cites_settings: CombineCitesSettings,
        thesaurus_settings: ThesaurusSettings,
        reference_settings: ReferenceSettings,
        ask_settings: AskSettings,
        shorten_settings: ShortenSettings,
        introduction_settings: IntroductionSettings,
        introduction_reply_settings: IntroductionSettings,
        conclusion_settings: ConclusionSettings,
        concl_no_issues_settings: ConclNoIssuesSettings,
        topic_sentence_settings: TopicSentenceSettings,
        concl_section_settings: ConclSectionSettings,
        translate_settings: TranslateSettings,
        shared_style_rules: str,
        editor_action_profile_defaults: dict[str, str | None],
        editor_pinned_action_ids: list[str],
        text_draft_pinned_action_ids: list[str],
        text_draft_template_dir: Path | None,
        text_draft_external_actions: list[TextDraftExternalAction],
        libreoffice_python_path: Path | None,
        concordance_file_path: Path | None,
    ) -> None:
        old_suggestions_json_path = self._proof_settings.suggestions_json_path
        self._model_profiles = model_profiles
        self._proof_settings = proof_settings
        self._spelling_settings = spelling_settings
        self._text_draft_spelling_prompt = text_draft_spelling_prompt.strip() or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        self._citation_validator_settings = citation_validator_settings
        self._improve1_settings = improve1_settings
        self._improve2_settings = improve2_settings
        self._combine_cites_settings = combine_cites_settings
        self._thesaurus_settings = thesaurus_settings
        self._reference_settings = reference_settings
        self._ask_settings = ask_settings
        self._shorten_settings = shorten_settings
        self._introduction_settings = introduction_settings
        self._introduction_reply_settings = introduction_reply_settings
        self._conclusion_settings = conclusion_settings
        self._concl_no_issues_settings = concl_no_issues_settings
        self._topic_sentence_settings = topic_sentence_settings
        self._concl_section_settings = concl_section_settings
        self._translate_settings = translate_settings
        self._shared_style_rules = str(shared_style_rules or "").strip()
        self._editor_action_profile_defaults = _sanitize_editor_action_profile_defaults(editor_action_profile_defaults)
        self._editor_pinned_action_ids = _sanitize_editor_pinned_actions(editor_pinned_action_ids)
        self._text_draft_pinned_action_ids = _sanitize_text_draft_pinned_actions(text_draft_pinned_action_ids)
        self._text_draft_template_dir = (
            text_draft_template_dir.expanduser().resolve(strict=False) if text_draft_template_dir else None
        )
        self._text_draft_external_actions = text_draft_external_actions
        self._libreoffice_python_path = libreoffice_python_path.expanduser().resolve(strict=False) if libreoffice_python_path else None
        self._concordance_file_path = (
            concordance_file_path.expanduser().resolve(strict=False) if concordance_file_path else None
        )
        self._load_text_draft_case_index(force=True)
        save_model_profiles(model_profiles)
        save_proofread_settings(proof_settings)
        save_spellingstyle_settings(spelling_settings)
        save_text_draft_spellingstyle_prompt(self._text_draft_spelling_prompt)
        save_citation_validator_settings(citation_validator_settings)
        save_improve1_settings(improve1_settings)
        save_improve2_settings(improve2_settings)
        save_combine_cites_settings(combine_cites_settings)
        save_thesaurus_settings(thesaurus_settings)
        save_reference_settings(reference_settings)
        save_ask_settings(ask_settings)
        save_shorten_settings(shorten_settings)
        save_introduction_settings(introduction_settings)
        save_introduction_reply_settings(introduction_reply_settings)
        save_conclusion_settings(conclusion_settings)
        save_concl_no_issues_settings(concl_no_issues_settings)
        save_topic_sentence_settings(topic_sentence_settings)
        save_concl_section_settings(concl_section_settings)
        save_translate_settings(translate_settings)
        save_shared_style_rules(self._shared_style_rules)
        save_editor_action_profile_defaults(self._editor_action_profile_defaults)
        save_editor_pinned_actions(self._editor_pinned_action_ids)
        save_text_draft_pinned_actions(self._text_draft_pinned_action_ids)
        save_text_draft_template_dir(self._text_draft_template_dir)
        save_text_draft_external_actions(self._text_draft_external_actions)
        save_libreoffice_python_path(self._libreoffice_python_path)
        save_concordance_file_path(self._concordance_file_path)
        self._rebuild_transform_action_buttons()
        self._rebuild_text_draft_action_buttons()
        self._rebuild_text_draft_external_action_buttons()
        self._rebuild_regenerate_profile_chips()
        self._refresh_text_draft_templates()
        _import_uno_from_candidates(self._libreoffice_python_path, force_retry=True)
        self._ctx = None
        self._desktop = None
        self._update_uno_status()
        if old_suggestions_json_path != self._proof_settings.suggestions_json_path:
            self._auto_suggestions_json_attempt_signature = None
            if getattr(self, "_main_stack", None) is not None and self._main_stack.get_visible_child_name() == "proof":
                self._try_auto_load_suggestions_json()

    def _on_settings_closed(self, _window: Gtk.Window) -> bool:
        self._settings_window = None
        return False

    def _on_editor_commands_closed(self, _window: Gtk.Window) -> bool:
        self._editor_commands_window = None
        return False

    def _on_launch_clicked(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose a Writer document")
        dialog.open(self, None, self._on_file_chosen)

    def _on_open_last_odt_clicked(self, _button: Gtk.Button) -> None:
        path = self._last_odt_path
        if not path:
            self._show_toast("No recent Writer document saved.")
            return
        if not path.exists():
            self._show_toast("Last Writer document no longer exists.")
            self._set_last_odt_path(None)
            return
        self._current_doc_path = path
        self._launch_writer_document(path)

    def _on_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._current_doc_path = path
        self._launch_writer_document(path)

    def show_open_error(self, message: str) -> None:
        self._status_label.set_label(message)
        self._show_toast(message)

    def open_odt_path(self, path: Path) -> None:
        path = path.expanduser().resolve(strict=False) if path else path
        if not path:
            self.show_open_error("Choose a local .odt file.")
            return
        if path.suffix.lower() != ".odt":
            self.show_open_error("Choose an .odt file.")
            return
        if not path.exists():
            self.show_open_error(f"Writer document not found: {path}")
            return
        self._current_doc_path = path
        self._launch_writer_document(path)

    def _on_choose_source_file(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose a source text file")
        dialog.open(self, None, self._on_source_file_chosen)

    def _on_source_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_editor_source_file(path)

    def _on_import_socf_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        dialog = Gtk.FileDialog(title="Choose SOCF ODT to import")
        dialog.open(self, None, self._on_import_socf_file_chosen)

    def _on_import_socf_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            source_path = Path(file.get_path() or "")
        except Exception:
            return
        if not source_path:
            return
        if source_path.suffix.lower() != ".odt":
            self._show_toast("Choose an .odt file.")
            return
        if not source_path.exists():
            self._show_toast("Selected SOCF file no longer exists.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        self._ensure_writer_frame_active(desktop, doc)
        self._set_busy(True)
        self._status_label.set_label(f"Importing {source_path.name} without page-number resets…")
        temp_path: Path | None = None
        try:
            temp_path = self._create_page_safe_odt_import_copy(source_path)
            self._insert_odt_at_current_cursor(doc, temp_path)
        except Exception as exc:  # noqa: BLE001
            self._set_busy(False)
            self._status_label.set_label("SOCF import failed.")
            self._show_toast(f"SOCF import failed: {exc}")
            return
        finally:
            if temp_path is not None:
                shutil.rmtree(temp_path.parent, ignore_errors=True)
        self._set_busy(False)
        self._status_label.set_label(f"Imported {source_path.name} without page-number reset metadata.")
        self._show_toast(f"Imported {source_path.name}.")

    def _on_editor_source_file_updated(self, path: Path | None) -> None:
        self._set_editor_source_file(path)

    def _set_editor_source_file(self, path: Path | None) -> None:
        self._editor_source_file = path.expanduser().resolve(strict=False) if path else None
        if self._settings_window:
            self._settings_window.set_source_file(self._editor_source_file)
        save_editor_source_file(self._editor_source_file)

    def _copy_normal_profile_to_prose(self) -> tuple[bool, str]:
        source = _normal_libreoffice_profile_path()
        target = LIBREOFFICE_PROFILE.expanduser().resolve(strict=False)
        if not source.exists():
            return False, f"Normal LibreOffice profile not found: {source}"
        if not (source / "user").exists():
            return False, f"LibreOffice profile is missing its user directory: {source}"
        try:
            same_profile = source.samefile(target)
        except OSError:
            same_profile = source == target
        if same_profile:
            return False, "Normal and Prose LibreOffice profiles are the same path."
        stopped, stop_message = self._stop_listener_for_profile_copy()
        if not stopped:
            return False, stop_message
        backup_path: Path | None = None
        try:
            if target.exists():
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                backup_path = target.with_name(f"{target.name}.backup-{stamp}")
                shutil.move(str(target), str(backup_path))
            shutil.copytree(source, target)
            self._ctx = None
            self._desktop = None
            self._active_doc = None
            return True, "Copied normal LibreOffice profile into Prose."
        except Exception as exc:  # noqa: BLE001
            try:
                if target.exists():
                    shutil.rmtree(target)
                if backup_path and backup_path.exists():
                    shutil.move(str(backup_path), str(target))
            except Exception:
                pass
            return False, f"Unable to copy LibreOffice profile: {exc}"

    def _stop_listener_for_profile_copy(self, timeout: float = 5.0) -> tuple[bool, str]:
        self._ctx = None
        self._desktop = None
        self._active_doc = None
        proc = self._listener_proc
        if not proc or proc.poll() is not None:
            self._listener_proc = None
            return True, ""
        try:
            proc.terminate()
            proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            return False, "Close the Prose LibreOffice window, then try profile copy again."
        except Exception as exc:  # noqa: BLE001
            return False, f"Unable to stop LibreOffice before copying profile: {exc}"
        self._listener_proc = None
        return True, ""

    def _on_request_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        if not self._proof_settings.is_configured():
            self._show_toast("Add Proof Reading API URL, API key, and prompt in Settings. Model ID may be required.")
            return
        profile = self._proofread_runtime_profile()
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        entire_document = self._entire_document_check.get_active()
        start: int | None = None
        end: int | None = None
        if not entire_document:
            start = int(self._start_spin.get_value())
            end = int(self._end_spin.get_value())
            if end < start:
                end = start
                self._end_spin.set_value(start)
        scope_label = self._proofreading_scope_label(start, end)
        self._set_busy(True)
        self._status_label.set_label(f"Proofreading {scope_label} with {profile.display_name()}…")
        thread = threading.Thread(target=self._gather_and_request, args=(doc, start, end, profile), daemon=True)
        thread.start()

    def _on_load_suggestions_json_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        dialog = Gtk.FileDialog(title="Choose suggestions JSON")
        dialog.open(self, None, self._on_suggestions_json_chosen)

    def _on_suggestions_json_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._load_suggestions_json_path(path)

    def _on_proof_entire_document_toggled(self, button: Gtk.CheckButton) -> None:
        enabled = not button.get_active()
        self._start_spin.set_sensitive(enabled)
        self._end_spin.set_sensitive(enabled)

    def _on_spellingstyle_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("spelling", None)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        if profile.api_url.rstrip("/").endswith("/responses"):
            self._show_toast("SpellingStyle uses a chat endpoint. Update the API URL in Settings.")
            return
        source_text = self._read_editor_source_text()
        if source_text is None:
            return
        source_text = self._apply_word_substitutions(source_text)
        if not source_text.strip():
            self._show_toast("Source file is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        GLib.idle_add(self._set_spelling_output_text, "")
        self._set_busy(True)
        self._status_label.set_label("Streaming SpellingStyle output…")
        thread = threading.Thread(target=self._run_spellingstyle, args=(doc, source_text, profile), daemon=True)
        thread.start()

    def _run_direct_input(self, add_trailing_space: bool) -> None:
        if self._busy:
            return
        source_text = self._read_editor_source_text()
        if source_text is None:
            return
        source_text = self._apply_word_substitutions(source_text)
        self._insert_text_into_writer(
            source_text,
            status_text="Inserting source text…",
            completed_text="Source text inserted.",
            add_trailing_space=add_trailing_space,
        )

    def _read_editor_source_text(self) -> str | None:
        source_path = self._editor_source_file
        if not source_path or not source_path.exists():
            self._show_toast("Choose a source text file first.")
            return None
        try:
            return source_path.read_text(encoding="utf-8")
        except OSError as exc:
            self._show_toast(f"Unable to read source file: {exc}")
            return None

    def _insert_text_into_writer(
        self,
        source_text: str,
        *,
        status_text: str,
        completed_text: str,
        add_trailing_space: bool,
        apply_case_citation_italics: bool = False,
    ) -> None:
        if not source_text.strip():
            self._show_toast("Source file is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        if source_text.startswith(" ") and self._get_preceding_char(doc) == " ":
            source_text = source_text.lstrip(" ")
        self._set_busy(True)
        self._status_label.set_label(status_text)
        self._append_editor_text(source_text)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            if apply_case_citation_italics:
                self._apply_case_citation_italics_to_recent_editor_insert(source_text)
            if add_trailing_space:
                self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            else:
                self._trim_trailing_whitespace(self._editor_insert_doc, self._editor_insert_cursor)
        self._capture_spellingstyle_range_end()
        self._set_busy(False)
        self._status_label.set_label(completed_text)

    def _preprocess_citation_input(self, source_text: str) -> str:
        normalized = source_text.lower()
        for word, digit in CITATION_NUMBER_WORDS.items():
            normalized = re.sub(rf"\b{word}\b", digit, normalized)
        normalized = (
            normalized.replace("\u00a0", " ")
            .replace("\u202f", " ")
            .replace("\u200a", " ")
            .replace("\u200b", "")
            .replace("\u3000", " ")
        )
        normalized = re.sub(r"(\d)\s*,\s*(\d)", r"\1\2", normalized)
        while re.search(r"(\d)\s*,\s*(\d)", normalized):
            normalized = re.sub(r"(\d)\s*,\s*(\d)", r"\1\2", normalized)
        normalized = re.sub(r"(\d)\s*:\s*(\d)", r"\1\2", normalized)
        normalized = re.sub(r"(\d)\s*\.\s*(\d)", r"\1\2", normalized)
        normalized = re.sub(r"\$\s*(\d)", r"\1", normalized)
        normalized = re.sub(r"\b(\d{1,5}) to (\d{1,5})\b", r"\1-\2", normalized)
        normalized = normalized.replace(" – ", "-").replace(" - ", "-").replace(" -- ", "-")
        normalized = re.sub(r"\sdash\s", "-", normalized)
        normalized = re.sub(r"^\s*-\s*", "", normalized)
        normalized = re.sub(r"(\d{1,5})\s*[–-]\s*(\d{1,5})", r"\1-\2", normalized)
        normalized = re.sub(r"(\d{1,5})-(\d{1,5})", r"\1–\2", normalized)
        normalized = re.sub(r"(\d)\s+(\d)", r"\1\2", normalized)
        while re.search(r"(\d)\s+(\d)", normalized):
            normalized = re.sub(r"(\d)\s+(\d)", r"\1\2", normalized)
        return normalized

    def _normalize_citation_input(self, source_text: str) -> tuple[int, int | None] | None:
        normalized = self._preprocess_citation_input(source_text)
        match = re.search(r"(\d{1,5})(?:\s*[–-]\s*(\d{1,5}))?", normalized)
        if not match:
            return None
        start = int(match.group(1))
        end = int(match.group(2)) if match.group(2) else None
        if end is not None and end < start:
            start, end = end, start
        return start, end

    def _format_citation_pages_source(self, pages: tuple[int, int | None]) -> str:
        start, end = pages
        if end is None:
            return str(start)
        return f"{start}-{end}"

    def _build_prefixed_citation(self, label: str, prefix: str, source_text: str) -> str | None:
        pages = self._normalize_citation_input(source_text)
        if not pages:
            return None
        start, end = pages
        prefix_text = "".join(ch for ch in prefix.strip() if ch.isdigit())
        cite_label = f"{prefix_text}{label}"
        if end is None:
            return f"({cite_label} {start}.)"
        return f"({cite_label} {start}\u2013{end}.)"

    def _run_citation_input(self, label: str, prefix: str) -> None:
        if self._busy:
            return
        source_text = self._read_editor_source_text()
        if source_text is None:
            return
        self._set_spelling_output_text(source_text)
        if not source_text.strip():
            self._show_toast("Source file is empty.")
            return
        self._set_busy(True)
        self._status_label.set_label(f"Validating {label} citation number…")
        thread = threading.Thread(
            target=self._run_validated_citation_input,
            args=(label, prefix, source_text),
            daemon=True,
        )
        thread.start()

    def _run_validated_citation_input(self, label: str, prefix: str, source_text: str) -> None:
        validator_error = ""
        validated_source = ""
        try:
            validated_source = self._validate_citation_number_with_llm(source_text)
        except Exception as exc:  # noqa: BLE001
            validator_error = str(exc)

        citation_source = validated_source if self._normalize_citation_input(validated_source) else source_text
        citation_pages = self._normalize_citation_input(citation_source)
        citation = self._build_prefixed_citation(label, prefix, citation_source) if citation_pages else None
        display_text = self._format_citation_pages_source(citation_pages) if citation_pages else source_text
        GLib.idle_add(
            self._finish_validated_citation_input,
            label,
            citation,
            display_text,
            validator_error,
        )

    def _validate_citation_number_with_llm(self, source_text: str) -> str:
        settings = self._citation_validator_settings
        if not settings.is_configured():
            raise ValueError("Configure Citation Number Validator API URL, API key, and prompt in Settings.")
        if settings.api_url.rstrip("/").endswith("/responses"):
            raise ValueError("Citation Number Validator uses a chat endpoint. Update the API URL in Settings.")
        payload = self._compose_citation_number_validator_payload(source_text)
        raw = self._post_json_and_read(
            payload,
            settings.api_url,
            settings.api_key,
            request_title="Citation Number Validator",
        )
        text = self._normalize_generated_output_text("".join(self._extract_response_text(raw)))
        cleaned = self._clean_citation_number_validator_output(text)
        if not cleaned:
            return ""
        pages = self._normalize_citation_input(cleaned)
        return self._format_citation_pages_source(pages) if pages else ""

    def _compose_citation_number_validator_payload(self, source_text: str) -> dict[str, Any]:
        settings = self._citation_validator_settings
        system_prompt = _expand_shared_prompt_parts(settings.prompt or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT)
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": source_text},
            ],
            "stream": False,
        }
        return self._add_model_id(
            payload,
            settings.model_id,
            disable_reasoning=settings.disable_reasoning,
        )

    def _clean_citation_number_validator_output(self, text: str) -> str:
        cleaned = self._normalize_generated_output_text(text)
        if cleaned.startswith("```"):
            fence_match = re.match(r"^```(?:text)?\s*([\s\S]*?)\s*```$", cleaned, flags=re.IGNORECASE)
            if fence_match:
                cleaned = fence_match.group(1).strip()
        cleaned = cleaned.strip().strip("\"'`")
        if cleaned.casefold() in {"none", "no number", "n/a", "na", "null"}:
            return ""
        return cleaned

    def _finish_validated_citation_input(
        self,
        label: str,
        citation: str | None,
        display_text: str,
        validator_error: str,
    ) -> bool:
        self._set_spelling_output_text(display_text)
        if not citation:
            self._set_busy(False)
            if validator_error:
                self._status_label.set_label(f"Citation validator failed: {validator_error}")
            else:
                self._status_label.set_label(f"No {label} page number found in source file.")
            self._show_toast(f"No {label} page number found in source file.")
            return False
        self._set_busy(False)
        self._insert_text_into_writer(
            citation,
            status_text=f"Inserting {label} citation…",
            completed_text=f"{label} citation inserted.",
            add_trailing_space=True,
        )
        return False

    def _on_direct_input_clicked(self, _button: Gtk.Button) -> None:
        self._run_direct_input(add_trailing_space=True)

    def _on_direct_input_no_trailing_space_clicked(self, _button: Gtk.Button) -> None:
        self._run_direct_input(add_trailing_space=False)

    def _on_speech_find_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        source_text = self._read_editor_source_text()
        if source_text is None:
            return
        query = self._normalize_speech_find_query(source_text)
        if not query:
            self._show_toast("Source file has no searchable speech text.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        self._ensure_writer_frame_active(desktop, doc)
        found = self._find_speech_query_range(doc, query)
        if not found:
            self._status_label.set_label(f'No match found for "{query}".')
            self._show_toast("No matching text found in Writer.")
            return
        try:
            doc.getCurrentController().select(found)
        except Exception:
            self._show_toast("Found text, but unable to select it in Writer.")
            return
        self._status_label.set_label(f'Found "{query}".')

    def _normalize_speech_find_query(self, source_text: str) -> str:
        normalized = (
            source_text.replace("\u00a0", " ")
            .replace("\u202f", " ")
            .replace("\u200a", " ")
            .replace("\u200b", "")
            .replace("\u3000", " ")
        )
        words = re.findall(r"[^\W_]+", normalized.casefold(), flags=re.UNICODE)
        return " ".join(words).strip()

    def _on_paste_clean_italics_action(self, source_path_raw: str) -> None:
        if self._busy:
            return
        source_path = Path(source_path_raw).expanduser()
        try:
            source_text = source_path.read_text(encoding="utf-8")
        except OSError as exc:
            self._show_toast(f"Unable to read paste text: {exc}")
            return
        self._insert_text_into_writer(
            source_text,
            status_text="Pasting cleaned text...",
            completed_text="Cleaned text pasted.",
            add_trailing_space=True,
            apply_case_citation_italics=True,
        )

    def _on_text_draft_spellingstyle_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("spelling", None)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        if profile.api_url.rstrip("/").endswith("/responses"):
            self._show_toast("SpellingStyle uses a chat endpoint. Update the API URL in Settings.")
            return
        source_text = self._read_editor_source_text()
        if source_text is None:
            return
        source_text = self._apply_word_substitutions(source_text)
        if not source_text.strip():
            self._show_toast("Source file is empty.")
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        if not route_to_terminal and not self._prepare_text_draft_append_output():
            self._show_toast("Unable to prepare Text Draft output.")
            return
        if route_to_terminal:
            self._text_draft_terminal_generated_output = ""
            self._text_draft_terminal_replace_text = ""
        else:
            GLib.idle_add(self._set_text_draft_original_output_text, "")
        self._clear_pending_text_draft_regenerate_context()
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label("Streaming SpellingStyle output into embedded Codex…")
        else:
            self._status_label.set_label("Streaming SpellingStyle output into Text Draft…")
        thread = threading.Thread(
            target=self._run_text_draft_spellingstyle,
            args=(source_text, profile, route_to_terminal, self._text_draft_terminal_pid),
            daemon=True,
        )
        thread.start()

    def _on_text_draft_improve_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("improve-generated", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        source_text = self._get_text_draft_generated_source_text().strip()
        if not source_text:
            self._show_toast("Original Output is empty.")
            return
        if not route_to_terminal and not self._prepare_text_draft_generated_replace():
            self._show_toast("Unable to replace the last generated draft output.")
            return
        if route_to_terminal and not self._delete_text_draft_terminal_generated_output():
            return
        self._set_pending_text_draft_regenerate_context("improve-generated", source_text)
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label(f"Streaming improved output into embedded Codex with {profile.display_name()}…")
        else:
            self._status_label.set_label(f"Improving Text Draft output with {profile.display_name()}…")
        thread = threading.Thread(
            target=self._run_text_draft_improve,
            args=(source_text, profile, route_to_terminal, self._text_draft_terminal_pid),
            daemon=True,
        )
        thread.start()

    def _on_text_draft_generated_choices_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        source_text = self._get_text_draft_generated_source_text().strip()
        if not source_text:
            self._show_toast("Original Output is empty.")
            return
        if (
            not route_to_terminal
            and (
                self._text_draft_buffer is None
                or self._text_draft_insert_start_mark is None
                or self._text_draft_insert_end_mark is None
            )
        ):
            self._show_toast("Unable to replace the last generated draft output.")
            return
        requests = self._build_editor_improve_rephrase_choice_requests("Text Draft Generated Choice")
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return
        self._start_multi_draft_choices_for_text(
            title="Text Draft Generated Choices",
            source_text=source_text,
            payload_builder=self._compose_improve_payload,
            prompt_text=self._improve1_settings.prompt,
            default_prompt=DEFAULT_IMPROVE_PROMPT,
            request_title="Text Draft Generated Choice",
            requests=requests,
        )
        self._multi_draft_insert_mode = "text-draft-terminal-replace-generated" if route_to_terminal else "text-draft-replace-generated"
        self._multi_draft_replace_doc = None
        self._multi_draft_replace_start = None
        self._multi_draft_replace_end = None

    def _on_text_draft_rephrase_generated_clicked(
        self,
        _button: Gtk.Button | None,
        profile_nickname: str | None = None,
    ) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("rephrase-generated", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        source_text = self._get_text_draft_generated_source_text().strip()
        if not source_text:
            self._show_toast("Original Output is empty.")
            return
        if not route_to_terminal and not self._prepare_text_draft_generated_replace():
            self._show_toast("Unable to replace the last generated draft output.")
            return
        if route_to_terminal and not self._delete_text_draft_terminal_generated_output():
            return
        self._set_pending_text_draft_regenerate_context("rephrase-generated", source_text)
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label(f"Streaming rephrased output into embedded Codex with {profile.display_name()}…")
        else:
            self._status_label.set_label(f"Rephrasing Text Draft output with {profile.display_name()}…")
        thread = threading.Thread(
            target=self._run_text_draft_rephrase_generated,
            args=(source_text, profile, route_to_terminal, self._text_draft_terminal_pid),
            daemon=True,
        )
        thread.start()

    def _on_text_draft_improve_selected_clicked(
        self,
        _button: Gtk.Button | None,
        profile_nickname: str | None = None,
    ) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("improve-selected", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        source_text = self._get_text_draft_selected_text().strip()
        if not source_text:
            self._show_toast("Select text in the Draft box first.")
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        if not route_to_terminal and not self._prepare_text_draft_selection_replace():
            self._show_toast("Unable to prepare Draft selection replacement.")
            return
        if not route_to_terminal:
            self._set_text_draft_original_output_text(source_text)
        self._set_pending_text_draft_regenerate_context("improve-selected", source_text)
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label(f"Streaming improved selection into embedded Codex with {profile.display_name()}…")
        else:
            self._status_label.set_label(f"Improving Draft selection with {profile.display_name()}…")
        thread = threading.Thread(
            target=self._run_text_draft_improve_selected,
            args=(source_text, profile, route_to_terminal, self._text_draft_terminal_pid),
            daemon=True,
        )
        thread.start()

    def _on_text_draft_selected_choices_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        source_text = self._get_text_draft_selected_text().strip()
        if not source_text:
            self._show_toast("Select text in the Draft box first.")
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        offsets = None if route_to_terminal else self._get_text_draft_selection_offsets()
        if not route_to_terminal and offsets is None:
            self._show_toast("Unable to remember the Draft selection.")
            return
        if not route_to_terminal:
            self._set_text_draft_original_output_text(source_text)
        requests = self._build_editor_improve_rephrase_choice_requests("Text Draft Selected Choice")
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return
        self._start_multi_draft_choices_for_text(
            title="Text Draft Selected Choices",
            source_text=source_text,
            payload_builder=self._compose_improve_payload,
            prompt_text=self._improve1_settings.prompt,
            default_prompt=DEFAULT_IMPROVE_PROMPT,
            request_title="Text Draft Selected Choice",
            requests=requests,
        )
        self._multi_draft_insert_mode = "text-draft-terminal" if route_to_terminal else "text-draft-replace-selection"
        self._multi_draft_replace_doc = None
        self._multi_draft_replace_start = offsets[0] if offsets is not None else None
        self._multi_draft_replace_end = offsets[1] if offsets is not None else None

    def _on_text_draft_keep_original_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        source_text = self._get_text_draft_generated_source_text().strip()
        if not source_text:
            self._show_toast("Original Output is empty.")
            return
        if not route_to_terminal and not self._prepare_text_draft_generated_replace():
            self._show_toast("Unable to replace the last generated draft output.")
            return
        if route_to_terminal and not self._delete_text_draft_terminal_generated_output():
            return
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label("Streaming original output into embedded Codex…")
            if not self._feed_text_draft_terminal_text(source_text):
                return
            self._set_text_draft_terminal_generated_output(source_text)
        else:
            self._status_label.set_label("Restoring Original Output in Draft…")
            self._append_text_draft_inserted_text(source_text)
            self._ensure_single_text_draft_trailing_space()
        self._set_busy(False)
        if route_to_terminal:
            self._status_label.set_label("Original output streamed to embedded Codex.")
        else:
            self._status_label.set_label("Original output restored in Draft.")

    def _on_text_draft_wrap_quotes_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        source_text = self._get_text_draft_selected_text()
        if not source_text.strip():
            self._show_toast("Select text in the Draft box first.")
            return
        wrapped = self._wrap_text_in_curly_quotes(source_text)
        route_to_terminal = self._text_draft_terminal_target_available()
        if not route_to_terminal and not self._prepare_text_draft_selection_replace():
            self._show_toast("Unable to prepare Draft selection replacement.")
            return
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label("Streaming quoted selection into embedded Codex…")
            if not self._feed_text_draft_terminal_text(wrapped):
                return
            self._set_text_draft_terminal_generated_output(wrapped)
        else:
            self._status_label.set_label("Wrapping Draft selection in quotes…")
            self._append_text_draft_inserted_text(wrapped)
        self._set_busy(False)
        if route_to_terminal:
            self._status_label.set_label("Quoted selection streamed to embedded Codex.")
        else:
            self._status_label.set_label("Draft selection wrapped in quotes.")

    def _on_copy_text_draft_clicked(self, _button: Gtk.Button | None) -> None:
        text = self._get_text_draft_submission_text()
        if not text.strip():
            self._show_toast("Draft is empty.")
            return
        clipboard = self.get_clipboard()
        clipboard.set(text)
        self._status_label.set_label("Draft copied to clipboard.")
        self._show_toast("Draft copied to clipboard.")

    def _on_clear_text_draft_clicked(self, _button: Gtk.Button) -> None:
        if not self._text_draft_has_user_text():
            self._refresh_text_draft_clear_button()
            return
        self._clear_text_draft_text()
        self._status_label.set_label("Draft cleared.")
        self._show_toast("Draft cleared.")

    def _expand_text_draft_external_action_value(self, value: str, draft_path: Path) -> str:
        expanded = value.replace(TEXT_DRAFT_EXTERNAL_ACTION_DRAFT_FILE_TOKEN, str(draft_path))
        return os.path.expandvars(os.path.expanduser(expanded))

    @staticmethod
    def _is_text_draft_case_log_action(action: TextDraftExternalAction) -> bool:
        if action.label.strip().lower() in {"log", "log entry", "case log"}:
            return True
        return any(
            script_name in part
            for part in action.command
            for script_name in (
                "prose-text-draft-case-log",
                "case_log_actions.py",
            )
        )

    @staticmethod
    def _text_draft_external_action_display_label(action: TextDraftExternalAction) -> str:
        if ProseWindow._is_text_draft_case_log_action(action):
            return "Log Entry"
        label = action.label.strip()
        return label or "External"

    @staticmethod
    def _text_draft_external_action_display_icon(action: TextDraftExternalAction) -> str:
        if ProseWindow._is_text_draft_case_log_action(action):
            return "document-edit-symbolic"
        if ProseWindow._is_text_draft_codex_action(action):
            return "send-to-symbolic"
        return action.icon_name.strip() or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON

    @staticmethod
    def _is_text_draft_codex_action(action: TextDraftExternalAction) -> bool:
        return _is_text_draft_codex_external_action(action)

    def _text_draft_codex_action(self) -> TextDraftExternalAction | None:
        return next(
            (
                action
                for action in self._text_draft_external_actions
                if self._is_text_draft_codex_action(action)
            ),
            None,
        )

    def _show_text_draft_terminal_controls(self) -> None:
        if self._text_draft_terminal_close_button is not None:
            self._text_draft_terminal_close_button.set_visible(True)
            self._text_draft_terminal_close_button.set_sensitive(True)

    def _hide_text_draft_terminal_controls(self) -> None:
        if self._text_draft_terminal_close_button is not None:
            self._text_draft_terminal_close_button.set_visible(False)

    def _text_draft_terminal_child_is_running(self, pid: int | None = None) -> bool:
        target_pid = self._text_draft_terminal_pid if pid is None else pid
        if target_pid is None:
            return self._text_draft_terminal_active
        try:
            os.kill(target_pid, 0)
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except OSError:
            return False
        return True

    def _text_draft_terminal_target_available(self) -> bool:
        if Vte is None or self._text_draft_terminal is None:
            return False
        if not self._text_draft_terminal_active:
            return False
        if self._text_draft_terminal_mode != "codex":
            return False
        if self._text_draft_terminal_child_is_running():
            return True
        self._text_draft_terminal_active = False
        self._text_draft_terminal_pid = None
        self._text_draft_terminal_mode = None
        return False

    def _feed_text_draft_terminal_text(self, text: str) -> bool:
        if not text:
            return True
        terminal = self._text_draft_terminal
        if terminal is None or not self._text_draft_terminal_target_available():
            self._on_text_draft_failed("Embedded Codex is no longer running.")
            return False
        try:
            terminal.feed_child(text.encode("utf-8"))
            terminal.grab_focus()
        except Exception as exc:  # noqa: BLE001
            self._on_text_draft_failed(f"Unable to stream text into embedded Codex: {exc}")
            return False
        return True

    def _append_text_draft_terminal_text(self, text: str) -> bool:
        self._feed_text_draft_terminal_text(text)
        return False

    def _set_text_draft_terminal_generated_output(self, text: str) -> bool:
        self._text_draft_terminal_generated_output = self._normalize_generated_output_text(text)
        self._text_draft_terminal_replace_text = text
        return False

    def _get_text_draft_generated_source_text(self) -> str:
        if self._text_draft_terminal_target_available() and self._text_draft_terminal_generated_output.strip():
            return self._text_draft_terminal_generated_output
        return self._get_text_draft_original_output_text()

    def _delete_text_draft_terminal_generated_output(self) -> bool:
        text = self._text_draft_terminal_replace_text or self._text_draft_terminal_generated_output
        if not text:
            return True
        backspaces = "\x7f" * len(text)
        if not self._feed_text_draft_terminal_text(backspaces):
            return False
        self._text_draft_terminal_generated_output = ""
        self._text_draft_terminal_replace_text = ""
        return True

    def _text_draft_embedded_terminal_unavailable(self) -> None:
        if self._text_draft_draft_surface is not None:
            self._text_draft_draft_surface.set_visible_child_name("terminal-missing")
            self._show_text_draft_terminal_controls()
        message = "Install gir1.2-vte-3.91 and libvte-2.91-gtk4-0 to use the embedded terminal."
        self._status_label.set_label(message)
        self._show_toast(message)

    def _reset_text_draft_terminal_state(self) -> None:
        self._text_draft_terminal_active = False
        self._text_draft_terminal_pid = None
        self._text_draft_terminal_mode = None
        self._text_draft_terminal_cwd = None
        self._text_draft_terminal_generated_output = ""
        self._text_draft_terminal_replace_text = ""

    def _spawn_text_draft_terminal(
        self,
        argv: list[str],
        cwd: str,
        env: dict[str, str],
        action_label: str,
        mode: str,
        clear_draft_on_return: bool = False,
    ) -> None:
        terminal = self._text_draft_terminal
        if Vte is None or terminal is None:
            self._text_draft_embedded_terminal_unavailable()
            return
        if self._text_draft_terminal_active:
            if self._text_draft_draft_surface is not None:
                self._text_draft_draft_surface.set_visible_child_name("terminal")
            terminal.grab_focus()
            self._show_toast("Embedded terminal is already running.")
            return

        try:
            terminal.reset(True, True)
            _apply_prose_terminal_theme(terminal)
            terminal.spawn_async(
                Vte.PtyFlags.DEFAULT,
                cwd,
                argv,
                [f"{key}={value}" for key, value in env.items()],
                GLib.SpawnFlags.DEFAULT,
                None,
                None,
                -1,
                None,
                self._on_text_draft_terminal_spawned,
                (action_label, mode),
            )
        except Exception as exc:  # noqa: BLE001
            self._status_label.set_label(f"Unable to start embedded {action_label}: {exc}")
            self._show_toast(f"Unable to start embedded {action_label}.")
            return

        self._text_draft_terminal_active = True
        self._text_draft_terminal_mode = mode
        self._text_draft_terminal_cwd = cwd
        self._text_draft_terminal_closing = False
        self._text_draft_terminal_generated_output = ""
        self._text_draft_terminal_replace_text = ""
        self._text_draft_clear_draft_on_terminal_return = clear_draft_on_return
        if self._text_draft_draft_surface is not None:
            self._text_draft_draft_surface.set_visible_child_name("terminal")
        self._show_text_draft_terminal_controls()
        self._status_label.set_label(f"Started embedded {action_label}.")
        terminal.grab_focus()

    def _start_text_draft_codex_terminal(
        self,
        action: TextDraftExternalAction,
        draft_path: Path,
        cwd_override: Path | None = None,
        reasoning_effort: str = "medium",
    ) -> None:
        terminal = self._text_draft_terminal
        if Vte is None or terminal is None:
            self._text_draft_embedded_terminal_unavailable()
            return
        if self._text_draft_terminal_active:
            if self._text_draft_draft_surface is not None:
                self._text_draft_draft_surface.set_visible_child_name("terminal")
            terminal.grab_focus()
            self._show_toast("Embedded Codex is already running.")
            return

        env = os.environ.copy()
        env["PROSE_TEXT_DRAFT_FILE"] = str(draft_path)
        env.update(
            {
                key: self._expand_text_draft_external_action_value(value, draft_path)
                for key, value in action.env.items()
            }
        )
        if "RR_CLI_KEEP_BROWSER_OPEN" not in action.env:
            env["RR_CLI_KEEP_BROWSER_OPEN"] = "1"
        env["CODEX_REASONING_EFFORT"] = reasoning_effort
        env.pop("PROSE_CODEX_EMPTY_SESSION", None)
        cwd_path = cwd_override or action.cwd
        cwd = str(cwd_path) if cwd_path is not None else os.getcwd()
        if cwd_override is not None:
            env["CODEX_CWD"] = cwd
        argv = [
            self._expand_text_draft_external_action_value(part, draft_path)
            for part in action.command
        ]
        if (
            not argv
            or any("prose-text-draft-codex-ghostty.sh" in part for part in argv)
        ):
            if not TEXT_DRAFT_CODEX_VTE_SCRIPT.is_file():
                message = f"Codex VTE script not found: {TEXT_DRAFT_CODEX_VTE_SCRIPT}"
                self._status_label.set_label(message)
                self._show_toast(message)
                return
            argv = ["bash", str(TEXT_DRAFT_CODEX_VTE_SCRIPT), str(draft_path)]

        self._spawn_text_draft_terminal(
            argv,
            cwd,
            env,
            action.label,
            "codex",
            clear_draft_on_return=True,
        )

    def _start_text_draft_shell_terminal(self, fallback: bool = False, cwd: str | None = None) -> None:
        terminal = self._text_draft_terminal
        if Vte is None or terminal is None:
            self._text_draft_embedded_terminal_unavailable()
            return
        shell = os.environ.get("SHELL") or "/bin/bash"
        if not Path(shell).is_file():
            shell = "/bin/bash"
        if not Path(shell).is_file():
            message = "Unable to find a shell for the embedded terminal."
            self._status_label.set_label(message)
            self._show_toast(message)
            return
        if self._text_draft_terminal_active:
            if self._text_draft_draft_surface is not None:
                self._text_draft_draft_surface.set_visible_child_name("terminal")
            terminal.grab_focus()
            self._show_toast("Embedded terminal is already running.")
            return

        label = "terminal" if not fallback else "terminal after Codex exited"
        shell_cwd = cwd or str(TEXT_DRAFT_PROJECTS_DIR)
        self._spawn_text_draft_terminal(
            [shell],
            shell_cwd,
            os.environ.copy(),
            label,
            "shell",
            clear_draft_on_return=(
                self._text_draft_clear_draft_on_terminal_return if fallback else False
            ),
        )

    def _on_text_draft_terminal_spawned(
        self,
        terminal: Any,
        pid: int,
        error: GLib.Error | None,
        user_data: tuple[str, str],
    ) -> None:
        action_label, mode = user_data
        if error is not None:
            self._reset_text_draft_terminal_state()
            self._show_text_draft_terminal_controls()
            self._status_label.set_label(f"Unable to start embedded {action_label}: {error.message}")
            self._show_toast(f"Unable to start embedded {action_label}.")
            return
        self._text_draft_terminal_pid = int(pid)
        self._text_draft_terminal_mode = mode

    def _on_text_draft_terminal_child_exited(self, _terminal: Any, _status: int) -> None:
        ended_mode = self._text_draft_terminal_mode
        ended_cwd = self._text_draft_terminal_cwd
        closing = self._text_draft_terminal_closing
        self._reset_text_draft_terminal_state()
        self._text_draft_terminal_closing = False
        if closing:
            self._hide_text_draft_terminal_controls()
            self._status_label.set_label("Embedded terminal closed.")
            return
        self._show_text_draft_terminal_controls()
        if ended_mode == "codex" and not closing:
            self._status_label.set_label("Embedded Codex session ended; opening terminal.")
            self._start_text_draft_shell_terminal(fallback=True, cwd=ended_cwd)
            return
        self._status_label.set_label("Embedded terminal session ended.")

    def _on_text_draft_terminal_style_changed(self, *_args: object) -> None:
        terminal = self._text_draft_terminal
        if terminal is not None:
            _apply_prose_terminal_theme(terminal)

    def _on_text_draft_terminal_key_pressed(
        self,
        _controller: Gtk.EventControllerKey,
        keyval: int,
        _keycode: int,
        state: Gdk.ModifierType,
    ) -> bool:
        terminal = self._text_draft_terminal
        if Vte is None or terminal is None:
            return False
        required_modifiers = Gdk.ModifierType.CONTROL_MASK | Gdk.ModifierType.SHIFT_MASK
        if state & required_modifiers != required_modifiers:
            return False
        if keyval in (Gdk.KEY_C, Gdk.KEY_c):
            terminal.copy_clipboard_format(Vte.Format.TEXT)
            return True
        if keyval in (Gdk.KEY_V, Gdk.KEY_v):
            terminal.paste_clipboard()
            return True
        return False

    def _on_text_draft_terminal_close_clicked(self, _button: Gtk.Button) -> None:
        clear_draft_on_return = self._text_draft_clear_draft_on_terminal_return
        if self._text_draft_terminal_active and self._text_draft_terminal_pid is not None:
            self._text_draft_terminal_closing = True
            try:
                os.kill(self._text_draft_terminal_pid, signal.SIGTERM)
            except OSError:
                pass
        self._reset_text_draft_terminal_state()
        self._text_draft_clear_draft_on_terminal_return = False
        if self._text_draft_draft_surface is not None:
            self._text_draft_draft_surface.set_visible_child_name("draft")
        if clear_draft_on_return:
            self._clear_text_draft_text()
        self._hide_text_draft_terminal_controls()
        if self._text_draft_view is not None:
            self._text_draft_view.grab_focus()

    def _on_text_draft_external_action_clicked(
        self,
        _button: Gtk.Button | None,
        action: TextDraftExternalAction | None = None,
    ) -> None:
        if action is None:
            action = next((item for item in self._text_draft_external_actions if item.is_configured()), None)
        if action is None or not action.is_configured():
            self._show_toast("No Text Draft external action is configured.")
            return
        if self._busy:
            return
        text = self._get_text_draft_submission_text()
        if not text.strip():
            self._show_toast("Draft is empty.")
            return
        draft_path = self._write_text_draft_temp_file_now()
        if draft_path is None:
            return

        if self._is_text_draft_codex_action(action):
            self._start_text_draft_codex_terminal(
                action,
                draft_path,
                reasoning_effort=_sanitize_text_draft_codex_reasoning_effort(
                    action.codex_reasoning_effort
                ),
            )
            return

        command = [
            self._expand_text_draft_external_action_value(part, draft_path)
            for part in action.command
        ]
        env = os.environ.copy()
        env["PROSE_TEXT_DRAFT_FILE"] = str(draft_path)
        env.update(
            {
                key: self._expand_text_draft_external_action_value(value, draft_path)
                for key, value in action.env.items()
            }
        )
        cwd = str(action.cwd) if action.cwd is not None else None
        try:
            subprocess.Popen(command, cwd=cwd, env=env, start_new_session=True)
        except (OSError, ValueError) as exc:
            self._status_label.set_label(f"Unable to launch {action.label}: {exc}")
            self._show_toast(f"Unable to launch {action.label}.")
            return
        display_label = self._text_draft_external_action_display_label(action)
        success_message = action.success_message.strip() or f"Launched {display_label}."
        self._status_label.set_label(success_message)
        self._show_toast(success_message)

    def _prepare_text_draft_regenerate_output_state(self, context: RegenerateContext) -> None:
        if context.action_key in REGENERATE_SOURCE_BUFFER_ACTION_KEYS:
            self._set_text_draft_original_output_text(context.source_text)

    def _run_text_draft_regenerate_command(
        self,
        context: RegenerateContext,
        profile: ModelProfile,
        route_to_terminal: bool = False,
        terminal_pid: int | None = None,
    ) -> None:
        if context.action_key == "improve-generated":
            self._run_text_draft_improve(context.source_text, profile, route_to_terminal, terminal_pid)
            return
        if context.action_key == "rephrase-generated":
            self._run_text_draft_rephrase_generated(context.source_text, profile, route_to_terminal, terminal_pid)
            return
        if context.action_key == "improve-selected":
            self._run_text_draft_improve_selected(context.source_text, profile, route_to_terminal, terminal_pid)
            return
        GLib.idle_add(self._on_text_draft_failed, f'Unsupported Text Draft regenerate action "{context.action_key}".')

    def _on_text_draft_regenerate_clicked(
        self,
        _button: Gtk.Button | None,
        profile_nickname: str | None = None,
    ) -> None:
        if self._busy:
            return
        context = self._text_draft_last_regenerate_context
        if context is None:
            self._show_toast("Run a supported Text Draft command first.")
            return
        profile = self._resolve_profile_for_action(context.action_key, profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        route_to_terminal = self._text_draft_terminal_target_available()
        if not route_to_terminal and not self._prepare_text_draft_generated_replace():
            self._show_toast("Unable to replace the last generated draft output.")
            return
        if route_to_terminal and not self._delete_text_draft_terminal_generated_output():
            return
        self._set_pending_text_draft_regenerate_context(context.action_key, context.source_text)
        if not route_to_terminal:
            self._prepare_text_draft_regenerate_output_state(context)
        self._set_busy(True)
        if route_to_terminal:
            self._status_label.set_label(
                f"Streaming regenerated {context.command_title.lower()} into embedded Codex with {profile.display_name()}…"
            )
        else:
            self._status_label.set_label(
                f"Regenerating {context.command_title.lower()} in Text Draft with {profile.display_name()}…"
            )
        thread = threading.Thread(
            target=self._run_text_draft_regenerate_command,
            args=(context, profile, route_to_terminal, self._text_draft_terminal_pid),
            daemon=True,
        )
        thread.start()

    def _on_input_rt_clicked(self, _button: Gtk.Button) -> None:
        self._run_citation_input("RT", self._prefix_settings.rt_prefix)

    def _on_input_ct_clicked(self, _button: Gtk.Button) -> None:
        self._run_citation_input("CT", self._prefix_settings.ct_prefix)

    def _apply_word_substitutions(self, source_text: str) -> str:
        updated = source_text
        for entry in self._prefix_settings.substitutions:
            if not entry.original or not entry.replacement:
                continue
            updated = updated.replace(entry.original, entry.replacement)
        return updated

    def _collect_prefix_settings_from_ui(self) -> PrefixSettings:
        substitutions: list[WordSubstitution] = []
        for original_entry, replacement_entry in self._substitution_rows[:MAX_WORD_SUBSTITUTIONS]:
            substitutions.append(
                WordSubstitution(
                    original=original_entry.get_text().strip(),
                    replacement=replacement_entry.get_text().strip(),
                )
            )
        rt_prefix = self._rt_prefix_entry.get_text().strip() if self._rt_prefix_entry else self._prefix_settings.rt_prefix
        ct_prefix = self._ct_prefix_entry.get_text().strip() if self._ct_prefix_entry else self._prefix_settings.ct_prefix
        return PrefixSettings(rt_prefix=rt_prefix, ct_prefix=ct_prefix, substitutions=substitutions)

    def _on_save_prefixes_clicked(self, _button: Gtk.Button) -> None:
        self._prefix_settings = self._collect_prefix_settings_from_ui()
        save_prefix_settings(self._prefix_settings)
        self._show_toast("Prefixes and substitutions saved.")

    def _on_combine_cites_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("combine", None)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        self._ensure_writer_frame_active(desktop, doc)
        found = self._find_first_uncombined_cites(doc)
        if not found:
            self._show_toast("No uncombined citation runs found.")
            return
        cite_text, cite_cursor = found
        self._combine_cites_doc = doc
        self._combine_cites_cursor = cite_cursor
        self._set_busy(True)
        self._status_label.set_label(f"Combining citations with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_combine_cites, args=(cite_text, profile), daemon=True)
        thread.start()

    def _on_improve_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("improve-generated", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        source_text = self._get_spelling_output_text().strip()
        if not source_text:
            self._show_toast("SpellingStyle output is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._select_spellingstyle_range(doc):
            self._show_toast("Unable to select the last SpellingStyle range.")
            return
        if not self._prepare_improve_insertion(doc):
            self._show_toast("Unable to prepare Improve Generated insertion point.")
            return
        self._set_pending_regenerate_context("improve-generated", source_text)
        self._set_spelling_output_text(source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Improving generated text with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_improve, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_rephrase_generated_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("rephrase-generated", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        source_text = self._get_spelling_output_text().strip()
        if not source_text:
            self._show_toast("SpellingStyle output is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._select_spellingstyle_range(doc):
            self._show_toast("Unable to select the last SpellingStyle range.")
            return
        if not self._prepare_improve_insertion(doc):
            self._show_toast("Unable to prepare Rephrase Generated insertion point.")
            return
        self._set_pending_regenerate_context("rephrase-generated", source_text)
        self._set_spelling_output_text(source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Rephrasing generated text with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_rephrase_generated, args=(source_text, profile), daemon=True)
        thread.start()

    def _prepare_editor_regenerate_insertion(self, doc: XTextDocument) -> bool:  # type: ignore[type-arg]
        if not self._editor_insert_end or self._last_insert_len <= 0:
            return False
        try:
            text = self._get_text_container(doc, self._editor_insert_end)
            if not text:
                return False
            range_cursor = text.createTextCursorByRange(self._editor_insert_end)
            if not range_cursor.goLeft(self._last_insert_len, True):
                return False
            try:
                range_cursor.setString("")
            except Exception:
                pass
            self._editor_insert_cursor = text.createTextCursorByRange(range_cursor)
            self._editor_insert_doc = doc
            self._editor_insert_end = None
            self._last_insert_len = 0
            self._editor_pending_newlines = 0
            return True
        except Exception:
            return False

    def _prepare_regenerate_output_state(self, context: RegenerateContext) -> None:
        if context.action_key in REGENERATE_SOURCE_BUFFER_ACTION_KEYS:
            self._set_spelling_output_text(context.source_text)
        else:
            self._set_spelling_output_text("")

    def _prepare_regenerate_insertion(self, doc: XTextDocument, context: RegenerateContext) -> bool:  # type: ignore[type-arg]
        if not self._select_spellingstyle_range(doc):
            return False
        if context.insert_mode == "editor":
            return self._prepare_editor_regenerate_insertion(doc)
        if context.insert_mode == "improve":
            return self._prepare_improve_insertion(doc)
        return False

    def _run_regenerate_command(self, context: RegenerateContext, profile: ModelProfile) -> None:
        action_key = context.action_key
        source_text = context.source_text
        if action_key == "improve-generated":
            self._run_improve(source_text, profile)
            return
        if action_key == "rephrase-generated":
            self._run_rephrase_generated(source_text, profile)
            return
        if action_key == "improve-selected":
            self._run_improve_selected(source_text, profile)
            return
        if action_key == "shorten":
            self._run_shorten(source_text, profile)
            return
        if action_key == "topic-sentence":
            self._run_topic_sentence(source_text, profile)
            return
        if action_key == "intro":
            self._run_introduction(source_text, profile)
            return
        if action_key == "intro-reply":
            self._run_introduction_reply(source_text, profile)
            return
        if action_key == "conclusion":
            self._run_conclusion(source_text, profile)
            return
        if action_key == "concl-no-issues":
            self._run_conclusion_no_issues(source_text, profile)
            return
        if action_key == "concl-section":
            self._run_concl_section(source_text, profile)
            return
        GLib.idle_add(self._on_spellingstyle_failed, f'Unsupported regenerate action "{action_key}".')

    def _on_regenerate_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        context = self._last_regenerate_context
        if context is None:
            self._show_toast("Run a supported text-generation command first.")
            return
        profile = self._resolve_profile_for_action(context.action_key, profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._prepare_regenerate_insertion(doc, context):
            self._show_toast("Unable to replace the last generated output.")
            return
        self._set_pending_regenerate_context(context.action_key, context.source_text)
        self._prepare_regenerate_output_state(context)
        self._set_busy(True)
        self._status_label.set_label(f"Regenerating {context.command_title.lower()} with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_regenerate_command, args=(context, profile), daemon=True)
        thread.start()

    def _on_improve_selected_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("improve-selected", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        self._set_spelling_output_text(source_text)
        if not self._prepare_selection_insertion(doc, view_cursor):
            self._show_toast("Unable to prepare selected text replacement.")
            return
        self._set_pending_regenerate_context("improve-selected", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Improving selected text with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_improve_selected, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_keep_original_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        source_text = self._get_spelling_output_text().strip()
        if not source_text:
            self._show_toast("SpellingStyle output is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._select_spellingstyle_range(doc):
            self._show_toast("Unable to select the last streamed range.")
            return
        if not self._prepare_improve_insertion(doc):
            self._show_toast("Unable to prepare insertion point.")
            return
        self._set_busy(True)
        self._status_label.set_label("Restoring SpellingStyle output…")
        self._append_improve1_text(source_text)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
        self._capture_improve1_range_end()
        self._set_busy(False)
        self._status_label.set_label("Original output restored.")

    def _on_thesaurus_button_clicked(self, _button: Gtk.Button) -> None:
        self._on_thesaurus_clicked(_button)

    def _on_thesaurus_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("thesaurus", None)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        self._show_thesaurus_output()
        self._set_thesaurus_placeholder("Looking up alternatives…")
        self._thesaurus_words = []
        self._render_thesaurus_words()
        self._set_busy(True)
        self._status_label.set_label(f"Fetching thesaurus alternatives with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_thesaurus, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_reference_button_clicked(self, _button: Gtk.Button) -> None:
        self._on_reference_clicked(_button)

    def _on_reference_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        if not self._reference_settings.is_configured():
            self._show_toast(
                "Add Reference API URL, API key, Tavily API key, and prompt in Settings. Tavily CLI (`tvly`) must also be installed."
            )
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        self._show_reference_output(mode="lookup")
        self._set_reference_placeholder("Looking up definition…")
        self._set_busy(True)
        self._status_label.set_label("Fetching reference definition…")
        thread = threading.Thread(target=self._run_reference, args=(source_text, None), daemon=True)
        thread.start()

    def _on_reference_question_activated(self, entry: Gtk.Entry) -> None:
        question = entry.get_text().strip()
        if not question:
            return
        if self._busy:
            return
        if not self._ask_settings.is_configured():
            self._show_toast(
                "Add Ask API URL, API key, Tavily API key, and prompt in Settings. Tavily CLI (`tvly`) must also be installed."
            )
            return
        self._show_reference_output(mode="ask")
        self._set_reference_placeholder("Answering question…")
        self._set_busy(True)
        self._status_label.set_label("Answering question…")
        thread = threading.Thread(target=self._run_ask, args=(question,), daemon=True)
        thread.start()

    def _on_focus_ask(self) -> None:
        if hasattr(self, "_reference_query_entry"):
            self._reference_query_entry.grab_focus()

    def _on_shorten_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("shorten", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        self._set_spelling_output_text(source_text)
        if not self._prepare_selection_insertion(doc, view_cursor):
            self._show_toast("Unable to prepare selected text replacement.")
            return
        self._set_pending_regenerate_context("shorten", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Shortening selection with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_shorten, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_wrap_quotes_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        wrapped = self._wrap_text_in_curly_quotes(source_text)
        self._replace_selected_text_no_trailing_space(wrapped)

    def _on_translate_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("translate", None)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        self._set_busy(True)
        self._status_label.set_label(f"Translating document to Spanish with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_translate_document, args=(doc, profile), daemon=True)
        thread.start()

    def _on_add_case_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        citation = " ".join(view_cursor.getString().split())
        if not citation:
            self._show_toast("Select a case citation in Writer first.")
            return
        self._set_busy(True)
        self._status_label.set_label("Adding case to concordance and AutoText…")
        selection_start = view_cursor.getStart()
        selection_end = view_cursor.getEnd()
        thread = threading.Thread(
            target=self._run_add_case,
            args=(citation, doc, selection_start, selection_end),
            daemon=True,
        )
        thread.start()

    def _run_add_case(
        self,
        citation: str,
        doc: XTextDocument,
        selection_start: Any,
        selection_end: Any,
    ) -> None:
        try:
            self._append_case_to_concordance(citation)
            added_autotext = self._add_case_to_autotext(citation, doc, selection_start, selection_end)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_add_case_failed, f"{type(exc).__name__}: {exc}")
            return
        GLib.idle_add(self._on_add_case_finished, added_autotext)

    def _on_add_case_finished(self, added_autotext: bool) -> bool:
        if added_autotext:
            self._show_toast("Added case to concordance and AutoText.")
        else:
            self._show_toast("Added case to concordance. AutoText already had it.")
        self._status_label.set_label("Ready.")
        self._set_busy(False)
        return GLib.SOURCE_REMOVE

    def _on_add_case_failed(self, message: str) -> bool:
        self._show_toast(f"Unable to add case: {message}")
        self._status_label.set_label("Ready.")
        self._set_busy(False)
        return GLib.SOURCE_REMOVE

    def _on_topic_sentence_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("topic-sentence", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_paragraph_at_cursor(doc)
        if not source_text:
            self._show_toast("Place the cursor in a paragraph to summarize.")
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("topic-sentence", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Creating topic sentence with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_topic_sentence, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_introduction_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("intro", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_section_between_headings(doc, "ARGUMENT", "CONCLUSION")
        if not source_text:
            self._show_toast('Unable to find text between "ARGUMENT" and "CONCLUSION".')
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("intro", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Writing introduction with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_introduction, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_introduction_reply_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("intro-reply", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_section_between_headings(doc, "ARGUMENT", "CONCLUSION")
        if not source_text:
            self._show_toast('Unable to find text between "ARGUMENT" and "CONCLUSION".')
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("intro-reply", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Writing introduction for reply with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_introduction_reply, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_conclusion_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("conclusion", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_section_between_headings(doc, "ARGUMENT", "CONCLUSION")
        if not source_text:
            self._show_toast('Unable to find text between "ARGUMENT" and "CONCLUSION".')
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("conclusion", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Writing conclusion with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_conclusion, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_concl_no_issues_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("concl-no-issues", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_section_between_headings(doc, "ISSUES CONSIDERED", "CONCLUSION")
        if not source_text:
            self._show_toast('Unable to find text between "ISSUES CONSIDERED" and "CONCLUSION".')
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("concl-no-issues", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Writing conclusion without issues with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_conclusion_no_issues, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_introduction_choices_clicked(self, _button: Gtk.Button | None) -> None:
        self._start_multi_draft_choices(
            title="Introduction Choices",
            source_start="ARGUMENT",
            source_end="CONCLUSION",
            payload_builder=self._compose_introduction_payload,
            prompt_text=self._introduction_settings.prompt,
            default_prompt=DEFAULT_INTRO_PROMPT,
            request_title="Introduction Choice",
        )

    def _on_introduction_reply_choices_clicked(self, _button: Gtk.Button | None) -> None:
        self._start_multi_draft_choices(
            title="Introduction Reply Choices",
            source_start="ARGUMENT",
            source_end="CONCLUSION",
            payload_builder=self._compose_introduction_reply_payload,
            prompt_text=self._introduction_reply_settings.prompt,
            default_prompt=DEFAULT_INTRO_REPLY_PROMPT,
            request_title="Introduction Reply Choice",
        )

    def _on_conclusion_choices_clicked(self, _button: Gtk.Button | None) -> None:
        self._start_multi_draft_choices(
            title="Conclusion Choices",
            source_start="ARGUMENT",
            source_end="CONCLUSION",
            payload_builder=self._compose_conclusion_payload,
            prompt_text=self._conclusion_settings.prompt,
            default_prompt=DEFAULT_CONCLUSION_PROMPT,
            request_title="Conclusion Choice",
        )

    def _on_concl_no_issues_choices_clicked(self, _button: Gtk.Button | None) -> None:
        self._start_multi_draft_choices(
            title="Conclusion No Issues Choices",
            source_start="ISSUES CONSIDERED",
            source_end="CONCLUSION",
            payload_builder=self._compose_conclusion_no_issues_payload,
            prompt_text=self._concl_no_issues_settings.prompt,
            default_prompt=DEFAULT_CONCL_NO_ISSUES_PROMPT,
            request_title="Conclusion No Issues Choice",
        )

    def _on_generated_choices_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        source_text = self._get_spelling_output_text().strip()
        if not source_text:
            self._show_toast("SpellingStyle output is empty.")
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if not self._select_spellingstyle_range(doc):
            self._show_toast("Unable to select the last SpellingStyle range.")
            return
        requests = self._build_editor_improve_rephrase_choice_requests("Generated Choice")
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return
        self._start_multi_draft_choices_for_text(
            title="Generated Choices",
            source_text=source_text,
            payload_builder=self._compose_improve_payload,
            prompt_text=self._improve1_settings.prompt,
            default_prompt=DEFAULT_IMPROVE_PROMPT,
            request_title="Generated Choice",
            requests=requests,
        )
        self._multi_draft_insert_mode = "replace-generated"
        self._multi_draft_replace_doc = None
        self._multi_draft_replace_start = None
        self._multi_draft_replace_end = None

    def _on_selected_choices_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = view_cursor.getString()
        if not source_text.strip():
            self._show_toast("Select text in Writer first.")
            return
        try:
            selection_start = view_cursor.getStart()
            selection_end = view_cursor.getEnd()
        except Exception:
            self._show_toast("Unable to remember the selected text range.")
            return
        requests = self._build_editor_improve_rephrase_choice_requests("Selected Choice")
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return
        self._start_multi_draft_choices_for_text(
            title="Selected Choices",
            source_text=source_text,
            payload_builder=self._compose_improve_payload,
            prompt_text=self._improve1_settings.prompt,
            default_prompt=DEFAULT_IMPROVE_PROMPT,
            request_title="Selected Choice",
            requests=requests,
        )
        self._multi_draft_insert_mode = "replace-selection"
        self._multi_draft_replace_doc = doc
        self._multi_draft_replace_start = selection_start
        self._multi_draft_replace_end = selection_end

    def _on_concl_section_clicked(self, _button: Gtk.Button | None, profile_nickname: str | None = None) -> None:
        if self._busy:
            return
        profile = self._resolve_profile_for_action("concl-section", profile_nickname)
        if profile is None:
            return
        if not profile.is_configured():
            self._show_toast(f'Configure the "{profile.display_name()}" model profile in Settings first.')
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = self._extract_section_to_last_heading(doc, view_cursor)
        if not source_text:
            self._show_toast("Unable to find a heading before the cursor.")
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._set_spelling_output_text("")
        self._set_pending_regenerate_context("concl-section", source_text)
        self._set_busy(True)
        self._status_label.set_label(f"Writing section conclusion with {profile.display_name()}…")
        thread = threading.Thread(target=self._run_concl_section, args=(source_text, profile), daemon=True)
        thread.start()

    def _on_concl_section_choices_clicked(self, _button: Gtk.Button | None) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        source_text = self._extract_section_to_last_heading(doc, view_cursor)
        if not source_text:
            self._show_toast("Unable to find a heading before the cursor.")
            return
        requests = self._build_editor_stacked_choice_requests(
            request_title="Section Conclusion Choice",
            payload_builder=self._compose_concl_section_payload,
            prompt_text=self._concl_section_settings.prompt,
            default_prompt=DEFAULT_CONCL_SECTION_PROMPT,
        )
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return
        self._multi_draft_insert_mode = "editor"
        self._multi_draft_replace_doc = None
        self._multi_draft_replace_start = None
        self._multi_draft_replace_end = None
        self._start_multi_draft_choices_for_text(
            title="Section Conclusion Choices",
            source_text=source_text,
            payload_builder=self._compose_concl_section_payload,
            prompt_text=self._concl_section_settings.prompt,
            default_prompt=DEFAULT_CONCL_SECTION_PROMPT,
            request_title="Section Conclusion Choice",
            requests=requests,
        )

    def _build_editor_stacked_choice_requests(
        self,
        *,
        request_title: str,
        payload_builder: Callable[[str, ModelProfile], dict[str, Any]],
        prompt_text: str,
        default_prompt: str,
    ) -> list[MultiDraftRequest]:
        requests: list[MultiDraftRequest] = []
        slot_specs = (
            ("choices-improve", "Choice 1"),
            ("choices-rephrase-1", "Choice 2"),
            ("choices-rephrase-2", "Choice 3"),
        )
        for slot_key, label in slot_specs:
            profile_key = self._editor_action_profile_defaults.get(slot_key)
            profile = self._model_profile_by_key(profile_key or "")
            if profile is None:
                continue
            requests.append(
                MultiDraftRequest(
                    key=slot_key,
                    label=profile.display_name(),
                    profile=profile,
                    payload_builder=payload_builder,
                    prompt_text=prompt_text,
                    default_prompt=default_prompt,
                    request_title=f"{request_title} {label}",
                    stacked=True,
                )
            )
        return requests

    def _build_editor_improve_rephrase_choice_requests(self, request_title: str) -> list[MultiDraftRequest]:
        requests: list[MultiDraftRequest] = []
        slot_specs = (
            (
                "choices-improve",
                "Improve",
                self._compose_improve_payload,
                self._improve1_settings.prompt,
                DEFAULT_IMPROVE_PROMPT,
                "improve",
            ),
            (
                "choices-rephrase-1",
                "Rephrase 1",
                self._compose_rephrase_generated_payload,
                self._improve2_settings.prompt,
                DEFAULT_IMPROVE2_PROMPT,
                "rephrase-1",
            ),
            (
                "choices-rephrase-2",
                "Rephrase 2",
                self._compose_rephrase_generated_payload,
                self._improve2_settings.prompt,
                DEFAULT_IMPROVE2_PROMPT,
                "rephrase-2",
            ),
        )
        for slot_key, label, payload_builder, prompt_text, default_prompt, variant in slot_specs:
            profile_key = self._editor_action_profile_defaults.get(slot_key)
            profile = self._model_profile_by_key(profile_key or "")
            if profile is None:
                continue
            requests.append(
                MultiDraftRequest(
                    key=slot_key,
                    label=profile.display_name(),
                    profile=profile,
                    payload_builder=payload_builder,
                    prompt_text=prompt_text,
                    default_prompt=default_prompt,
                    request_title=f"{request_title} {label}",
                    variant=variant,
                )
            )
        return requests

    def _start_multi_draft_choices(
        self,
        *,
        title: str,
        source_start: str,
        source_end: str,
        payload_builder: Callable[[str, ModelProfile], dict[str, Any]],
        prompt_text: str,
        default_prompt: str,
        request_title: str,
    ) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        source_text = self._extract_section_between_headings(doc, source_start, source_end)
        if not source_text:
            self._show_toast(f'Unable to find text between "{source_start}" and "{source_end}".')
            return
        requests = self._build_editor_stacked_choice_requests(
            request_title=request_title,
            payload_builder=payload_builder,
            prompt_text=prompt_text,
            default_prompt=default_prompt,
        )
        if len(requests) != len(STACKED_CHOICE_VARIANTS):
            self._show_toast("Choose Improve, Rephrase 1, and Rephrase 2 profiles in Settings first.")
            return

        self._multi_draft_insert_mode = "editor"
        self._multi_draft_replace_doc = None
        self._multi_draft_replace_start = None
        self._multi_draft_replace_end = None
        self._start_multi_draft_choices_for_text(
            title=title,
            source_text=source_text,
            payload_builder=payload_builder,
            prompt_text=prompt_text,
            default_prompt=default_prompt,
            request_title=request_title,
            requests=requests,
        )

    def _start_multi_draft_choices_for_text(
        self,
        *,
        title: str,
        source_text: str,
        payload_builder: Callable[[str, ModelProfile], dict[str, Any]],
        prompt_text: str,
        default_prompt: str,
        request_title: str,
        requests: list[MultiDraftRequest] | None = None,
    ) -> None:
        self._cancel_multi_draft_run()
        self._clear_pending_regenerate_context()
        if requests is None:
            requests = [
                MultiDraftRequest(
                    key=profile.key,
                    label=profile.display_name(),
                    profile=profile,
                    payload_builder=payload_builder,
                    prompt_text=prompt_text,
                    default_prompt=default_prompt,
                    request_title=request_title,
                )
                for profile in self._model_profiles
            ]
        if not requests:
            self._status_label.set_label(f"{title}: choose Choices profiles in Settings.")
            self._show_toast("Choose Choices profiles in Settings first.")
            return
        self._show_multi_draft_choices(title, requests, source_text)
        self._multi_draft_run_id += 1
        run_id = self._multi_draft_run_id
        cancel_event = threading.Event()
        self._multi_draft_cancel_event = cancel_event
        configured_requests = [request for request in requests if request.profile.is_configured()]
        for request in requests:
            if not request.profile.is_configured():
                self._mark_multi_draft_choice_unconfigured(run_id, request.key)
        if not configured_requests:
            self._status_label.set_label(f"{title}: no configured model profiles.")
            self._show_toast("Configure at least one model profile in Settings first.")
            return

        with self._multi_draft_lock:
            self._multi_draft_remaining = len(configured_requests)
        self._set_busy(True)
        self._status_label.set_label(f"Writing {title.lower()} with {len(configured_requests)} choices…")
        for request in configured_requests:
            thread = threading.Thread(
                target=self._run_multi_draft_choice,
                args=(
                    run_id,
                    source_text,
                    request,
                    cancel_event,
                ),
                daemon=True,
            )
            thread.start()

    def _multi_draft_variant_label(self, variant: str) -> str:
        labels = {
            "improve": "Improve",
            "rephrase": "Rephrase",
            "rephrase-1": "Rephrase 1",
            "rephrase-2": "Rephrase 2",
        }
        return labels.get(variant, variant.replace("-", " ").title())

    def _stacked_choice_default_height(self, source_text: str) -> int:
        text = (source_text or "").strip()
        if not text:
            return STACKED_CHOICES_MIN_HEIGHT

        estimated_lines = sum(
            max(
                1,
                (len(line) + STACKED_CHOICES_ESTIMATED_CHARS_PER_LINE - 1)
                // STACKED_CHOICES_ESTIMATED_CHARS_PER_LINE,
            )
            for line in text.splitlines()
        )

        def scaled_height(value: int, compact_value: int, full_value: int) -> int:
            if value <= compact_value:
                return STACKED_CHOICES_MIN_HEIGHT
            if value >= full_value:
                return STACKED_CHOICES_MAX_HEIGHT
            height_span = STACKED_CHOICES_MAX_HEIGHT - STACKED_CHOICES_MIN_HEIGHT
            value_span = full_value - compact_value
            return STACKED_CHOICES_MIN_HEIGHT + round(
                (value - compact_value) * height_span / value_span
            )

        return max(
            scaled_height(
                len(text),
                STACKED_CHOICES_COMPACT_SOURCE_CHARS,
                STACKED_CHOICES_FULL_SOURCE_CHARS,
            ),
            scaled_height(
                estimated_lines,
                STACKED_CHOICES_COMPACT_SOURCE_LINES,
                STACKED_CHOICES_FULL_SOURCE_LINES,
            ),
        )

    def _show_multi_draft_choices(
        self,
        title: str,
        requests: list[MultiDraftRequest],
        source_text: str,
    ) -> None:
        if self._multi_draft_window is not None:
            self._multi_draft_window.close()
            self._multi_draft_window = None
        app = self.get_application()
        if app is None:
            return

        self._multi_draft_choices = {}
        self._multi_draft_insert_buttons = []
        paired_prompt_choices = bool(requests) and all(
            request.variant in {"improve", "rephrase"} for request in requests
        )
        stacked_prompt_choices = (
            len(requests) == len(STACKED_CHOICE_VARIANTS)
            and (
                tuple(request.variant for request in requests) == STACKED_CHOICE_VARIANTS
                or all(request.stacked for request in requests)
            )
        )

        window_width = 980 if stacked_prompt_choices else 1280
        window_height = (
            self._stacked_choice_default_height(source_text)
            if stacked_prompt_choices
            else STACKED_CHOICES_MAX_HEIGHT
        )
        window = Adw.ApplicationWindow(
            application=app,
            title=title,
            default_width=window_width,
            default_height=window_height,
        )
        window.set_transient_for(self)
        window.connect("close-request", self._on_multi_draft_window_closed)

        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.add_css_class("flat")
        view.add_top_bar(header)

        grid = Gtk.Grid()
        grid.set_hexpand(True)
        grid.set_vexpand(True)
        grid.set_column_homogeneous(not stacked_prompt_choices)
        grid.set_row_homogeneous(stacked_prompt_choices or not paired_prompt_choices)
        grid.set_column_spacing(10)
        grid.set_row_spacing(10)
        grid.set_margin_top(12)
        grid.set_margin_bottom(12)
        grid.set_margin_start(12)
        grid.set_margin_end(12)
        grid.add_css_class("multi-draft-results")
        view.set_content(grid)
        window.set_content(view)
        self._multi_draft_window = window
        self._multi_draft_grid = grid

        if paired_prompt_choices:
            improve_header = Gtk.Label(label="Improve Prompt", xalign=0)
            improve_header.add_css_class("multi-draft-column-header")
            improve_header.set_margin_start(10)
            rephrase_header = Gtk.Label(label="Rephrase Prompt", xalign=0)
            rephrase_header.add_css_class("multi-draft-column-header")
            rephrase_header.set_margin_start(10)
            grid.attach(improve_header, 0, 0, 1, 1)
            grid.attach(rephrase_header, 1, 0, 1, 1)

        for index, request in enumerate(requests):
            profile = request.profile
            style_variant = request.variant
            if style_variant is None and request.stacked and index < len(STACKED_CHOICE_VARIANTS):
                style_variant = STACKED_CHOICE_VARIANTS[index]
            choice_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
            choice_box.set_hexpand(True)
            choice_box.set_vexpand(True)
            choice_box.add_css_class("multi-draft-choice")
            if style_variant:
                choice_box.add_css_class(f"multi-draft-choice-{style_variant}")

            header = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
            header.set_hexpand(True)
            if request.variant:
                variant_label = self._multi_draft_variant_label(request.variant)
                badge = Gtk.Label(label=variant_label)
                badge.add_css_class("multi-draft-variant-badge")
                badge.add_css_class(f"multi-draft-variant-badge-{request.variant}")
                header.append(badge)

            label = Gtk.Label(label=request.label, xalign=0)
            label.add_css_class("heading")
            label.set_tooltip_text(profile.display_name())
            header.append(label)

            status_label = Gtk.Label(label="Waiting…", xalign=0)
            status_label.add_css_class("dim-label")
            status_label.set_hexpand(True)
            header.append(status_label)

            insert_button = Gtk.Button(label="Insert")
            insert_button.add_css_class("flat")
            insert_button.add_css_class("suggested-action")
            insert_button.set_sensitive(False)
            insert_button.connect("clicked", self._on_multi_draft_insert_clicked, request.key)
            header.append(insert_button)
            choice_box.append(header)

            buffer = Gtk.TextBuffer()
            text_view = Gtk.TextView.new_with_buffer(buffer)
            text_view.set_editable(False)
            text_view.set_cursor_visible(False)
            text_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
            text_view.set_hexpand(True)
            text_view.set_vexpand(True)
            text_view.set_left_margin(8)
            text_view.set_right_margin(8)
            text_view.set_top_margin(8)
            text_view.set_bottom_margin(8)
            text_view.add_css_class("spelling-output-view")

            scroller = Gtk.ScrolledWindow()
            scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
            scroller.set_hexpand(True)
            scroller.set_vexpand(True)
            scroller.set_child(text_view)
            scroller.add_css_class("multi-draft-choice-output")
            choice_box.append(scroller)

            if stacked_prompt_choices:
                grid.attach(choice_box, 0, index, 1, 1)
            else:
                row_offset = 1 if paired_prompt_choices else 0
                grid.attach(choice_box, index % 2, index // 2 + row_offset, 1, 1)
            self._multi_draft_insert_buttons.append(insert_button)
            self._multi_draft_choices[request.key] = MultiDraftChoice(
                key=request.key,
                label=request.label,
                profile=profile,
                buffer=buffer,
                text_view=text_view,
                status_label=status_label,
                insert_button=insert_button,
                variant=request.variant,
                source_text=source_text,
            )
        window.present()

    def _on_multi_draft_window_closed(self, window: Adw.ApplicationWindow, *_args: object) -> bool:
        if self._multi_draft_window is window:
            self._cancel_multi_draft_run()
            self._multi_draft_window = None
            self._multi_draft_grid = None
            self._multi_draft_insert_mode = "editor"
            self._multi_draft_replace_doc = None
            self._multi_draft_replace_start = None
            self._multi_draft_replace_end = None
        return False

    def _mark_multi_draft_choice_unconfigured(self, run_id: int, profile_key: str) -> bool:
        if run_id != self._multi_draft_run_id:
            return False
        choice = self._multi_draft_choices.get(profile_key)
        if choice is None:
            return False
        choice.failed = True
        choice.complete = True
        choice.status_label.set_label("Not configured")
        choice.buffer.set_text("Configure this model profile in Settings to include it here.")
        choice.text_view.set_editable(False)
        choice.text_view.set_cursor_visible(False)
        choice.insert_button.set_sensitive(False)
        return False

    def _append_multi_draft_choice_text(self, run_id: int, profile_key: str, text: str) -> bool:
        if run_id != self._multi_draft_run_id or not text:
            return False
        choice = self._multi_draft_choices.get(profile_key)
        if choice is None:
            return False
        choice.text += text
        end_iter = choice.buffer.get_end_iter()
        choice.buffer.insert(end_iter, text)
        choice.status_label.set_label("Writing…")
        return False

    def _changed_text_word_spans(self, text: str) -> list[tuple[str, int, int]]:
        return [
            (match.group(0).casefold(), match.start(), match.end())
            for match in re.finditer(r"[^\W_]+(?:['’][^\W_]+)?", text, flags=re.UNICODE)
        ]

    def _changed_text_unit_boundary(self, char: str) -> bool:
        return char.isspace() or char in ',;:.!?()[]{}<>"“”'

    def _changed_text_connector_punctuation(self, text: str) -> bool:
        return any(char in "-‐‑‒–—―/" for char in text)

    def _merge_changed_text_ranges(self, ranges: list[tuple[int, int]]) -> list[tuple[int, int]]:
        if not ranges:
            return []
        merged: list[tuple[int, int]] = []
        for start, end in sorted(ranges):
            if start >= end:
                continue
            if not merged or start > merged[-1][1]:
                merged.append((start, end))
                continue
            previous_start, previous_end = merged[-1]
            merged[-1] = (previous_start, max(previous_end, end))
        return merged

    def _expand_changed_text_punctuation_range(
        self,
        text: str,
        start_offset: int,
        end_offset: int,
    ) -> tuple[int, int]:
        text_len = len(text)
        start = max(0, min(start_offset, text_len))
        end = max(start, min(end_offset, text_len))
        changed_text = text[start:end]

        if not changed_text:
            if start < text_len:
                end = start + 1
            elif start > 0:
                start -= 1
                end = start + 1
            else:
                return (0, 0)

        should_expand = self._changed_text_connector_punctuation(changed_text)
        if not should_expand:
            return (start, end)

        while start > 0 and not self._changed_text_unit_boundary(text[start - 1]):
            start -= 1
        while end < text_len and not self._changed_text_unit_boundary(text[end]):
            end += 1
        return (start, end)

    def _changed_text_punctuation_ranges(self, source_text: str, output_text: str) -> list[tuple[int, int]]:
        if source_text == output_text:
            return []
        if re.sub(r"\s+", " ", source_text).strip() == re.sub(r"\s+", " ", output_text).strip():
            return []

        matcher = difflib.SequenceMatcher(None, source_text, output_text, autojunk=False)
        ranges: list[tuple[int, int]] = []
        for tag, _source_start, _source_end, output_start, output_end in matcher.get_opcodes():
            if tag == "equal":
                continue
            ranges.append(
                self._expand_changed_text_punctuation_range(
                    output_text,
                    output_start,
                    output_end,
                )
            )
        return self._merge_changed_text_ranges(ranges)

    def _changed_text_word_ranges(self, source_text: str, output_text: str) -> list[tuple[int, int]]:
        source_words = self._changed_text_word_spans(source_text)
        output_words = self._changed_text_word_spans(output_text)
        ranges: list[tuple[int, int]] = []
        if source_words and output_words:
            matcher = difflib.SequenceMatcher(
                None,
                [word for word, _start, _end in source_words],
                [word for word, _start, _end in output_words],
                autojunk=False,
            )
            for tag, _source_start, _source_end, output_start, output_end in matcher.get_opcodes():
                if tag == "equal" or output_start == output_end:
                    continue
                ranges.append((output_words[output_start][1], output_words[output_end - 1][2]))
        if ranges or source_text == output_text:
            return ranges
        return self._changed_text_punctuation_ranges(source_text, output_text)

    def _changed_text_display_parts(
        self,
        source_text: str,
        output_text: str,
        show_deletions: bool = True,
    ) -> tuple[str, list[tuple[int, int]], list[tuple[int, int]]]:
        source_words = self._changed_text_word_spans(source_text)
        output_words = self._changed_text_word_spans(output_text)
        if not source_words or not output_words:
            return output_text, [], []

        matcher = difflib.SequenceMatcher(
            None,
            [word for word, _start, _end in source_words],
            [word for word, _start, _end in output_words],
            autojunk=False,
        )
        display_parts: list[str] = []
        changed_ranges: list[tuple[int, int]] = []
        deletion_marker_ranges: list[tuple[int, int]] = []
        last_output_offset = 0

        for tag, source_start, source_end, output_start, output_end in matcher.get_opcodes():
            output_offset = output_words[output_start][1] if output_start < len(output_words) else len(output_text)
            if output_offset > last_output_offset:
                display_parts.append(output_text[last_output_offset:output_offset])
                last_output_offset = output_offset

            display_offset = sum(len(part) for part in display_parts)
            if tag != "equal" and output_start < output_end:
                changed_start = display_offset
                changed_end = display_offset + output_words[output_end - 1][2] - output_offset
                changed_ranges.append((changed_start, changed_end))

            if show_deletions and tag == "delete":
                deleted_words = [
                    source_text[start:end]
                    for _word, start, end in source_words[source_start:source_end]
                ]
                marker = f"[removed: {' '.join(deleted_words)}]"
                if display_parts and not display_parts[-1].endswith((" ", "\n", "\t")):
                    marker = f" {marker}"
                if output_offset < len(output_text) and not output_text[output_offset].isspace():
                    marker = f"{marker} "
                marker_start = display_offset
                display_parts.append(marker)
                deletion_marker_ranges.append((marker_start, marker_start + len(marker)))

        if last_output_offset < len(output_text):
            display_parts.append(output_text[last_output_offset:])

        return "".join(display_parts), changed_ranges, deletion_marker_ranges

    def _set_multi_draft_choice_display_text(self, choice: MultiDraftChoice, text: str) -> None:
        if choice.variant not in MULTI_DRAFT_DIFF_VARIANTS:
            choice.buffer.set_text(text)
            return

        display_text, changed_ranges, deletion_marker_ranges = self._changed_text_display_parts(
            choice.source_text,
            text,
            show_deletions=False,
        )
        choice.buffer.set_text(display_text)

        changed_tag = choice.buffer.create_tag(None, weight=Pango.Weight.BOLD)
        deletion_marker_tag = None
        if deletion_marker_ranges:
            deletion_marker_tag = choice.buffer.create_tag(
                None,
                weight=Pango.Weight.BOLD,
                style=Pango.Style.ITALIC,
            )
        choice.deletion_marker_tag = deletion_marker_tag

        if changed_tag is not None:
            for start_offset, end_offset in changed_ranges:
                if start_offset >= end_offset:
                    continue
                start_iter = choice.buffer.get_iter_at_offset(start_offset)
                end_iter = choice.buffer.get_iter_at_offset(end_offset)
                choice.buffer.apply_tag(changed_tag, start_iter, end_iter)

        if deletion_marker_tag is None:
            return
        for start_offset, end_offset in deletion_marker_ranges:
            if start_offset >= end_offset:
                continue
            start_iter = choice.buffer.get_iter_at_offset(start_offset)
            end_iter = choice.buffer.get_iter_at_offset(end_offset)
            choice.buffer.apply_tag(deletion_marker_tag, start_iter, end_iter)

    def _finish_multi_draft_choice(self, run_id: int, profile_key: str, error: str | None = None) -> bool:
        if run_id != self._multi_draft_run_id:
            return False
        choice = self._multi_draft_choices.get(profile_key)
        if choice is not None:
            choice.complete = True
            if error:
                choice.failed = True
                choice.status_label.set_label("Failed")
                choice.buffer.set_text(f"Unable to generate this draft: {error}")
                choice.text_view.set_editable(False)
                choice.text_view.set_cursor_visible(False)
                choice.insert_button.set_sensitive(False)
            else:
                choice.text = self._normalize_generated_output_text(choice.text)
                self._set_multi_draft_choice_display_text(choice, choice.text)
                if choice.text:
                    choice.status_label.set_label("Ready")
                    choice.text_view.set_editable(True)
                    choice.text_view.set_cursor_visible(True)
                    choice.insert_button.set_sensitive(True)
                else:
                    choice.failed = True
                    choice.status_label.set_label("No text returned")
                    choice.text_view.set_editable(False)
                    choice.text_view.set_cursor_visible(False)
                    choice.insert_button.set_sensitive(False)

        remaining = 0
        with self._multi_draft_lock:
            self._multi_draft_remaining = max(0, self._multi_draft_remaining - 1)
            remaining = self._multi_draft_remaining
        if remaining == 0:
            self._set_busy(False)
            self._status_label.set_label("Draft choices complete.")
        return False

    def _run_multi_draft_choice(
        self,
        run_id: int,
        source_text: str,
        request: MultiDraftRequest,
        cancel_event: threading.Event,
    ) -> None:
        error = None
        profile = request.profile
        try:
            title = f"{request.request_title} - {request.label}"
            if self._is_gemini_generate_content_url(profile.api_url):
                if cancel_event.is_set():
                    return
                prompt = _expand_shared_prompt_parts(request.prompt_text or request.default_prompt)
                combined = f"{prompt}\n\n{source_text}" if prompt else source_text
                output = self._call_gemini_generate_content(
                    profile.api_url,
                    profile.api_key,
                    combined,
                    request_title=title,
                )
                if not cancel_event.is_set():
                    GLib.idle_add(self._append_multi_draft_choice_text, run_id, request.key, output)
            else:
                payload = request.payload_builder(source_text, profile)
                for chunk in self._stream_custom(
                    payload,
                    profile.api_url,
                    profile.api_key,
                    request_title=title,
                ):
                    if cancel_event.is_set():
                        return
                    GLib.idle_add(self._append_multi_draft_choice_text, run_id, request.key, chunk)
        except Exception as exc:  # noqa: BLE001
            error = str(exc)
        if not cancel_event.is_set():
            GLib.idle_add(self._finish_multi_draft_choice, run_id, request.key, error)

    def _cancel_multi_draft_run(self) -> None:
        cancel_event = self._multi_draft_cancel_event
        if cancel_event is not None:
            cancel_event.set()
            self._multi_draft_cancel_event = None
        self._multi_draft_run_id += 1
        with self._multi_draft_lock:
            self._multi_draft_remaining = 0
        self._set_busy(False)
        for insert_button in self._multi_draft_insert_buttons:
            insert_button.set_sensitive(False)

    def _get_multi_draft_choice_text(self, choice: MultiDraftChoice) -> str:
        tag = choice.deletion_marker_tag
        if tag is None:
            start_iter, end_iter = choice.buffer.get_bounds()
            text = choice.buffer.get_text(start_iter, end_iter, True)
            return self._normalize_generated_output_text(text)

        pieces: list[str] = []
        cursor = choice.buffer.get_start_iter()
        end_iter = choice.buffer.get_end_iter()
        while cursor.compare(end_iter) < 0:
            next_iter = cursor.copy()
            if not next_iter.forward_to_tag_toggle(tag):
                next_iter = end_iter.copy()
            if not cursor.has_tag(tag):
                pieces.append(choice.buffer.get_text(cursor, next_iter, True))
            cursor = next_iter
        text = "".join(pieces)
        return self._normalize_generated_output_text(text)

    def _multi_draft_choice_display_label(self, choice: MultiDraftChoice) -> str:
        if choice.variant:
            return f"{choice.label} {self._multi_draft_variant_label(choice.variant)}"
        return choice.label

    def _on_multi_draft_insert_clicked(self, _button: Gtk.Button, profile_key: str) -> None:
        choice = self._multi_draft_choices.get(profile_key)
        if choice is None:
            return
        insert_text = self._get_multi_draft_choice_text(choice)
        if not insert_text:
            return
        choice.text = insert_text
        if self._multi_draft_insert_mode in {"text-draft-terminal", "text-draft-terminal-replace-generated"}:
            if not self._text_draft_terminal_target_available():
                self._show_toast("Embedded Codex is no longer running.")
                return
            if (
                self._multi_draft_insert_mode == "text-draft-terminal-replace-generated"
                and not self._delete_text_draft_terminal_generated_output()
            ):
                return
            if not self._feed_text_draft_terminal_text(insert_text):
                return
            self._set_text_draft_terminal_generated_output(insert_text)
            self._finish_text_draft_terminal_choice_insert(choice)
            return
        if self._multi_draft_insert_mode == "text-draft-replace-generated":
            if not self._prepare_text_draft_generated_replace():
                self._show_toast("Unable to replace the last generated draft output.")
                return
            self._append_text_draft_inserted_text(insert_text)
            self._finish_text_draft_multi_choice_insert(choice)
            return
        if self._multi_draft_insert_mode == "text-draft-replace-selection":
            if not self._prepare_text_draft_saved_selection_replace():
                self._show_toast("Unable to replace the selected Draft text.")
                return
            self._append_text_draft_inserted_text(insert_text)
            self._finish_text_draft_multi_choice_insert(choice)
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        if self._multi_draft_insert_mode == "replace-selection":
            if not self._prepare_multi_draft_selection_replacement(doc):
                self._show_toast("Unable to replace the selected text.")
                return
            self._append_improve1_text(insert_text)
            if self._improve_insert_doc and self._improve_insert_cursor:
                self._flush_pending_newlines(
                    self._improve_insert_doc,
                    self._improve_insert_cursor,
                    "_improve_pending_newlines",
                )
                self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
                self._apply_case_citation_italics_to_recent_writer_insert(
                    self._improve_insert_doc,
                    self._improve_insert_cursor,
                )
            self._capture_improve1_range_end()
            for insert_button in self._multi_draft_insert_buttons:
                insert_button.set_sensitive(False)
            self._status_label.set_label(f"Inserted {self._multi_draft_choice_display_label(choice)} draft.")
            self._close_multi_draft_window_after_insert()
            return
        if self._multi_draft_insert_mode == "replace-generated":
            if not self._prepare_improve_insertion(doc):
                self._show_toast("Unable to prepare Improve Generated insertion point.")
                return
            self._append_improve1_text(insert_text)
            if self._improve_insert_doc and self._improve_insert_cursor:
                self._flush_pending_newlines(
                    self._improve_insert_doc,
                    self._improve_insert_cursor,
                    "_improve_pending_newlines",
                )
                self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
                self._apply_case_citation_italics_to_recent_writer_insert(
                    self._improve_insert_doc,
                    self._improve_insert_cursor,
                )
            self._capture_improve1_range_end()
            for insert_button in self._multi_draft_insert_buttons:
                insert_button.set_sensitive(False)
            self._status_label.set_label(f"Inserted {self._multi_draft_choice_display_label(choice)} draft.")
            self._close_multi_draft_window_after_insert()
            return
        if not self._prepare_editor_insertion(doc):
            self._show_toast("Unable to prepare Writer insertion point.")
            return
        self._append_editor_text(insert_text)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._capture_spellingstyle_range_end()
        self._status_label.set_label(f"Inserted {self._multi_draft_choice_display_label(choice)} draft.")
        self._close_multi_draft_window_after_insert()

    def _finish_text_draft_multi_choice_insert(self, choice: MultiDraftChoice) -> None:
        self._flush_text_draft_pending_newlines(is_original_output=False)
        self._ensure_single_text_draft_trailing_space()
        self._focus_text_draft_insert_end()
        self._text_draft_temp_dirty = True
        for insert_button in self._multi_draft_insert_buttons:
            insert_button.set_sensitive(False)
        self._status_label.set_label(f"Inserted {self._multi_draft_choice_display_label(choice)} draft.")
        self._close_multi_draft_window_after_insert()

    def _finish_text_draft_terminal_choice_insert(self, choice: MultiDraftChoice) -> None:
        for insert_button in self._multi_draft_insert_buttons:
            insert_button.set_sensitive(False)
        self._status_label.set_label(
            f"Streamed {self._multi_draft_choice_display_label(choice)} draft to embedded Codex."
        )
        self._close_multi_draft_window_after_insert()

    def _close_multi_draft_window_after_insert(self) -> None:
        self._cancel_multi_draft_run()
        window = self._multi_draft_window
        if window is not None:
            window.close()

    def _prepare_multi_draft_selection_replacement(self, _doc: XTextDocument) -> bool:  # type: ignore[type-arg]
        replace_doc = self._multi_draft_replace_doc
        selection_start = self._multi_draft_replace_start
        selection_end = self._multi_draft_replace_end
        if replace_doc is None or selection_start is None or selection_end is None:
            return False
        try:
            text = self._get_text_container(replace_doc, selection_start)
            if not text:
                return False
            range_cursor = text.createTextCursorByRange(selection_start)
            range_cursor.gotoRange(selection_end, True)
            try:
                range_cursor.setString("")
            except Exception:
                pass
            self._improve_insert_cursor = text.createTextCursorByRange(range_cursor)
            self._improve_insert_doc = replace_doc
            self._editor_insert_end = None
            self._last_insert_len = 0
            self._improve_pending_newlines = 0
            self._improve_started = False
            return True
        except Exception:
            return False


    # Thesaurus pipeline -------------------------------------------------
    def _run_thesaurus(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_thesaurus_payload(source_text, profile)
        try:
            words = self._call_thesaurus(payload, profile)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_thesaurus_failed, str(exc))
            return
        GLib.idle_add(self._on_thesaurus_ready, words)

    def _compose_thesaurus_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        prompt = _expand_shared_prompt_parts(self._thesaurus_settings.prompt or DEFAULT_THESAURUS_PROMPT)
        content = f"{prompt}\n\n{source_text}" if prompt else source_text
        payload = {
            "messages": [
                {"role": "user", "content": content},
            ],
            "stream": False,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _call_thesaurus(self, payload: dict[str, Any], profile: ModelProfile) -> list[str]:
        raw = self._post_json_and_read(
            payload,
            profile.api_url,
            profile.api_key,
            request_title="Thesaurus",
        )
        parts = list(self._extract_response_text(raw))
        if not parts:
            raise ValueError("Thesaurus returned empty output.")
        return self._parse_thesaurus_response("".join(parts).strip())

    def _parse_thesaurus_response(self, text: str) -> list[str]:
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Thesaurus response was not valid JSON: {exc}") from exc
        words = None
        if isinstance(data, dict):
            words = data.get("alternatives") or data.get("synonyms") or data.get("words")
        elif isinstance(data, list):
            words = data
        if not isinstance(words, list):
            raise ValueError("Thesaurus response did not include an alternatives list.")
        cleaned: list[str] = []
        seen: set[str] = set()
        for item in words:
            word = str(item).strip()
            if not word or word in seen:
                continue
            seen.add(word)
            cleaned.append(word)
        if not cleaned:
            raise ValueError("Thesaurus response contained no alternatives.")
        return cleaned

    def _on_thesaurus_ready(self, words: list[str]) -> bool:
        self._set_busy(False)
        self._thesaurus_words = words
        self._set_thesaurus_placeholder("")
        self._render_thesaurus_words()
        self._status_label.set_label("Thesaurus alternatives ready.")
        return False

    def _on_thesaurus_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._set_thesaurus_placeholder("Unable to load alternatives.")
        self._thesaurus_words = []
        self._render_thesaurus_words()
        self._notify_llm_error(message)
        return False

    # Reference pipeline ------------------------------------------------
    def _run_reference(self, source_text: str, prompt_override: str | None) -> None:
        search_query = f"Define the following word or phrase and find authoritative sources: {source_text}"
        try:
            search_bundle = self._run_tavily_search(
                search_query,
                self._reference_settings.tavily_api_key,
                request_title="Look Up",
            )
            payload = self._compose_reference_payload(source_text, prompt_override, search_bundle)
            received = False
            if self._endpoint_uses_responses_api(self._reference_settings.api_url):
                stream = self._stream_responses(
                    payload,
                    self._reference_settings.api_url,
                    self._reference_settings.api_key,
                    request_title="Look Up",
                )
            else:
                stream = self._stream_custom(
                    payload,
                    self._reference_settings.api_url,
                    self._reference_settings.api_key,
                    request_title="Look Up",
                )
            for chunk in stream:
                received = True
                GLib.idle_add(self._append_reference_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_reference_failed, str(exc))
            return
        if not received:
            GLib.idle_add(self._on_reference_failed, "Reference returned empty output.")
            return
        GLib.idle_add(self._on_reference_finished, "Reference ready.")

    def _run_ask(self, question: str) -> None:
        try:
            search_bundle = self._run_tavily_search(
                question,
                self._ask_settings.tavily_api_key,
                request_title="Ask",
            )
            payload = self._compose_ask_payload(question, search_bundle)
            received = False
            if self._endpoint_uses_responses_api(self._ask_settings.api_url):
                stream = self._stream_responses(
                    payload,
                    self._ask_settings.api_url,
                    self._ask_settings.api_key,
                    request_title="Ask",
                )
            else:
                stream = self._stream_custom(
                    payload,
                    self._ask_settings.api_url,
                    self._ask_settings.api_key,
                    request_title="Ask",
                )
            for chunk in stream:
                received = True
                GLib.idle_add(self._append_reference_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_reference_failed, str(exc))
            return
        if not received:
            GLib.idle_add(self._on_reference_failed, "Ask returned empty output.")
            return
        GLib.idle_add(self._on_reference_finished, "Answer ready.")

    def _endpoint_uses_responses_api(self, api_url: str) -> bool:
        return api_url.rstrip("/").endswith("/responses")

    def _run_tavily_search(
        self,
        query: str,
        tavily_api_key: str,
        *,
        request_title: str,
    ) -> TavilySearchBundle:
        tvly_path = _resolve_tavily_cli_path()
        if not tvly_path:
            raise ValueError(f"Tavily CLI (`tvly`) was not found on PATH. {TAVILY_CLI_INSTALL_HINT}")
        command = [
            tvly_path,
            "search",
            query,
            "--depth",
            "advanced",
            "--max-results",
            str(TAVILY_MAX_RESULTS),
            "--include-raw-content",
            "text",
            "--json",
        ]
        env = os.environ.copy()
        env["TAVILY_API_KEY"] = tavily_api_key
        self._remember_editor_request(
            f"{request_title} Tavily Search",
            tvly_path,
            {
                "command": command[1:],
                "query": query,
            },
        )
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                env=env,
                check=False,
            )
        except OSError as exc:
            raise ValueError(f"Unable to run Tavily CLI: {exc}. {TAVILY_CLI_INSTALL_HINT}") from exc
        detail = completed.stderr.strip() or completed.stdout.strip()
        if completed.returncode != 0:
            if completed.returncode == 3:
                raise ValueError("Tavily CLI authentication failed. Check the Tavily API key in Settings.")
            raise ValueError(f"Tavily CLI search failed: {detail or f'exit code {completed.returncode}'}")
        raw = completed.stdout.strip()
        if not raw:
            raise ValueError("Tavily CLI search returned empty JSON output.")
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Tavily CLI search returned invalid JSON: {exc}") from exc
        sources = self._select_tavily_sources(data)
        if not sources:
            raise ValueError("Tavily CLI did not return any usable sources.")
        notice = None
        if len(sources) < TAVILY_MAX_SOURCES:
            notice = "Only one usable Tavily source was available for this query. Do not invent additional citations."
        return TavilySearchBundle(query=query, sources=sources, notice=notice)

    def _select_tavily_sources(self, data: Any) -> list[TavilySearchSource]:
        raw_results = data.get("results") if isinstance(data, dict) else data
        if not isinstance(raw_results, list):
            return []
        usable: list[TavilySearchSource] = []
        for item in raw_results:
            if not isinstance(item, dict):
                continue
            url = str(item.get("url") or "").strip()
            title = str(item.get("title") or item.get("site_name") or "").strip()
            excerpt = str(item.get("raw_content") or item.get("content") or item.get("snippet") or "").strip()
            if not url or not title or not excerpt:
                continue
            domain = urllib.parse.urlparse(url).netloc.lower().strip()
            usable.append(
                TavilySearchSource(
                    title=title,
                    url=url,
                    excerpt=self._trim_tavily_excerpt(excerpt),
                    domain=domain,
                )
            )
        selected: list[TavilySearchSource] = []
        seen_domains: set[str] = set()
        for source in usable:
            if source.domain and source.domain in seen_domains:
                continue
            selected.append(source)
            if source.domain:
                seen_domains.add(source.domain)
            if len(selected) >= TAVILY_MAX_SOURCES:
                return selected
        for source in usable:
            if any(existing.url == source.url for existing in selected):
                continue
            selected.append(source)
            if len(selected) >= TAVILY_MAX_SOURCES:
                break
        return selected

    def _trim_tavily_excerpt(self, text: str) -> str:
        normalized = re.sub(r"\n{3,}", "\n\n", text or "").strip()
        if len(normalized) <= TAVILY_SOURCE_EXCERPT_CHARS:
            return normalized
        truncated = normalized[:TAVILY_SOURCE_EXCERPT_CHARS].rstrip()
        if " " in truncated:
            truncated = truncated.rsplit(" ", 1)[0].rstrip()
        return f"{truncated}..."

    def _format_tavily_sources_for_prompt(self, sources: list[TavilySearchSource]) -> str:
        blocks: list[str] = []
        for index, source in enumerate(sources, start=1):
            blocks.append(
                "\n".join(
                    (
                        f"SOURCE {index}",
                        f"Title: {source.title}",
                        f"URL: {source.url}",
                        "Excerpt:",
                        source.excerpt,
                    )
                )
            )
        return "\n\n".join(blocks)

    def _compose_search_grounded_payload(
        self,
        *,
        api_url: str,
        model_id: str,
        disable_reasoning: bool,
        system_prompt: str,
        user_content: str,
    ) -> dict[str, Any]:
        if self._endpoint_uses_responses_api(api_url):
            payload = {
                "input": f"{system_prompt}\n\n{user_content}".strip(),
                "stream": True,
            }
        else:
            payload = {
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_content},
                ],
                "stream": True,
            }
        return self._add_model_id(
            payload,
            model_id,
            disable_reasoning=disable_reasoning,
        )

    def _compose_reference_payload(
        self,
        source_text: str,
        prompt_override: str | None,
        search_bundle: TavilySearchBundle,
    ) -> dict[str, Any]:
        prompt = _expand_shared_prompt_parts(prompt_override or self._reference_settings.prompt or DEFAULT_REFERENCE_PROMPT)
        system_prompt = (
            "The app already performed the Tavily search. "
            "Use only the provided search results. "
            "Do not claim to have browsed or searched on your own. "
            "Cite only the provided source URLs.\n\n"
            f"{prompt}"
        ).strip()
        user_parts = [
            f"Word or phrase: {source_text}",
        ]
        if search_bundle.notice:
            user_parts.append(f"SEARCH NOTICE\n{search_bundle.notice}")
        user_parts.extend(
            (
                "TAVILY SEARCH RESULTS",
                self._format_tavily_sources_for_prompt(search_bundle.sources),
            )
        )
        user_content = "\n\n".join(part for part in user_parts if part)
        return self._compose_search_grounded_payload(
            api_url=self._reference_settings.api_url,
            model_id=self._reference_settings.model_id,
            disable_reasoning=self._reference_settings.disable_reasoning,
            system_prompt=system_prompt,
            user_content=user_content,
        )

    def _compose_ask_payload(self, question: str, search_bundle: TavilySearchBundle) -> dict[str, Any]:
        prompt = _expand_shared_prompt_parts(self._ask_settings.prompt or DEFAULT_ASK_PROMPT)
        today = datetime.now().strftime("%B %d, %Y")
        system_prompt = (
            "The app already performed the Tavily search. "
            "Use only the provided search results. "
            "Do not claim to have browsed or searched on your own. "
            "Cite only the provided source URLs.\n\n"
            f"{prompt}"
        ).strip()
        user_parts = [
            f"Today is {today}.",
            f"Question: {question}",
        ]
        if search_bundle.notice:
            user_parts.append(f"SEARCH NOTICE\n{search_bundle.notice}")
        user_parts.extend(
            (
                "TAVILY SEARCH RESULTS",
                self._format_tavily_sources_for_prompt(search_bundle.sources),
            )
        )
        user_content = "\n\n".join(part for part in user_parts if part)
        return self._compose_search_grounded_payload(
            api_url=self._ask_settings.api_url,
            model_id=self._ask_settings.model_id,
            disable_reasoning=self._ask_settings.disable_reasoning,
            system_prompt=system_prompt,
            user_content=user_content,
        )

    def _on_reference_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        self._trim_reference_output_edges()
        return False

    def _on_reference_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._set_reference_placeholder("Unable to load reference.")
        self._notify_llm_error(message)
        return False

    def _set_thesaurus_placeholder(self, text: str) -> None:
        if hasattr(self, "_thesaurus_placeholder"):
            self._thesaurus_placeholder.set_label(text)
            self._thesaurus_placeholder.set_visible(bool(text) and not self._thesaurus_words)

    def _show_thesaurus_output(self) -> None:
        if hasattr(self, "_thesaurus_stack"):
            self._thesaurus_stack.set_visible_child_name("thesaurus")
        self._set_reference_toggle_state("thesaurus")

    def _show_reference_output(self, *, mode: str = "lookup") -> None:
        if hasattr(self, "_thesaurus_stack"):
            self._thesaurus_stack.set_visible_child_name("reference")
        self._set_reference_toggle_state(mode)

    def _set_reference_toggle_state(self, active: str) -> None:
        self._reference_output_mode = active
        if not hasattr(self, "_thesaurus_btn") or not hasattr(self, "_reference_btn"):
            return
        if active == "lookup":
            self._reference_btn.add_css_class("reference-toggle-active")
            self._thesaurus_btn.remove_css_class("reference-toggle-active")
        elif active == "thesaurus":
            self._thesaurus_btn.add_css_class("reference-toggle-active")
            self._reference_btn.remove_css_class("reference-toggle-active")
        else:
            self._thesaurus_btn.remove_css_class("reference-toggle-active")
            self._reference_btn.remove_css_class("reference-toggle-active")

    def _set_reference_placeholder(self, text: str) -> None:
        self._reference_output_text = text or ""
        self._reference_placeholder_active = True
        if hasattr(self, "_reference_label"):
            self._reference_label.add_css_class("dim-label")
            self._reference_label.set_text(text or "")

    def _set_reference_output_text(self, text: str) -> bool:
        self._reference_output_text = text or ""
        self._reference_placeholder_active = False
        self._render_reference_output()
        return False

    def _append_reference_output_text(self, text: str) -> bool:
        if not text:
            return False
        if self._reference_placeholder_active:
            self._reference_output_text = ""
            self._reference_placeholder_active = False
            if hasattr(self, "_reference_label"):
                self._reference_label.remove_css_class("dim-label")
        self._reference_output_text += text
        self._render_reference_output()
        return False

    def _render_reference_output(self) -> None:
        if not hasattr(self, "_reference_label"):
            return
        markup = self._format_reference_markup(self._reference_output_text)
        if markup:
            self._reference_label.remove_css_class("dim-label")
            self._reference_label.set_markup(markup)
        else:
            self._reference_label.add_css_class("dim-label")
            self._reference_label.set_text("No reference yet.")

    def _format_reference_markup(self, text: str) -> str:
        if not text:
            return ""
        parts: list[str] = []
        last = 0
        for match in REFERENCE_URL_RE.finditer(text):
            start, end = match.span()
            if start > last:
                parts.append(GLib.markup_escape_text(text[last:start]))
            url = match.group(0)
            escaped_url = GLib.markup_escape_text(url)
            parts.append(f'<a href="{escaped_url}">{escaped_url}</a>')
            last = end
        if last < len(text):
            parts.append(GLib.markup_escape_text(text[last:]))
        return "".join(parts)

    def _on_reference_link_clicked(self, _label: Gtk.Label, uri: str) -> bool:
        try:
            Gio.AppInfo.launch_default_for_uri(uri, None)
        except Exception:
            self._show_toast("Unable to open link.")
        return True

    def _render_thesaurus_words(self) -> None:
        if not hasattr(self, "_thesaurus_list"):
            return
        for row in self._thesaurus_rows:
            self._thesaurus_list.remove(row)
        self._thesaurus_rows.clear()
        if hasattr(self, "_thesaurus_placeholder"):
            self._thesaurus_placeholder.set_visible(bool(self._thesaurus_placeholder.get_label()) and not self._thesaurus_words)
        for word in self._thesaurus_words:
            row = self._build_thesaurus_row(word)
            self._thesaurus_list.append(row)
            self._thesaurus_rows.append(row)

    def _build_thesaurus_row(self, word: str) -> Gtk.Box:
        row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        row.add_css_class("thesaurus-row")
        row.set_margin_top(6)
        row.set_margin_bottom(6)
        row.set_margin_start(SPELLING_OUTPUT_PADDING_PX)
        row.set_margin_end(SPELLING_OUTPUT_PADDING_PX)
        label = Gtk.Label(label=word, xalign=0)
        label.set_wrap(True)
        label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        label.set_hexpand(True)
        row.append(label)
        use_btn = Gtk.Button(label="Use")
        use_btn.add_css_class("flat")
        use_btn.add_css_class("thesaurus-use-button")
        use_btn.connect("clicked", self._on_thesaurus_word_clicked, word)
        row.append(use_btn)
        click = Gtk.GestureClick()
        click.connect("released", lambda *_args: self._on_thesaurus_word_clicked(None, word))
        row.add_controller(click)
        return row

    def _on_thesaurus_word_clicked(self, _button: Gtk.Button | None, word: str) -> None:
        self._replace_selected_text_no_trailing_space(word)

    def _replace_selected_text_no_trailing_space(self, text: str) -> None:
        if self._busy:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("Unable to reach LibreOffice listener. Is the service running?")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document (File → Launch Writer).")
            return
        view_cursor = doc.getCurrentController().getViewCursor()
        if not view_cursor.getString().strip():
            self._show_toast("Select text in Writer first.")
            return
        if not self._prepare_selection_insertion(doc, view_cursor):
            self._show_toast("Unable to prepare selected text replacement.")
            return
        self._set_busy(True)
        self._status_label.set_label("Replacing selected text…")
        self._append_improve1_text(text)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._trim_trailing_whitespace(self._improve_insert_doc, self._improve_insert_cursor)
        self._capture_improve1_range_end()
        self._set_busy(False)
        self._status_label.set_label("Selected text replaced.")

    def _wrap_text_in_curly_quotes(self, text: str) -> str:
        normalized = self._normalize_curly_quotes(text)
        return f"“{normalized}”"

    def _append_case_to_concordance(self, citation: str) -> None:
        concordance_file = self._concordance_file_path
        if not concordance_file:
            raise FileNotFoundError("No concordance file configured.")
        concordance_file.parent.mkdir(parents=True, exist_ok=True)
        short = re.sub(r" \(\d{4}\) ", ",supra, ", citation)
        short = re.sub(r"Cal\.App\.\d+[a-z]*.*", lambda m: m.group(0).split()[0], short)
        short = re.sub(r"Cal\.\d+[a-z]*.*", lambda m: m.group(0).split()[0], short)
        short = re.sub(r"U\.S\..*", "U.S.", short)
        lines = [
            f"{citation};{citation};Cases;;0;0",
            f"{short};{citation};Cases;;0;0",
        ]
        with concordance_file.open("a", encoding="utf-8") as handle:
            for line in lines:
                handle.write("\n" + line)

    def _autotext_display_name(self, citation: str) -> str:
        if citation.startswith("In re "):
            return citation[len("In re ") :]
        if citation.startswith("People v. "):
            return citation[len("People v. ") :]
        return citation

    def _generate_autotext_key(self, display_name: str, existing: set[str]) -> str:
        base = re.sub(r"[^A-Za-z0-9]+", "", display_name).upper() or "CASE"
        base = base[:6]
        for idx in range(1, 10000):
            candidate = f"{base}{idx:03d}"
            if candidate not in existing:
                return candidate
        raise RuntimeError("Unable to generate a unique AutoText key.")

    def _split_citation_for_italics(self, citation: str) -> tuple[str, str]:
        match = re.search(r" \(\d{4}\)", citation)
        if match:
            return citation[:match.start()], citation[match.start():]
        return citation, ""

    def _get_font_slant(self, name: str) -> int:
        if uno is not None and hasattr(uno, "getConstantByName"):
            try:
                return uno.getConstantByName(f"com.sun.star.awt.FontSlant.{name}")
            except Exception:
                pass
        try:
            from com.sun.star.awt import FontSlant as _FontSlant  # type: ignore

            return getattr(_FontSlant, name)
        except Exception:
            return 0 if name == "NONE" else 2

    def _case_citation_italic_spans(self, text: str) -> list[tuple[int, int]]:
        if not text.strip():
            return []

        word = r"[A-Z][A-Za-z0-9.&'’\-]*"
        lower_word = r"(?:a|an|and|as|at|by|de|del|for|in|la|of|on|or|the|to)"
        entity_word = (
            r"(?:Agency|Association|Bank|Bd\.|Board|Bureau|Cal\.|City|Co\.|Corp\.|"
            r"Corporation|County|Court|Department|Dept\.|District|Division|Inc\.|"
            r"LLC|L\.L\.C\.|People|Regents|Services|State|Superior|University)"
        )
        name_part = rf"(?:{word}|{lower_word}|{entity_word}|ex|rel\.)"
        party = rf"{word}(?:\s+{name_part}){{0,10}}"
        in_re_prefix = (
            r"(?:In re|Estate of|Conservatorship of|Guardianship of|Marriage of|"
            r"Adoption of|Matter of)"
        )
        citation_signal = (
            r"(?=\s*(?:"
            r"\(\d{4}\)|"
            r",\s+supra\b|"
            r"\d{1,4}\s+(?:Cal\.|Cal\.App\.|Cal\.Rptr\.|P\.\d+d|U\.S\.|S\.Ct\.|"
            r"L\.Ed\.|F\.\d+d|F\.Supp\.)"
            r"))"
        )
        patterns = (
            re.compile(
                rf"(?<![\w'’])(?P<name>{party}\s+v\.?\s+{party}){citation_signal}"
            ),
            re.compile(rf"(?<![\w'’])(?P<name>{in_re_prefix}\s+{party}){citation_signal}"),
        )

        spans: list[tuple[int, int]] = []
        for pattern in patterns:
            for match in pattern.finditer(text):
                start, end = match.span("name")
                start, end = self._trim_citation_intro_words(text, start, end)
                if end > start:
                    spans.append((start, end))

        citation_term_pattern = re.compile(r"(?<![\w'’])(?:Id\.|Ibid\.|supra)(?![\w'’])")
        for match in citation_term_pattern.finditer(text):
            spans.append(match.span())

        spans.sort()
        merged: list[tuple[int, int]] = []
        for start, end in spans:
            if not merged or start > merged[-1][1]:
                merged.append((start, end))
                continue
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        return merged

    def _trim_citation_intro_words(self, text: str, start: int, end: int) -> tuple[int, int]:
        candidate = text[start:end]
        intro_pattern = re.compile(
            r"(?i)^(?:(?:see also|but see|pursuant to|based on|relying on|"
            r"according to|the court in|see|cf\.|compare|accord|contra|e\.g\.|"
            r"under|following|applying),?\s+)+"
        )
        match = intro_pattern.match(candidate)
        if not match:
            return start, end
        trimmed_start = start + match.end()
        trimmed = text[trimmed_start:end].lstrip()
        if not trimmed or trimmed.startswith(("v.", "v ")):
            return start, end
        return trimmed_start, end

    def _replace_writer_range_text(  # type: ignore[type-arg]
        self,
        doc: XTextDocument,
        text_range: Any,
        replacement: str,
    ):
        text = self._get_text_container(doc, text_range)
        if not text:
            text_range.setString(replacement)
            return text_range
        try:
            start_anchor = text.createTextCursorByRange(text_range.getStart())
            replace_cursor = text.createTextCursorByRange(text_range.getStart())
            replace_cursor.gotoRange(text_range.getEnd(), True)
            replace_cursor.setString(replacement)
            try:
                if str(replace_cursor.getString() or "") == replacement:
                    return replace_cursor
            except Exception:
                pass
            inserted_cursor = text.createTextCursorByRange(start_anchor)
            if replacement and inserted_cursor.goRight(len(replacement), True):
                return inserted_cursor
            return replace_cursor
        except Exception:
            text_range.setString(replacement)
            return text_range

    def _apply_case_citation_italics_to_range(
        self,
        doc: XTextDocument,  # type: ignore[type-arg]
        text_range: Any,
        replacement: str,
    ) -> None:
        spans = self._case_citation_italic_spans(replacement)
        if not spans:
            return
        text = self._get_text_container(doc, text_range)
        if not text:
            return
        try:
            selected_text = str(text_range.getString() or "")
        except Exception:
            selected_text = ""
        if selected_text != replacement:
            return
        italic = self._get_font_slant("ITALIC")
        for start, end in spans:
            try:
                cursor = text.createTextCursorByRange(text_range.getStart())
                if start and not cursor.goRight(start, False):
                    continue
                if not cursor.goRight(end - start, True):
                    continue
                for prop_name in ("CharPosture", "CharPostureAsian", "CharPostureComplex"):
                    try:
                        cursor.setPropertyValue(prop_name, italic)
                    except Exception:
                        continue
            except Exception:
                continue

    def _apply_case_citation_italics_to_recent_editor_insert(self, inserted_text: str) -> None:
        if not inserted_text or not self._editor_insert_doc or not self._editor_insert_cursor:
            return
        if self._last_insert_len <= 0:
            return
        try:
            text = self._get_text_container(self._editor_insert_doc, self._editor_insert_cursor)
            if not text:
                return
            inserted_range = text.createTextCursorByRange(self._editor_insert_cursor)
            if not inserted_range.goLeft(self._last_insert_len, True):
                return
            self._apply_case_citation_italics_to_range(
                self._editor_insert_doc,
                inserted_range,
                inserted_text,
            )
        except Exception:
            return

    def _apply_case_citation_italics_to_recent_writer_insert(
        self,
        doc: XTextDocument,  # type: ignore[type-arg]
        insert_cursor: Any,
    ) -> None:
        if not doc or not insert_cursor or self._last_insert_len <= 0:
            return
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return
            inserted_range = text.createTextCursorByRange(insert_cursor)
            if not inserted_range.goLeft(self._last_insert_len, True):
                return
            inserted_text = str(inserted_range.getString() or "")
            if not inserted_text.strip():
                return
            self._apply_case_citation_italics_to_range(doc, inserted_range, inserted_text)
        except Exception:
            return

    def _get_autotext_group(self, container, title: str):
        if hasattr(container, "hasByName") and container.hasByName(title):
            return container.getByName(title)
        if hasattr(container, "getElementNames"):
            for name in container.getElementNames():
                try:
                    group = container.getByName(name)
                except Exception:
                    continue
                group_title = getattr(group, "Title", None)
                if not group_title:
                    try:
                        group_title = group.getTitle()
                    except Exception:
                        group_title = None
                if group_title == title:
                    return group
        if hasattr(container, "insertNewByName"):
            return container.insertNewByName(title, title)
        raise RuntimeError("Unable to access AutoText group.")

    def _add_case_to_autotext(
        self,
        citation: str,
        doc: XTextDocument,
        selection_start: Any,
        selection_end: Any,
    ) -> bool:
        if self._ctx is None:
            raise RuntimeError("LibreOffice context unavailable.")
        container = self._ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.text.AutoTextContainer", self._ctx
        )
        group = self._get_autotext_group(container, "Cases")
        display_name = self._autotext_display_name(citation)
        existing_keys = set(group.getElementNames()) if hasattr(group, "getElementNames") else set()
        if hasattr(group, "hasByName") and group.hasByName(display_name):
            return False
        for key in existing_keys:
            try:
                entry = group.getByName(key)
            except Exception:
                continue
            title = getattr(entry, "Title", None)
            if not title:
                try:
                    title = entry.getTitle()
                except Exception:
                    title = None
            if title == display_name:
                return False
        key = self._generate_autotext_key(display_name, existing_keys)
        text = doc.getText()
        range_cursor = text.createTextCursorByRange(selection_start)
        range_cursor.gotoRange(selection_end, True)
        group.insertNewByName(key, display_name, range_cursor)
        if hasattr(group, "store"):
            group.store()
        return True

    def _normalize_curly_quotes(self, text: str) -> str:
        if not text:
            return ""
        normalized = re.sub(r"(?<=\\w)'(?=\\w)", "’", text)
        normalized = self._toggle_ascii_quotes(normalized, '"', "“", "”")
        normalized = self._toggle_ascii_quotes(normalized, "'", "‘", "’")
        return normalized

    def _toggle_ascii_quotes(self, text: str, ascii_char: str, open_char: str, close_char: str) -> str:
        parts: list[str] = []
        open_state = True
        for char in text:
            if char == ascii_char:
                parts.append(open_char if open_state else close_char)
                open_state = not open_state
            else:
                parts.append(char)
        return "".join(parts)

    # UNO helpers --------------------------------------------------------
    def _ensure_writer_frame_active(self, desktop, doc: XTextDocument) -> None:  # type: ignore[type-arg]
        try:
            frame = doc.getCurrentController().getFrame()
            if frame:
                frame.activate()
        except Exception:
            pass
        try:
            desktop.setCurrentComponent(doc)
        except Exception:
            pass

    def _create_page_safe_odt_import_copy(self, source_path: Path) -> Path:
        temp_dir = Path(tempfile.mkdtemp(prefix="prose-odt-import-"))
        output_path = temp_dir / f"{source_path.stem}-page-safe.odt"
        try:
            with zipfile.ZipFile(source_path, "r") as zin:
                infos = zin.infolist()
                if not any(info.filename == "content.xml" for info in infos):
                    raise ValueError("Selected ODT has no content.xml.")
                contents = {info.filename: zin.read(info.filename) for info in infos}
        except zipfile.BadZipFile as exc:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise ValueError("Selected file is not a valid ODT package.") from exc
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise

        for name in ODT_IMPORT_XML_MEMBERS:
            raw = contents.get(name)
            if raw is None:
                continue
            try:
                xml_text = raw.decode("utf-8")
            except UnicodeDecodeError as exc:
                shutil.rmtree(temp_dir, ignore_errors=True)
                raise ValueError(f"Unable to read {name} as UTF-8.") from exc
            stripped, _count = _strip_odt_page_style_metadata(xml_text)
            contents[name] = stripped.encode("utf-8")

        try:
            with zipfile.ZipFile(output_path, "w") as zout:
                def _write_member(info: zipfile.ZipInfo) -> None:
                    zip_info = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
                    zip_info.comment = info.comment
                    zip_info.extra = info.extra
                    zip_info.internal_attr = info.internal_attr
                    zip_info.external_attr = info.external_attr
                    zip_info.create_system = info.create_system
                    zip_info.compress_type = zipfile.ZIP_STORED if info.filename == "mimetype" else info.compress_type
                    zout.writestr(zip_info, contents[info.filename])

                mimetype_info = next((info for info in infos if info.filename == "mimetype"), None)
                if mimetype_info is not None:
                    _write_member(mimetype_info)
                for info in infos:
                    if info.filename == "mimetype":
                        continue
                    _write_member(info)
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise
        return output_path

    def _insert_odt_at_current_cursor(self, doc: XTextDocument, odt_path: Path) -> None:  # type: ignore[type-arg]
        if uno is None:
            raise ValueError("python-uno unavailable. Set the LibreOffice Python path in Settings.")
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            text = self._get_text_container(doc, view_cursor)
            if not text:
                raise ValueError("Unable to find the current Writer text cursor.")
            try:
                if view_cursor.getString():
                    view_cursor.setString("")
            except Exception:
                pass
            insert_cursor = text.createTextCursorByRange(view_cursor)
            insert_url = uno.systemPathToFileUrl(str(odt_path))
            insertable = getattr(insert_cursor, "insertDocumentFromURL", None)
            if not callable(insertable):
                insertable = getattr(view_cursor, "insertDocumentFromURL", None)
            if not callable(insertable):
                raise ValueError("LibreOffice cursor does not support document insertion.")
            insertable(insert_url, ())
        except Exception as exc:  # noqa: BLE001
            if isinstance(exc, ValueError):
                raise
            raise ValueError(f"Unable to insert sanitized ODT: {exc}") from exc

    def _extract_section_between_headings(
        self, doc: XTextDocument, start_heading: str, end_heading: str
    ) -> str | None:  # type: ignore[type-arg]
        try:
            doc_text = doc.getText()
        except Exception:
            return None
        section = None
        try:
            enum = doc_text.createEnumeration()
            collecting = False
            parts: list[str] = []
            while enum.hasMoreElements():
                element = enum.nextElement()
                if not hasattr(element, "getString"):
                    continue
                paragraph = element.getString()
                paragraph_stripped = paragraph.strip()
                if not collecting:
                    if paragraph_stripped == start_heading:
                        collecting = True
                    continue
                if paragraph_stripped == end_heading:
                    break
                parts.append(paragraph)
            if collecting and parts:
                section = "\n".join(part.strip() for part in parts).strip() or None
        except Exception:
            section = None
        if section:
            return section
        try:
            text = doc_text.getString()
        except Exception:
            return None
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        start_re = re.compile(rf"(?mi)^\\s*{re.escape(start_heading)}\\s*$")
        end_re = re.compile(rf"(?mi)^\\s*{re.escape(end_heading)}\\s*$")
        start_match = start_re.search(text)
        if not start_match:
            return None
        end_match = end_re.search(text, start_match.end())
        if not end_match:
            return None
        section = text[start_match.end() : end_match.start()].strip()
        return section or None

    def _extract_paragraph_at_cursor(self, doc: XTextDocument) -> str | None:  # type: ignore[type-arg]
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            text = doc.getText()
            range_cursor = text.createTextCursorByRange(view_cursor)
            range_cursor.gotoStartOfParagraph(False)
            range_cursor.gotoEndOfParagraph(True)
            paragraph = range_cursor.getString()
        except Exception:
            return None
        paragraph = paragraph.strip()
        return paragraph or None

    def _is_heading_paragraph(self, paragraph: Any) -> bool:  # type: ignore[type-arg]
        try:
            outline_level = paragraph.getPropertyValue("OutlineLevel")
            if isinstance(outline_level, int) and outline_level > 0:
                return True
        except Exception:
            pass
        try:
            style = paragraph.getPropertyValue("ParaStyleName")
            if isinstance(style, str) and style.lower().startswith("heading"):
                return True
        except Exception:
            pass
        return False

    def _is_heading_cursor(self, cursor: Any) -> bool:  # type: ignore[type-arg]
        try:
            outline_level = cursor.getPropertyValue("OutlineLevel")
            if isinstance(outline_level, int) and outline_level > 0:
                return True
        except Exception:
            pass
        try:
            style = cursor.getPropertyValue("ParaStyleName")
            if isinstance(style, str) and style.lower().startswith("heading"):
                return True
        except Exception:
            pass
        return False

    def _extract_section_to_last_heading(self, doc: XTextDocument, view_cursor: Any) -> str | None:  # type: ignore[type-arg]
        try:
            text = doc.getText()
            cursor_start = view_cursor.getStart()
            scan_cursor = text.createTextCursorByRange(cursor_start)
        except Exception:
            return None
        try:
            scan_cursor.gotoStartOfParagraph(False)
        except Exception:
            return None
        heading_end = None
        while True:
            if self._is_heading_cursor(scan_cursor):
                heading_end = scan_cursor.getEnd()
                break
            if not scan_cursor.goLeft(1, False):
                break
            try:
                scan_cursor.gotoStartOfParagraph(False)
            except Exception:
                break
        if not heading_end:
            return None
        try:
            section_cursor = text.createTextCursorByRange(heading_end)
            section_cursor.gotoRange(cursor_start, True)
            section = section_cursor.getString()
        except Exception:
            return None
        section = section.strip()
        return section or None

    def _get_desktop(self):
        if uno is None:
            self._status_label.set_label(_uno_status_message(self._libreoffice_python_path))
            return None
        if self._desktop is not None:
            return self._desktop
        connected = self._connect_desktop(retries=1, delay=0)
        if not connected:
            self._start_listener()
            connected = self._connect_desktop(retries=8, delay=0.25)
        if not connected:
            self._status_label.set_label("LibreOffice listener unreachable.")
            self._desktop = None
        return self._desktop

    def _get_active_writer(self, desktop) -> XTextDocument | None:  # type: ignore[type-arg]
        try:
            comp = desktop.getCurrentComponent()
            if comp and comp.supportsService("com.sun.star.text.TextDocument"):
                self._active_doc = comp
                return comp
        except Exception:
            pass
        # Fallback to the last loaded document if UNO does not return a current component (headless mode)
        if self._active_doc and getattr(self._active_doc, "supportsService", lambda *_: False)(
            "com.sun.star.text.TextDocument"
        ):
            return self._active_doc
        return None

    def _find_first_uncombined_cites(self, doc: XTextDocument) -> tuple[str, Any] | None:  # type: ignore[type-arg]
        pattern = "[(][^)]*[)]([[:space:]]+[(][^)]*[)])+"
        pattern_re = re.compile(r"\([^)]*\)(\s+\([^)]*\))+")
        try:
            doc_text = doc.getText()
        except Exception:
            doc_text = None
        found = self._find_first_uncombined_cites_in_searchable(doc, doc_text, pattern)
        if not found:
            for note_text in self._iter_note_texts(doc):
                found = self._find_first_uncombined_cites_in_text(note_text, pattern_re)
                if found:
                    break
        if not found:
            return None
        try:
            cite_text, cite_cursor = found
            doc.getCurrentController().select(cite_cursor)
            return cite_text, cite_cursor
        except Exception:
            return None

    def _find_first_uncombined_cites_in_searchable(self, searchable: Any, text: Any, pattern: str) -> tuple[str, Any] | None:
        try:
            if not searchable or not text:
                return None
            search_desc = searchable.createSearchDescriptor()
            search_desc.SearchString = pattern
            search_desc.SearchRegularExpression = True
            found = searchable.findFirst(search_desc)
            if not found:
                return None
            cite_cursor = text.createTextCursorByRange(found)
            cite_text = cite_cursor.getString()
            if not cite_text.strip():
                return None
            return cite_text, cite_cursor
        except Exception:
            return None

    def _find_first_uncombined_cites_in_text(self, text: Any, pattern: re.Pattern[str]) -> tuple[str, Any] | None:
        try:
            raw = text.getString()
            if not raw:
                return None
            match = pattern.search(raw)
            if not match:
                return None
            start, end = match.span()
            cursor = text.createTextCursor()
            if not cursor.goRight(start, False):
                return None
            if not cursor.goRight(end - start, True):
                return None
            cite_text = cursor.getString()
            if not cite_text.strip():
                return None
            return cite_text, cursor
        except Exception:
            return None

    def _iter_note_texts(self, doc: XTextDocument):  # type: ignore[type-arg]
        for attr_name in ("getFootnotes", "getEndnotes"):
            try:
                notes = getattr(doc, attr_name)()
            except Exception:
                continue
            try:
                count = notes.getCount()
            except Exception:
                continue
            for idx in range(count):
                try:
                    note = notes.getByIndex(idx)
                    text = note.getText()
                except Exception:
                    continue
                if text is not None:
                    yield text

    def _run_combine_cites(self, cite_text: str, profile: ModelProfile) -> None:
        payload = self._compose_combine_cites_payload(cite_text, profile)
        try:
            combined_text = self._call_combine_cites(payload, profile)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_combine_cites_failed, str(exc))
            return
        GLib.idle_add(self._apply_combined_cites, combined_text)

    def _compose_combine_cites_payload(self, cite_text: str, profile: ModelProfile) -> dict[str, Any]:
        prompt = _expand_shared_prompt_parts(self._combine_cites_settings.prompt or DEFAULT_COMBINE_CITES_PROMPT)
        content = f"{prompt}\n\n{cite_text}" if prompt else cite_text
        payload = {
            "messages": [
                {"role": "user", "content": content},
            ],
            "stream": False,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _call_combine_cites(self, payload: dict[str, Any], profile: ModelProfile) -> str:
        raw = self._post_json_and_read(
            payload,
            profile.api_url,
            profile.api_key,
            request_title="Combine Cites",
        )
        parts = list(self._extract_response_text(raw))
        combined_text = "".join(parts).strip()
        if not combined_text:
            raise ValueError("Combine Cites returned empty output.")
        return combined_text

    def _apply_combined_cites(self, combined_text: str) -> bool:
        try:
            if not self._combine_cites_doc or not self._combine_cites_cursor:
                self._status_label.set_label("Unable to replace citations.")
                return False
            self._combine_cites_cursor.setString(combined_text)
            text = self._get_text_container(self._combine_cites_doc, self._combine_cites_cursor)
            if not text:
                self._status_label.set_label("Unable to replace citations.")
                return False
            self._apply_case_citation_italics_to_range(
                self._combine_cites_doc,
                self._combine_cites_cursor,
                combined_text,
            )
            insert_cursor = text.createTextCursorByRange(self._combine_cites_cursor.getEnd())
            previous_insert_len = self._last_insert_len
            self._last_insert_len = len(combined_text)
            self._ensure_single_trailing_space(self._combine_cites_doc, insert_cursor)
            self._last_insert_len = previous_insert_len
            try:
                view_cursor = self._combine_cites_doc.getCurrentController().getViewCursor()
                move_cursor = text.createTextCursorByRange(insert_cursor)
                if self._get_following_char_at_cursor(self._combine_cites_doc, insert_cursor) == " ":
                    move_cursor.goRight(1, False)
                view_cursor.gotoRange(move_cursor, False)
            except Exception:
                pass
            self._status_label.set_label("Citations combined.")
            return False
        except Exception as exc:  # noqa: BLE001
            self._status_label.set_label(f"Unable to replace citations: {exc}")
            return False
        finally:
            self._set_busy(False)
            self._combine_cites_doc = None
            self._combine_cites_cursor = None

    def _on_combine_cites_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._notify_llm_error(message)
        return False

    def _on_action_suggestion_goto(self, _action: Gio.SimpleAction, param: GLib.Variant) -> None:
        if param is None:
            return
        self._on_goto_clicked(None, int(param.get_int32()))

    def _on_action_suggestion_accept(self, _action: Gio.SimpleAction, param: GLib.Variant) -> None:
        if param is None:
            return
        self._on_accept_clicked(None, int(param.get_int32()))

    def _on_action_suggestion_reject(self, _action: Gio.SimpleAction, param: GLib.Variant) -> None:
        if param is None:
            return
        self._on_reject_clicked(None, int(param.get_int32()))

    def _on_action_save_settings(self) -> None:
        if not self._settings_window:
            self._show_toast("Settings window is not open.")
            return
        self._settings_window.trigger_save()

    def _get_preceding_char(self, doc: XTextDocument) -> str:  # type: ignore[type-arg]
        try:
            view_cursor = doc.getCurrentController().getViewCursor()
            text = self._get_text_container(doc, view_cursor)
            if not text:
                return ""
            range_cursor = text.createTextCursorByRange(view_cursor)
            if not range_cursor.goLeft(1, True):
                return ""
            return range_cursor.getString() or ""
        except Exception:
            return ""

    def _get_following_char_at_cursor(self, doc: XTextDocument, insert_cursor: Any) -> str:  # type: ignore[type-arg]
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return ""
            range_cursor = text.createTextCursorByRange(insert_cursor)
            if not range_cursor.goRight(1, True):
                return ""
            return range_cursor.getString() or ""
        except Exception:
            return ""

    def _get_insert_end_cursor(self, doc: XTextDocument, insert_cursor: Any) -> Any | None:  # type: ignore[type-arg]
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return None
            end_range = insert_cursor.getEnd() if hasattr(insert_cursor, "getEnd") else insert_cursor
            return text.createTextCursorByRange(end_range)
        except Exception:
            return None

    def _move_view_cursor_to_range(self, doc: XTextDocument, text_range: Any) -> Any | None:  # type: ignore[type-arg]
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            view_cursor.gotoRange(text_range, False)
            return view_cursor
        except Exception:
            return None

    def _insert_space_at_cursor(self, doc: XTextDocument, insert_cursor: Any) -> bool:  # type: ignore[type-arg]
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return False
            end_cursor = self._get_insert_end_cursor(doc, insert_cursor)
            if end_cursor:
                view_cursor = self._move_view_cursor_to_range(doc, end_cursor)
                if view_cursor:
                    text.insertString(view_cursor, " ", False)
                    return True
                text.insertString(end_cursor, " ", False)
                return True
            view_cursor = doc.getCurrentController().getViewCursor()
            text.insertString(view_cursor, " ", False)
            return True
        except Exception:
            return False

    def _launch_writer_document(self, path: Path) -> None:
        desktop = self._get_desktop()
        if not desktop:
            if uno is None:
                self._show_toast("python-uno unavailable. Set the LibreOffice Python path in Settings.")
            else:
                self._show_toast("LibreOffice listener unavailable; cannot open document.")
            return
        try:
            url = uno.systemPathToFileUrl(str(path))

            def _prop(name: str, value: Any) -> PropertyValue:
                pv = PropertyValue()
                pv.Name = name
                pv.Value = value
                return pv

            args = (
                _prop("Hidden", False),
                _prop("ReadOnly", False),
                _prop("LockContentExtraction", False),
                _prop("AsTemplate", False),
            )
            doc = desktop.loadComponentFromURL(url, "_blank", 0, args)
            if not doc:
                self._show_toast("Document did not open via UNO.")
                return
            self._active_doc = doc  # Remember for headless sessions with no visible window
            self._maybe_save_last_odt(path)
            try:
                readonly_attr = getattr(doc, "isReadonly", False)
                readonly = readonly_attr() if callable(readonly_attr) else bool(readonly_attr)
                if readonly:
                    self._show_toast("Document opened read-only. Close other LibreOffice instances or clear lock file.")
            except Exception:
                pass
            self._status_label.set_label(f"Opened {path.name} in UNO session.")
            self._show_toast(f"Loaded {path.name} via listener.")
        except Exception as exc:  # noqa: BLE001
            self._show_toast(f"Open failed: {exc}")

    def _maybe_save_last_odt(self, path: Path) -> None:
        if path.suffix.lower() != ".odt":
            return
        self._set_last_odt_path(path)

    def _set_last_odt_path(self, path: Path | None) -> None:
        self._last_odt_path = path.expanduser().resolve(strict=False) if path else None
        save_last_odt_file(self._last_odt_path)
        if hasattr(self, "_open_last_btn") and self._open_last_btn:
            self._open_last_btn.set_sensitive(bool(self._last_odt_path))

    def _collect_proofreading_paragraphs(self, doc: XTextDocument) -> list[ProofreadParagraph]:  # type: ignore[type-arg]
        try:
            doc_text = doc.getText()
            enum = doc_text.createEnumeration()
        except Exception:
            try:
                fallback = str(doc.getText().getString() or "").strip()
            except Exception:
                fallback = ""
            return [ProofreadParagraph(text=fallback, page=None)] if fallback else []

        controller = None
        view_cursor = None
        original_range = None
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            original_range = view_cursor.getStart()
        except Exception:
            view_cursor = None

        paragraphs: list[ProofreadParagraph] = []
        try:
            while enum.hasMoreElements():
                paragraph = enum.nextElement()
                try:
                    paragraph_text = str(paragraph.getString() or "")
                except Exception:
                    paragraph_text = ""
                cleaned = paragraph_text.strip()
                if not cleaned:
                    continue
                page_number: int | None = None
                if view_cursor is not None:
                    try:
                        view_cursor.gotoRange(paragraph.getStart(), False)
                        page_getter = getattr(view_cursor, "getPage", None)
                        if callable(page_getter):
                            page_value = page_getter()
                            if isinstance(page_value, int) and page_value > 0:
                                page_number = page_value
                    except Exception:
                        page_number = None
                paragraphs.append(ProofreadParagraph(text=cleaned, page=page_number))
        finally:
            if view_cursor is not None and original_range is not None:
                try:
                    view_cursor.gotoRange(original_range, False)
                except Exception:
                    pass
        return paragraphs

    def _find_range(self, doc: XTextDocument, snippet: str):  # type: ignore[type-arg]
        def _loose_literal_pattern(text: str) -> str:
            punctuation_variants = {
                '"': '"“”',
                "“": '"“”',
                "”": '"“”',
                "'": "'‘’",
                "‘": "'‘’",
                "’": "'‘’",
                "-": "-–—",
                "–": "-–—",
                "—": "-–—",
            }
            parts: list[str] = []
            last_was_space = False
            for char in text.strip():
                if char.isspace() or char == "\u00a0":
                    if not last_was_space:
                        parts.append(r"[ \t\r\n\u00a0]+")
                        last_was_space = True
                    continue
                last_was_space = False
                variants = punctuation_variants.get(char)
                if variants is not None:
                    parts.append(f"[{re.escape(variants)}]")
                else:
                    parts.append(re.escape(char))
            return "".join(parts)

        try:
            search = doc.createSearchDescriptor()
            search.SearchString = snippet
            search.SearchRegularExpression = False
            search.SearchCaseSensitive = False  # type: ignore[attr-defined]
            found = doc.findFirst(search)
            if found:
                return found

            regex_pattern = _loose_literal_pattern(snippet)
            if not regex_pattern:
                return None
            search = doc.createSearchDescriptor()
            search.SearchString = regex_pattern
            search.SearchRegularExpression = True
            search.SearchCaseSensitive = False  # type: ignore[attr-defined]
            search.SearchWords = False  # type: ignore[attr-defined]
            return doc.findFirst(search)
        except Exception:
            return None

    def _find_speech_query_range(self, doc: XTextDocument, query: str):  # type: ignore[type-arg]
        words = query.split()
        if not words:
            return None
        separator = r"[[:space:][:punct:]“”‘’–—]+"
        pattern = separator.join(re.escape(word) for word in words)
        try:
            search = doc.createSearchDescriptor()
            search.SearchString = pattern
            search.SearchRegularExpression = True
            search.SearchCaseSensitive = False  # type: ignore[attr-defined]
            search.SearchWords = False  # type: ignore[attr-defined]

            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            start_range = view_cursor.getEnd()
            try:
                found = doc.findNext(start_range, search)
            except Exception:
                found = None
            if found:
                return found
            return doc.findFirst(search)
        except Exception:
            return None

    def _get_text_container(self, doc: XTextDocument, text_range: Any | None):  # type: ignore[type-arg]
        if text_range is not None:
            try:
                text = text_range.getText()
                if text is not None:
                    return text
            except Exception:
                pass
        try:
            return doc.getText()
        except Exception:
            return None

    def _prepare_editor_insertion(self, doc: XTextDocument) -> bool:  # type: ignore[type-arg]
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            try:
                if view_cursor.getString():
                    view_cursor.setString("")
            except Exception:
                pass
            text = self._get_text_container(doc, view_cursor)
            if not text:
                return False
            self._editor_insert_cursor = text.createTextCursorByRange(view_cursor)
            self._editor_insert_end = None
            self._editor_insert_doc = doc
            self._last_insert_len = 0
            self._editor_pending_newlines = 0
            return True
        except Exception:
            return False

    def _append_editor_text(self, text: str) -> bool:
        if not text or not self._editor_insert_cursor or not self._editor_insert_doc:
            return False
        try:
            self._append_writer_text(
                text,
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
        except Exception as exc:  # noqa: BLE001
            self._show_toast(f"Unable to insert text: {exc}")
        return False

    def _prepare_improve_insertion(self, doc: XTextDocument) -> bool:  # type: ignore[type-arg]
        if not self._editor_insert_end or self._last_insert_len <= 0:
            return False
        try:
            text = self._get_text_container(doc, self._editor_insert_end)
            if not text:
                return False
            range_cursor = text.createTextCursorByRange(self._editor_insert_end)
            if not range_cursor.goLeft(self._last_insert_len, True):
                return False
            try:
                range_cursor.setString("")
            except Exception:
                pass
            self._improve_insert_cursor = text.createTextCursorByRange(range_cursor)
            self._improve_insert_doc = doc
            self._editor_insert_end = None
            self._last_insert_len = 0
            self._improve_pending_newlines = 0
            self._improve_started = False
            return True
        except Exception:
            return False

    def _prepare_selection_insertion(self, doc: XTextDocument, view_cursor: Any) -> bool:  # type: ignore[type-arg]
        try:
            text = self._get_text_container(doc, view_cursor)
            if not text:
                return False
            selection_start = view_cursor.getStart()
            selection_end = view_cursor.getEnd()
            range_cursor = text.createTextCursorByRange(selection_start)
            range_cursor.gotoRange(selection_end, True)
            try:
                range_cursor.setString("")
            except Exception:
                pass
            self._improve_insert_cursor = text.createTextCursorByRange(range_cursor)
            self._improve_insert_doc = doc
            self._editor_insert_end = None
            self._last_insert_len = 0
            self._improve_pending_newlines = 0
            self._improve_started = False
            return True
        except Exception:
            return False

    def _append_improve1_text(self, text: str) -> bool:
        if not text or not self._improve_insert_cursor or not self._improve_insert_doc:
            return False
        if not self._improve_started:
            text = text.lstrip()
            if not text:
                return False
        try:
            self._append_writer_text(
                text,
                self._improve_insert_doc,
                self._improve_insert_cursor,
                "_improve_pending_newlines",
            )
            self._improve_started = True
        except Exception as exc:  # noqa: BLE001
            self._show_toast(f"Unable to insert improved text: {exc}")
        return False

    def _trim_trailing_whitespace(self, doc: XTextDocument, insert_cursor: Any) -> None:  # type: ignore[type-arg]
        if not doc or not insert_cursor or self._last_insert_len <= 0:
            return
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return
            range_cursor = text.createTextCursorByRange(insert_cursor)
            if not range_cursor.goLeft(self._last_insert_len, True):
                return
            current = range_cursor.getString()
            if not current:
                return
            trailing_len = len(current) - len(current.rstrip())
            if trailing_len:
                trim_cursor = text.createTextCursorByRange(insert_cursor)
                if trim_cursor.goLeft(trailing_len, True):
                    trim_cursor.setString("")
                    self._last_insert_len -= trailing_len
        except Exception:
            return

    def _ensure_single_trailing_space(self, doc: XTextDocument, insert_cursor: Any) -> None:  # type: ignore[type-arg]
        if not doc or not insert_cursor:
            return
        if self._last_insert_len > 0:
            self._trim_trailing_whitespace(doc, insert_cursor)
        try:
            end_cursor = self._get_insert_end_cursor(doc, insert_cursor)
            if end_cursor and self._get_following_char_at_cursor(doc, end_cursor) == " ":
                return
            inserted = self._insert_space_at_cursor(doc, insert_cursor)
            end_cursor = self._get_insert_end_cursor(doc, insert_cursor)
            following = self._get_following_char_at_cursor(doc, end_cursor) if end_cursor else ""
            if self._last_insert_len > 0:
                self._last_insert_len += 1
        except Exception:
            return

    def _normalize_newlines(self, text: str) -> str:
        return text.replace("\r\n", "\n").replace("\r", "\n")

    def _normalize_generated_output_chunk(self, text: str, *, is_first_chunk: bool) -> str:
        normalized = self._normalize_newlines(text)
        if is_first_chunk:
            normalized = normalized.lstrip()
        return normalized

    def _normalize_generated_output_text(self, text: str) -> str:
        return self._normalize_newlines(text).strip()

    def _trim_spelling_output_edges(self) -> None:
        current = self._get_spelling_output_text()
        cleaned = self._normalize_generated_output_text(current)
        if cleaned != current:
            self._set_spelling_output_text(cleaned)

    def _trim_reference_output_edges(self) -> None:
        current = self._reference_output_text or ""
        cleaned = self._normalize_generated_output_text(current)
        if cleaned != current:
            self._set_reference_output_text(cleaned)

    def _split_text_for_paragraphs(self, text: str, pending_newlines: int) -> tuple[list[tuple[str, str]], int]:
        normalized = self._normalize_newlines(text)
        if pending_newlines:
            normalized = ("\n" * pending_newlines) + normalized
        parts: list[tuple[str, str]] = []
        buffer: list[str] = []
        idx = 0
        length = len(normalized)
        while idx < length:
            char = normalized[idx]
            if char != "\n":
                buffer.append(char)
                idx += 1
                continue
            run = 1
            idx += 1
            while idx < length and normalized[idx] == "\n":
                run += 1
                idx += 1
            if run >= 2:
                if buffer:
                    parts.append(("text", "".join(buffer)))
                    buffer = []
                breaks = run // 2
                for _ in range(breaks):
                    parts.append(("paragraph", ""))
                if run % 2:
                    buffer.append("\n")
            else:
                buffer.append("\n")
        if buffer:
            parts.append(("text", "".join(buffer)))
        new_pending = 0
        if parts and parts[-1][0] == "text" and parts[-1][1].endswith("\n"):
            text_part = parts[-1][1][:-1]
            if text_part:
                parts[-1] = ("text", text_part)
            else:
                parts.pop()
            new_pending = 1
        return parts, new_pending

    def _insert_paragraph_break(self, doc: XTextDocument, insert_cursor: Any) -> None:  # type: ignore[type-arg]
        try:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return
            if ControlCharacter is None:
                raise ValueError("UNO ControlCharacter unavailable")
            text.insertControlCharacter(insert_cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        except Exception:
            text = self._get_text_container(doc, insert_cursor)
            if not text:
                return
            text.insertString(insert_cursor, "\n", False)

    def _append_writer_text(
        self,
        text: str,
        doc: XTextDocument,  # type: ignore[type-arg]
        insert_cursor: Any,
        pending_attr: str,
    ) -> None:
        pending_newlines = int(getattr(self, pending_attr, 0) or 0)
        parts, pending_newlines = self._split_text_for_paragraphs(text, pending_newlines)
        for kind, chunk in parts:
            if kind == "paragraph":
                self._insert_paragraph_break(doc, insert_cursor)
                self._last_insert_len += 1
                continue
            if not chunk:
                continue
            text_obj = self._get_text_container(doc, insert_cursor)
            if not text_obj:
                return
            text_obj.insertString(insert_cursor, chunk, False)
            self._last_insert_len += len(chunk)
        setattr(self, pending_attr, pending_newlines)

    def _flush_pending_newlines(
        self,
        doc: XTextDocument,  # type: ignore[type-arg]
        insert_cursor: Any,
        pending_attr: str,
    ) -> None:
        pending_newlines = int(getattr(self, pending_attr, 0) or 0)
        if pending_newlines <= 0:
            return
        try:
            breaks = pending_newlines // 2
            for _ in range(breaks):
                self._insert_paragraph_break(doc, insert_cursor)
                self._last_insert_len += 1
            if pending_newlines % 2:
                text_obj = self._get_text_container(doc, insert_cursor)
                if not text_obj:
                    return
                text_obj.insertString(insert_cursor, "\n", False)
                self._last_insert_len += 1
        except Exception:
            return
        setattr(self, pending_attr, 0)

    def _set_spelling_output_text(self, text: str) -> bool:
        if self._editor_output_surface is not None:
            self._editor_output_surface.set_visible_child_name("output")
        if self._spelling_output_buffer is None:
            return False
        self._spelling_output_buffer.set_text(text or "")
        return False

    def _append_spelling_output_text(self, text: str) -> bool:
        if self._editor_output_surface is not None:
            self._editor_output_surface.set_visible_child_name("output")
        if not text or self._spelling_output_buffer is None:
            return False
        end_iter = self._spelling_output_buffer.get_end_iter()
        self._spelling_output_buffer.insert(end_iter, text)
        return False

    def _get_spelling_output_text(self) -> str:
        if self._spelling_output_buffer is None:
            return ""
        start, end = self._spelling_output_buffer.get_bounds()
        return self._spelling_output_buffer.get_text(start, end, True)

    def _mark_text_draft_temp_dirty(self) -> None:
        self._text_draft_temp_dirty = True
        if self._text_draft_temp_flush_source_id:
            return
        self._text_draft_temp_flush_source_id = GLib.timeout_add(
            TEXT_DRAFT_TEMP_FLUSH_DELAY_MS,
            self._flush_text_draft_temp_file,
        )

    def _on_text_draft_buffer_changed(self, _buffer: Gtk.TextBuffer) -> None:
        self._mark_text_draft_temp_dirty()
        self._refresh_text_draft_clear_button()

    def _on_text_draft_surface_visible_child_name_changed(self, _stack: Gtk.Stack, _pspec: object) -> None:
        self._refresh_text_draft_clear_button()

    def _on_text_draft_key_pressed(
        self,
        _controller: Gtk.EventControllerKey,
        keyval: int,
        _keycode: int,
        state: Gdk.ModifierType,
    ) -> bool:
        return False

    def _hide_text_draft_case_suggestions(self) -> None:
        if self._text_draft_case_results_scroller is not None:
            self._text_draft_case_results_scroller.set_visible(False)
        self._text_draft_case_suggestions = []
        self._text_draft_case_selected_index = 0
        self._text_draft_case_prefix = ""
        self._text_draft_case_prefix_bounds = None

    def _normalize_text_draft_case_lookup_text(self, text: str) -> str:
        return re.sub(r"\s+", " ", text).strip().lower()

    def _compact_text_draft_case_lookup_text(self, text: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", text.lower())

    def _text_draft_case_search_terms(self, display_name: str) -> tuple[str, ...]:
        candidates = [display_name]
        seen: set[str] = set()
        terms: list[str] = []
        for candidate in candidates:
            for normalized in (
                self._normalize_text_draft_case_lookup_text(candidate),
                self._compact_text_draft_case_lookup_text(candidate),
            ):
                if not normalized or normalized in seen:
                    continue
                seen.add(normalized)
                terms.append(normalized)
        return tuple(terms)

    def _is_text_draft_slip_opinion_case(self, term: str, citation: str) -> bool:
        return any("slip opn" in value.lower() for value in (term, citation))

    def _load_text_draft_case_index(self, *, force: bool = False) -> list[TextDraftCaseSuggestion]:
        concordance_file = self._concordance_file_path
        if concordance_file is None:
            self._text_draft_case_index = []
            self._text_draft_case_index_path = None
            self._text_draft_case_index_mtime_ns = None
            return []
        try:
            stat = concordance_file.stat()
        except OSError:
            self._text_draft_case_index = []
            self._text_draft_case_index_path = concordance_file
            self._text_draft_case_index_mtime_ns = None
            return []
        if (
            not force
            and self._text_draft_case_index_path == concordance_file
            and self._text_draft_case_index_mtime_ns == stat.st_mtime_ns
        ):
            return self._text_draft_case_index

        suggestions: list[TextDraftCaseSuggestion] = []
        seen_citations: set[str] = set()
        try:
            lines = concordance_file.read_text(encoding="utf-8", errors="ignore").splitlines()
        except OSError:
            lines = []
        for raw_line in lines:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [part.strip() for part in line.split(";")]
            if len(parts) < 3 or parts[2] != "Cases":
                continue
            term = parts[0]
            citation = parts[1] or term
            if self._is_text_draft_slip_opinion_case(term, citation):
                continue
            if not citation or citation in seen_citations:
                continue
            display_name = self._autotext_display_name(citation)
            search_terms = self._text_draft_case_search_terms(display_name)
            if not search_terms:
                continue
            seen_citations.add(citation)
            suggestions.append(
                TextDraftCaseSuggestion(
                    citation=citation,
                    display_name=display_name,
                    search_terms=search_terms,
                )
            )
        suggestions.sort(key=lambda suggestion: suggestion.display_name.lower())
        self._text_draft_case_index = suggestions
        self._text_draft_case_index_path = concordance_file
        self._text_draft_case_index_mtime_ns = stat.st_mtime_ns
        return suggestions

    def _matching_text_draft_case_suggestions(
        self,
        prefix: str,
        *,
        limit: int | None = TEXT_DRAFT_CASE_SUGGESTION_LIMIT,
    ) -> tuple[list[TextDraftCaseSuggestion], int]:
        normalized_prefix = self._normalize_text_draft_case_lookup_text(prefix)
        compact_prefix = self._compact_text_draft_case_lookup_text(prefix)
        if not normalized_prefix and not compact_prefix:
            return ([], 0)
        prefixes = tuple(
            candidate
            for candidate in (normalized_prefix, compact_prefix)
            if candidate
        )
        ranked_matches: list[tuple[int, TextDraftCaseSuggestion]] = []
        for suggestion in self._load_text_draft_case_index():
            rank: int | None = None
            for term in suggestion.search_terms:
                if any(
                    term.startswith(prefix_candidate)
                    for prefix_candidate in prefixes
                ):
                    rank = 0
                    break
                if any(prefix_candidate in term for prefix_candidate in prefixes):
                    rank = 1
            if rank is not None:
                ranked_matches.append((rank, suggestion))
        ranked_matches.sort(
            key=lambda match: (match[0], match[1].display_name.lower())
        )
        matches = [suggestion for _rank, suggestion in ranked_matches]
        total = len(matches)
        if limit is not None:
            matches = matches[:limit]
        return (matches, total)

    def _refresh_text_draft_case_suggestions(self) -> None:
        entry = self._text_draft_case_search_entry
        if entry is None:
            return
        entry_had_focus = entry.has_focus()
        prefix = entry.get_text().strip()
        suggestions, _total = self._matching_text_draft_case_suggestions(prefix)
        if not suggestions:
            self._hide_text_draft_case_suggestions()
            if entry_had_focus:
                entry.grab_focus()
            return
        if prefix != self._text_draft_case_prefix:
            self._text_draft_case_selected_index = 0
            self._text_draft_case_prefix = prefix
        self._text_draft_case_suggestions = suggestions
        self._text_draft_case_prefix_bounds = None
        self._text_draft_case_selected_index = min(
            self._text_draft_case_selected_index,
            len(suggestions) - 1,
        )
        self._show_text_draft_case_suggestions()
        if entry_had_focus:
            entry.grab_focus()

    def _show_text_draft_case_suggestions(self) -> None:
        scroller = self._text_draft_case_results_scroller
        list_box = self._text_draft_case_list_box
        if scroller is None or list_box is None:
            return
        self._clear_list_box(list_box)
        for suggestion in self._text_draft_case_suggestions:
            row = Gtk.ListBoxRow()
            row.set_selectable(True)
            row.set_activatable(True)
            row.add_css_class("text-draft-case-suggestion-row")
            label = Gtk.Label(label=suggestion.citation, xalign=0)
            label.set_ellipsize(Pango.EllipsizeMode.END)
            label.set_tooltip_text(suggestion.citation)
            row.set_child(label)
            list_box.append(row)
        selected_row = list_box.get_row_at_index(self._text_draft_case_selected_index)
        if selected_row is not None:
            list_box.select_row(selected_row)
        scroller.set_visible(True)

    def _move_text_draft_case_selection(self, direction: int) -> None:
        if not self._text_draft_case_suggestions:
            return
        self._text_draft_case_selected_index = (
            self._text_draft_case_selected_index + direction
        ) % len(self._text_draft_case_suggestions)
        list_box = self._text_draft_case_list_box
        if list_box is not None:
            row = list_box.get_row_at_index(self._text_draft_case_selected_index)
            if row is not None:
                list_box.select_row(row)

    def _on_text_draft_case_row_activated(self, _list_box: Gtk.ListBox, row: Gtk.ListBoxRow) -> None:
        row_index = row.get_index()
        if 0 <= row_index < len(self._text_draft_case_suggestions):
            self._text_draft_case_selected_index = row_index
        self._attach_selected_text_draft_case_suggestion()

    def _on_text_draft_case_search_changed(self, _entry: Gtk.Entry) -> None:
        self._refresh_text_draft_case_suggestions()

    def _on_text_draft_case_search_key_pressed(
        self,
        _controller: Gtk.EventControllerKey,
        keyval: int,
        _keycode: int,
        state: Gdk.ModifierType,
    ) -> bool:
        if keyval in (Gdk.KEY_Escape,):
            self._hide_text_draft_case_suggestions()
            return True
        if (
            self._text_draft_case_results_scroller is None
            or not self._text_draft_case_suggestions
            or not self._text_draft_case_results_scroller.get_visible()
        ):
            return False
        if keyval in (Gdk.KEY_Return, Gdk.KEY_KP_Enter):
            self._attach_selected_text_draft_case_suggestion()
            return True
        if keyval in (Gdk.KEY_Down, Gdk.KEY_Tab, Gdk.KEY_ISO_Left_Tab):
            direction = -1 if state & Gdk.ModifierType.SHIFT_MASK else 1
            if keyval == Gdk.KEY_ISO_Left_Tab:
                direction = -1
            self._move_text_draft_case_selection(direction)
            return True
        if keyval == Gdk.KEY_Up:
            self._move_text_draft_case_selection(-1)
            return True
        return False

    def _attach_selected_text_draft_case_suggestion(self) -> None:
        if not self._text_draft_case_suggestions:
            self._hide_text_draft_case_suggestions()
            return
        suggestion = self._text_draft_case_suggestions[self._text_draft_case_selected_index]
        if suggestion.citation not in self._text_draft_case_attachments:
            self._text_draft_case_attachments.append(suggestion.citation)
            self._render_text_draft_case_attachments()
            self._mark_text_draft_temp_dirty()
            self._status_label.set_label(f"Attached {self._autotext_display_name(suggestion.citation)}.")
        entry = self._text_draft_case_search_entry
        if entry is not None:
            entry.set_text("")
            entry.grab_focus()
        self._hide_text_draft_case_suggestions()

    def _remove_text_draft_case_attachment(self, _button: Gtk.Button, citation: str) -> None:
        self._text_draft_case_attachments = [
            existing for existing in self._text_draft_case_attachments if existing != citation
        ]
        self._render_text_draft_case_attachments()
        self._mark_text_draft_temp_dirty()

    def _clear_text_draft_case_attachments(self) -> None:
        if not self._text_draft_case_attachments:
            return
        self._text_draft_case_attachments = []
        self._render_text_draft_case_attachments()
        self._mark_text_draft_temp_dirty()

    def _render_text_draft_case_attachments(self) -> None:
        box = self._text_draft_case_attachments_box
        if box is None:
            return
        self._clear_box(box)
        for start in range(0, len(self._text_draft_case_attachments), 2):
            row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
            row.set_hexpand(True)
            row.set_halign(Gtk.Align.FILL)
            row.set_homogeneous(True)
            for citation in self._text_draft_case_attachments[start:start + 2]:
                chip = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
                chip.set_hexpand(True)
                chip.set_halign(Gtk.Align.FILL)
                chip.add_css_class("text-draft-case-attachment")
                label = Gtk.Label(label=citation, xalign=0)
                label.set_hexpand(True)
                label.set_halign(Gtk.Align.FILL)
                label.set_ellipsize(Pango.EllipsizeMode.END)
                label.set_tooltip_text(citation)
                chip.append(label)
                remove_button = Gtk.Button(icon_name="window-close-symbolic")
                remove_button.add_css_class("flat")
                remove_button.add_css_class("text-draft-case-remove")
                remove_button.set_tooltip_text("Remove attached case.")
                remove_button.connect("clicked", self._remove_text_draft_case_attachment, citation)
                chip.append(remove_button)
                row.append(chip)
            box.append(row)
        box.set_visible(bool(self._text_draft_case_attachments))
        self._refresh_text_draft_clear_button()

    def _ensure_text_draft_temp_path(self) -> Path | None:
        if self._text_draft_temp_path is not None:
            return self._text_draft_temp_path
        try:
            with tempfile.NamedTemporaryFile(prefix="prose-text-draft-", suffix=".txt", delete=False) as handle:
                self._text_draft_temp_path = Path(handle.name)
        except OSError as exc:
            self._show_toast(f"Unable to create text draft temp file: {exc}")
            return None
        return self._text_draft_temp_path

    def _write_text_draft_temp_file_now(self) -> Path | None:
        if self._text_draft_temp_flush_source_id:
            GLib.Source.remove(self._text_draft_temp_flush_source_id)
            self._text_draft_temp_flush_source_id = 0
        path = self._ensure_text_draft_temp_path()
        if path is None:
            return None
        try:
            path.write_text(self._get_text_draft_submission_text(), encoding="utf-8")
            self._text_draft_temp_dirty = False
        except OSError as exc:
            self._show_toast(f"Unable to sync text draft temp file: {exc}")
            return None
        return path

    def _flush_text_draft_temp_file(self) -> bool:
        self._text_draft_temp_flush_source_id = 0
        if not self._text_draft_temp_dirty:
            return GLib.SOURCE_REMOVE
        path = self._ensure_text_draft_temp_path()
        if path is None:
            return GLib.SOURCE_REMOVE
        try:
            path.write_text(self._get_text_draft_submission_text(), encoding="utf-8")
            self._text_draft_temp_dirty = False
        except OSError as exc:
            self._show_toast(f"Unable to sync text draft temp file: {exc}")
        return GLib.SOURCE_REMOVE

    def _cleanup_text_draft_temp_file(self) -> None:
        if self._text_draft_temp_flush_source_id:
            GLib.Source.remove(self._text_draft_temp_flush_source_id)
            self._text_draft_temp_flush_source_id = 0
        self._text_draft_temp_dirty = False
        if self._text_draft_temp_path is None:
            return
        try:
            self._text_draft_temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        self._text_draft_temp_path = None

    def _on_window_close_request(self, _window: Gtk.Window) -> bool:
        if self._text_draft_terminal_active and self._text_draft_terminal_pid is not None:
            try:
                os.kill(self._text_draft_terminal_pid, signal.SIGTERM)
            except OSError:
                pass
            self._text_draft_terminal_active = False
            self._text_draft_terminal_pid = None
        self._cleanup_text_draft_temp_file()
        return False

    def _set_text_draft_original_output_text(self, text: str) -> bool:
        if self._text_draft_original_output_buffer is None:
            return False
        self._text_draft_original_pending_newlines = 0
        self._text_draft_original_output_buffer.set_text(text or "")
        return False

    def _get_text_draft_original_output_text(self) -> str:
        if self._text_draft_original_output_buffer is None:
            return ""
        start, end = self._text_draft_original_output_buffer.get_bounds()
        return self._text_draft_original_output_buffer.get_text(start, end, True)

    def _trim_text_draft_original_output_edges(self) -> None:
        current = self._get_text_draft_original_output_text()
        cleaned = self._normalize_generated_output_text(current)
        if cleaned != current:
            self._set_text_draft_original_output_text(cleaned)

    def _get_text_draft_text(self) -> str:
        if self._text_draft_buffer is None:
            return ""
        start, end = self._text_draft_buffer.get_bounds()
        return self._text_draft_buffer.get_text(start, end, True)

    def _get_text_draft_submission_text(self) -> str:
        draft_text = self._get_text_draft_text()
        if not self._text_draft_case_attachments:
            return draft_text
        cases_text = "\n".join(self._text_draft_case_attachments)
        if draft_text.strip():
            return f"{draft_text.rstrip()}\n\nCases:\n{cases_text}"
        return f"Cases:\n{cases_text}"

    def _clear_text_draft_text(self) -> None:
        if self._text_draft_buffer is None:
            return
        self._clear_text_draft_insert_marks()
        self._text_draft_pending_newlines = 0
        self._text_draft_buffer.set_text("")
        self._clear_text_draft_case_attachments()
        self._hide_text_draft_case_suggestions()
        self._text_draft_temp_dirty = True

    def _text_draft_has_user_text(self) -> bool:
        return bool(self._get_text_draft_submission_text().strip())

    def _refresh_text_draft_clear_button(self) -> None:
        button = self._text_draft_clear_button
        surface = self._text_draft_draft_surface
        if button is None or surface is None:
            return
        button.set_visible(
            surface.get_visible_child_name() == "draft"
            and self._text_draft_has_user_text()
        )

    def _parse_text_draft_template_content(self, raw_text: str) -> tuple[list[tuple[str, str]], str]:
        text = raw_text[1:] if raw_text.startswith("\ufeff") else raw_text
        if not text:
            return [], ""

        emails: list[tuple[str, str]] = []
        position = 0
        text_length = len(text)

        while position < text_length:
            line_end = text.find("\n", position)
            next_position = text_length if line_end == -1 else line_end + 1
            line = text[position:next_position]
            line_for_match = line[:-1] if line.endswith("\n") else line
            if line_for_match.endswith("\r"):
                line_for_match = line_for_match[:-1]
            match = TEXT_DRAFT_TEMPLATE_EMAIL_RE.fullmatch(line_for_match)
            if match is not None:
                emails.append((match.group("label").strip(), match.group("address").strip()))
                position = next_position
                continue

            fallback_match = TEXT_DRAFT_TEMPLATE_EMAIL_FALLBACK_RE.search(line_for_match)
            if fallback_match is None:
                break

            address = fallback_match.group("address").strip()
            label = (
                line_for_match[:fallback_match.start()] + line_for_match[fallback_match.end():]
            ).strip(" <>[]():;,-\t")
            emails.append((label or address, address))
            position = next_position

        if not emails:
            return [], text

        if position < text_length:
            line_end = text.find("\n", position)
            next_position = text_length if line_end == -1 else line_end + 1
            line = text[position:next_position]
            blank_line = line[:-1] if line.endswith("\n") else line
            if blank_line.endswith("\r"):
                blank_line = blank_line[:-1]
            if not blank_line.strip():
                position = next_position

        return emails, text[position:]

    def _list_text_draft_template_paths_in_directory(self, directory: Path) -> list[Path]:
        if not directory.exists() or not directory.is_dir():
            return []
        try:
            return sorted(
                (
                    child for child in directory.iterdir()
                    if child.is_file() and child.suffix.lower() == ".txt"
                ),
                key=lambda path: path.name.lower(),
            )
        except OSError:
            return []

    def _list_text_draft_template_categories(self) -> list[TextDraftTemplateCategory]:
        directory = self._text_draft_template_dir
        if directory is None or not directory.exists() or not directory.is_dir():
            return []

        categories: list[TextDraftTemplateCategory] = []
        try:
            child_directories = sorted(
                (child for child in directory.iterdir() if child.is_dir()),
                key=lambda path: path.name.lower(),
            )
        except OSError:
            return []

        for child_directory in child_directories:
            template_paths = self._list_text_draft_template_paths_in_directory(child_directory)
            if template_paths:
                categories.append(
                    TextDraftTemplateCategory(
                        key=f"dir:{child_directory.name}",
                        name=child_directory.name,
                        template_paths=template_paths,
                    )
                )

        root_templates = self._list_text_draft_template_paths_in_directory(directory)
        if root_templates:
            categories.append(
                TextDraftTemplateCategory(
                    key="__uncategorized__",
                    name="Uncategorized",
                    template_paths=root_templates,
                )
            )
        return categories

    def _find_text_draft_template_category(self, key: str | None) -> TextDraftTemplateCategory | None:
        if key is None:
            return None
        for category in self._text_draft_template_categories:
            if category.key == key:
                return category
        return None

    def _text_draft_template_display_name(self, path: Path) -> str:
        directory = self._text_draft_template_dir
        if directory is None:
            return path.name
        try:
            return str(path.relative_to(directory))
        except ValueError:
            return path.name

    def _show_text_draft_template_categories(self) -> None:
        title = self._text_draft_templates_title
        back_button = self._text_draft_templates_back_button
        grid = self._text_draft_templates_grid
        if title is None or back_button is None or grid is None:
            return
        self._text_draft_template_current_category_key = None
        title.set_label("Templates")
        back_button.set_visible(False)
        self._clear_flow_box(grid)
        self._text_draft_template_menu_buttons = []
        for category in self._text_draft_template_categories:
            button = Gtk.Button(label=category.name)
            count = len(category.template_paths)
            button.set_tooltip_text(f"{count} template{'s' if count != 1 else ''}")
            self._apply_quick_action_button_classes(button)
            button.connect("clicked", self._on_text_draft_template_category_clicked, category.key)
            grid.append(button)
            self._text_draft_template_menu_buttons.append(button)
        self._update_text_draft_template_controls()

    def _show_text_draft_template_category(self, key: str) -> None:
        category = self._find_text_draft_template_category(key)
        if category is None:
            self._show_text_draft_template_categories()
            return
        title = self._text_draft_templates_title
        back_button = self._text_draft_templates_back_button
        grid = self._text_draft_templates_grid
        if title is None or back_button is None or grid is None:
            return
        self._text_draft_template_current_category_key = category.key
        title.set_label(category.name)
        back_button.set_visible(True)
        self._clear_flow_box(grid)
        self._text_draft_template_menu_buttons = []
        popover = self._text_draft_templates_popover
        for path in category.template_paths:
            button = Gtk.Button(label=path.stem)
            button.set_tooltip_text(self._text_draft_template_display_name(path))
            self._apply_quick_action_button_classes(button)
            button.connect("clicked", self._on_text_draft_template_menu_clicked, path, popover)
            grid.append(button)
            self._text_draft_template_menu_buttons.append(button)
        self._update_text_draft_template_controls()

    def _update_text_draft_template_controls(self) -> None:
        button = self._text_draft_template_button
        has_templates = bool(self._text_draft_template_categories)
        if button is not None:
            button.set_sensitive(has_templates and not self._busy)
        if self._text_draft_templates_back_button is not None:
            self._text_draft_templates_back_button.set_sensitive(
                self._text_draft_template_current_category_key is not None and not self._busy
            )
        for widget in self._text_draft_template_menu_buttons:
            widget.set_sensitive(not self._busy)
        for email_button in self._text_draft_template_email_buttons:
            email_button.set_sensitive(not self._busy)

    def _refresh_text_draft_template_email_buttons(self) -> None:
        box = self._text_draft_template_email_box
        if box is None:
            return
        self._clear_box(box)
        self._text_draft_template_email_buttons = []
        for label, address in self._text_draft_template_emails:
            button = Gtk.Button(label=label, icon_name="mail-send-symbolic")
            button.add_css_class("flat")
            button.add_css_class("transform-pill")
            button.add_css_class("transform-pill-compact")
            button.set_tooltip_text(address)
            button.connect("clicked", self._on_copy_text_draft_template_email_clicked, label, address)
            button.set_sensitive(not self._busy)
            box.append(button)
            self._text_draft_template_email_buttons.append(button)
        box.set_visible(bool(self._text_draft_template_emails))

    def _build_text_draft_templates_button(self) -> Gtk.MenuButton:
        popover = Gtk.Popover()
        popover.set_autohide(True)
        popover.set_cascade_popdown(True)
        popover.set_position(Gtk.PositionType.BOTTOM)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=10)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(10)
        content.set_margin_end(10)

        header = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        back_button = Gtk.Button(label="Back", icon_name="go-previous-symbolic")
        back_button.add_css_class("flat")
        back_button.connect("clicked", self._on_text_draft_template_back_clicked)
        back_button.set_visible(False)
        header.append(back_button)

        title = Gtk.Label(label="Templates", xalign=0)
        title.add_css_class("caption")
        title.add_css_class("dim-label")
        title.set_hexpand(True)
        title.set_halign(Gtk.Align.START)
        header.append(title)
        content.append(header)

        action_grid = Gtk.FlowBox()
        action_grid.set_selection_mode(Gtk.SelectionMode.NONE)
        action_grid.set_column_spacing(6)
        action_grid.set_row_spacing(6)
        action_grid.set_max_children_per_line(2)
        content.append(action_grid)

        popover.set_child(content)

        button = Gtk.MenuButton(icon_name="document-open-symbolic")
        button.set_tooltip_text("Templates")
        button.set_popover(popover)
        button.connect("notify::active", self._on_text_draft_template_button_active_changed)
        self._apply_quick_action_button_classes(button)
        self._text_draft_templates_popover = popover
        self._text_draft_templates_grid = action_grid
        self._text_draft_templates_title = title
        self._text_draft_templates_back_button = back_button
        return button

    def _refresh_text_draft_templates(self) -> None:
        if self._text_draft_templates_grid is None:
            return
        self._text_draft_template_categories = self._list_text_draft_template_categories()
        self._show_text_draft_template_categories()

    def _replace_text_draft_with_text(self, text: str) -> None:
        buffer = self._text_draft_buffer
        if buffer is None:
            return
        self._clear_text_draft_insert_marks()
        self._clear_pending_text_draft_regenerate_context()
        self._text_draft_last_regenerate_context = None
        self._set_text_draft_original_output_text("")
        self._clear_text_draft_case_attachments()
        self._text_draft_pending_newlines = 0
        buffer.set_text(text)
        end_iter = buffer.get_end_iter()
        self._set_text_draft_insert_marks(end_iter)
        self._text_draft_temp_dirty = True
        self._focus_text_draft_insert_end()

    def _apply_text_draft_template_path(self, path: Path) -> None:
        try:
            raw_text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            self._status_label.set_label(f"Unable to load template: {exc}")
            self._show_toast(f"Unable to load template: {exc}")
            return
        emails, text = self._parse_text_draft_template_content(raw_text)
        self._text_draft_template_emails = emails
        self._refresh_text_draft_template_email_buttons()
        self._replace_text_draft_with_text(text)
        display_name = self._text_draft_template_display_name(path)
        self._status_label.set_label(f"Loaded text template: {display_name}")
        self._show_toast(f"Loaded template: {display_name}")

    def _load_text_draft_template_path(self, path: Path) -> None:
        if self._busy:
            return
        self._apply_text_draft_template_path(path)

    def _on_text_draft_template_category_clicked(self, _button: Gtk.Button, key: str) -> None:
        if self._busy:
            return
        self._show_text_draft_template_category(key)

    def _on_text_draft_template_back_clicked(self, _button: Gtk.Button) -> None:
        if self._busy:
            return
        self._show_text_draft_template_categories()

    def _on_text_draft_template_button_active_changed(
        self,
        button: Gtk.MenuButton,
        _param_spec: object | None = None,
    ) -> None:
        if button.get_active():
            self._show_text_draft_template_categories()

    def _on_text_draft_template_menu_clicked(
        self,
        _button: Gtk.Button,
        path: Path,
        popover: Gtk.Popover | None,
    ) -> None:
        if popover is not None:
            popover.popdown()
        self._load_text_draft_template_path(path)

    def _on_copy_text_draft_template_email_clicked(
        self,
        _button: Gtk.Button,
        label: str,
        address: str,
    ) -> None:
        if not address:
            self._show_toast("Template email is empty.")
            return
        clipboard = self.get_clipboard()
        clipboard.set(address)
        self._status_label.set_label(f"Copied {label} email to clipboard.")
        self._show_toast(f"Copied {label} email.")

    def _clear_text_draft_insert_marks(self) -> None:
        buffer = self._text_draft_buffer
        if buffer is None:
            self._text_draft_insert_start_mark = None
            self._text_draft_insert_end_mark = None
            return
        for mark in (self._text_draft_insert_start_mark, self._text_draft_insert_end_mark):
            if mark is None:
                continue
            try:
                buffer.delete_mark(mark)
            except Exception:
                pass
        self._text_draft_insert_start_mark = None
        self._text_draft_insert_end_mark = None

    def _set_text_draft_insert_marks(self, start_iter: Gtk.TextIter) -> None:
        buffer = self._text_draft_buffer
        if buffer is None:
            return
        self._clear_text_draft_insert_marks()
        self._text_draft_insert_start_mark = buffer.create_mark(None, start_iter, True)
        self._text_draft_insert_end_mark = buffer.create_mark(None, start_iter, False)

    def _prepare_text_draft_append_output(self) -> bool:
        buffer = self._text_draft_buffer
        if buffer is None:
            return False
        if buffer.get_has_selection():
            return self._prepare_text_draft_selection_replace()
        insert_iter = buffer.get_iter_at_mark(buffer.get_insert())
        self._set_text_draft_insert_marks(insert_iter)
        self._text_draft_pending_newlines = 0
        return True

    def _prepare_text_draft_generated_replace(self) -> bool:
        buffer = self._text_draft_buffer
        start_mark = self._text_draft_insert_start_mark
        end_mark = self._text_draft_insert_end_mark
        if buffer is None or start_mark is None or end_mark is None:
            return False
        start_iter = buffer.get_iter_at_mark(start_mark)
        end_iter = buffer.get_iter_at_mark(end_mark)
        start_offset = start_iter.get_offset()
        buffer.delete(start_iter, end_iter)
        insert_iter = buffer.get_iter_at_offset(start_offset)
        self._set_text_draft_insert_marks(insert_iter)
        self._text_draft_pending_newlines = 0
        self._text_draft_temp_dirty = True
        return True

    def _get_text_draft_selected_text(self) -> str:
        buffer = self._text_draft_buffer
        if buffer is None:
            return ""
        if not buffer.get_has_selection():
            return ""
        selection = buffer.get_selection_bounds()
        if len(selection) == 3:
            has_selection, start, end = selection
            if not has_selection:
                return ""
        else:
            start, end = selection
        return buffer.get_text(start, end, True)

    def _get_text_draft_selection_offsets(self) -> tuple[int, int] | None:
        buffer = self._text_draft_buffer
        if buffer is None or not buffer.get_has_selection():
            return None
        selection = buffer.get_selection_bounds()
        if len(selection) == 3:
            has_selection, start_iter, end_iter = selection
            if not has_selection:
                return None
        else:
            start_iter, end_iter = selection
        return (start_iter.get_offset(), end_iter.get_offset())

    def _prepare_text_draft_selection_replace(self) -> bool:
        buffer = self._text_draft_buffer
        if buffer is None or not buffer.get_has_selection():
            return False
        selection = buffer.get_selection_bounds()
        if len(selection) == 3:
            has_selection, start_iter, end_iter = selection
            if not has_selection:
                return False
        else:
            start_iter, end_iter = selection
        start_offset = start_iter.get_offset()
        buffer.delete(start_iter, end_iter)
        insert_iter = buffer.get_iter_at_offset(start_offset)
        self._set_text_draft_insert_marks(insert_iter)
        self._text_draft_pending_newlines = 0
        self._text_draft_temp_dirty = True
        return True

    def _prepare_text_draft_saved_selection_replace(self) -> bool:
        buffer = self._text_draft_buffer
        start_offset = self._multi_draft_replace_start
        end_offset = self._multi_draft_replace_end
        if buffer is None or not isinstance(start_offset, int) or not isinstance(end_offset, int):
            return False
        try:
            char_count = buffer.get_char_count()
            start_offset = max(0, min(start_offset, char_count))
            end_offset = max(start_offset, min(end_offset, char_count))
            start_iter = buffer.get_iter_at_offset(start_offset)
            end_iter = buffer.get_iter_at_offset(end_offset)
            buffer.delete(start_iter, end_iter)
            insert_iter = buffer.get_iter_at_offset(start_offset)
            self._set_text_draft_insert_marks(insert_iter)
            self._text_draft_pending_newlines = 0
            self._text_draft_temp_dirty = True
            return True
        except Exception:
            return False

    def _append_text_draft_inserted_text(self, text: str) -> bool:
        if not text:
            return False
        self._append_text_draft_buffer_text(
            text,
            is_original_output=False,
        )
        return False

    def _append_text_draft_buffer_text(self, text: str, *, is_original_output: bool) -> bool:
        buffer = self._text_draft_original_output_buffer if is_original_output else self._text_draft_buffer
        if not text or buffer is None:
            return False
        if is_original_output:
            end_iter = buffer.get_end_iter()
            insert = lambda chunk: buffer.insert(end_iter, chunk)
            pending_attr = "_text_draft_original_pending_newlines"
        else:
            if self._text_draft_insert_end_mark is None:
                return False
            insert_iter = buffer.get_iter_at_mark(self._text_draft_insert_end_mark)
            insert = lambda chunk: buffer.insert(insert_iter, chunk)
            pending_attr = "_text_draft_pending_newlines"

        pending_newlines = int(getattr(self, pending_attr, 0) or 0)
        parts, pending_newlines = self._split_text_for_paragraphs(text, pending_newlines)
        for kind, chunk in parts:
            if kind == "paragraph":
                insert("\n\n")
                continue
            if not chunk:
                continue
            insert(chunk)
        setattr(self, pending_attr, pending_newlines)
        self._text_draft_temp_dirty = True
        return False

    def _append_text_draft_original_output_text(self, text: str) -> bool:
        if not text:
            return False
        self._append_text_draft_buffer_text(text, is_original_output=True)
        return False

    def _flush_text_draft_pending_newlines(self, *, is_original_output: bool) -> None:
        buffer = self._text_draft_original_output_buffer if is_original_output else self._text_draft_buffer
        if buffer is None:
            return
        pending_attr = "_text_draft_original_pending_newlines" if is_original_output else "_text_draft_pending_newlines"
        pending_newlines = int(getattr(self, pending_attr, 0) or 0)
        if pending_newlines <= 0:
            return
        if is_original_output:
            end_iter = buffer.get_end_iter()
            insert = lambda chunk: buffer.insert(end_iter, chunk)
        else:
            if self._text_draft_insert_end_mark is None:
                return
            insert_iter = buffer.get_iter_at_mark(self._text_draft_insert_end_mark)
            insert = lambda chunk: buffer.insert(insert_iter, chunk)
        breaks = pending_newlines // 2
        for _ in range(breaks):
            insert("\n\n")
        if pending_newlines % 2:
            insert("\n")
        setattr(self, pending_attr, 0)
        self._text_draft_temp_dirty = True

    def _trim_text_draft_trailing_whitespace(self) -> None:
        buffer = self._text_draft_buffer
        start_mark = self._text_draft_insert_start_mark
        end_mark = self._text_draft_insert_end_mark
        if buffer is None or start_mark is None or end_mark is None:
            return
        start_iter = buffer.get_iter_at_mark(start_mark)
        end_iter = buffer.get_iter_at_mark(end_mark)
        current = buffer.get_text(start_iter, end_iter, True)
        if not current:
            return
        trimmed = current.rstrip()
        if trimmed == current:
            return
        start_offset = start_iter.get_offset()
        end_offset = end_iter.get_offset()
        delete_start = buffer.get_iter_at_offset(start_offset)
        delete_end = buffer.get_iter_at_offset(end_offset)
        buffer.delete(delete_start, delete_end)
        insert_iter = buffer.get_iter_at_offset(start_offset)
        self._set_text_draft_insert_marks(insert_iter)
        if trimmed:
            self._append_text_draft_inserted_text(trimmed)

    def _ensure_single_text_draft_trailing_space(self) -> None:
        buffer = self._text_draft_buffer
        end_mark = self._text_draft_insert_end_mark
        if buffer is None or end_mark is None:
            return
        self._trim_text_draft_trailing_whitespace()
        end_iter = buffer.get_iter_at_mark(end_mark)
        if end_iter.ends_line() or end_iter.is_end():
            self._append_text_draft_inserted_text(" ")
            return
        following = end_iter.get_char()
        if following != " ":
            self._append_text_draft_inserted_text(" ")

    def _focus_text_draft_insert_end(self) -> None:
        buffer = self._text_draft_buffer
        view = self._text_draft_view
        end_mark = self._text_draft_insert_end_mark
        if buffer is None or view is None or end_mark is None:
            return
        end_iter = buffer.get_iter_at_mark(end_mark)
        buffer.place_cursor(end_iter)
        insert_mark = buffer.get_insert()
        view.scroll_mark_onscreen(insert_mark)
        view.grab_focus()

    def _capture_spellingstyle_range_end(self) -> None:
        if not self._editor_insert_doc or not self._editor_insert_cursor:
            return
        try:
            text = self._get_text_container(self._editor_insert_doc, self._editor_insert_cursor)
            if not text:
                return
            self._editor_insert_end = text.createTextCursorByRange(self._editor_insert_cursor)
        except Exception:
            self._editor_insert_end = None

    def _capture_improve1_range_end(self) -> None:
        if not self._improve_insert_doc or not self._improve_insert_cursor:
            return
        try:
            text = self._get_text_container(self._improve_insert_doc, self._improve_insert_cursor)
            if not text:
                return
            self._editor_insert_end = text.createTextCursorByRange(self._improve_insert_cursor)
        except Exception:
            self._editor_insert_end = None

    def _capture_improve2_range_end(self) -> None:
        if not self._improve_insert_doc or not self._improve_insert_cursor:
            return
        try:
            text = self._get_text_container(self._improve_insert_doc, self._improve_insert_cursor)
            if not text:
                return
            self._editor_insert_end = text.createTextCursorByRange(self._improve_insert_cursor)
        except Exception:
            self._editor_insert_end = None

    def _select_spellingstyle_range(self, doc: XTextDocument) -> bool:  # type: ignore[type-arg]
        if not self._editor_insert_end or self._last_insert_len <= 0:
            return False
        try:
            text = self._get_text_container(doc, self._editor_insert_end)
            if not text:
                return False
            range_cursor = text.createTextCursorByRange(self._editor_insert_end)
            if not range_cursor.goLeft(self._last_insert_len, True):
                return False
            view_cursor = doc.getCurrentController().getViewCursor()
            view_cursor.gotoRange(range_cursor.getStart(), False)
            view_cursor.gotoRange(range_cursor.getEnd(), True)
            return True
        except Exception:
            return False

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if hasattr(self, "_status_spinner"):
            self._status_spinner.set_visible(busy)
            self._status_spinner.set_spinning(busy)
        if hasattr(self, "_run_btn"):
            self._run_btn.set_sensitive(not busy)
        if hasattr(self, "_load_json_btn"):
            self._load_json_btn.set_sensitive(not busy)
        if hasattr(self, "_spell_btn"):
            self._spell_btn.set_sensitive(not busy)
        if hasattr(self, "_direct_input_btn"):
            self._direct_input_btn.set_sensitive(not busy)
        if hasattr(self, "_combine_cites_btn"):
            self._combine_cites_btn.set_sensitive(not busy)
        if hasattr(self, "_improve1_btn"):
            self._improve1_btn.set_sensitive(not busy)
        if hasattr(self, "_improve2_btn"):
            self._improve2_btn.set_sensitive(not busy)
        if hasattr(self, "_improve_selected_btn"):
            self._improve_selected_btn.set_sensitive(not busy)
        if hasattr(self, "_keep_original_btn"):
            self._keep_original_btn.set_sensitive(not busy)
        if hasattr(self, "_thesaurus_btn"):
            self._thesaurus_btn.set_sensitive(not busy)
        if hasattr(self, "_reference_btn"):
            self._reference_btn.set_sensitive(not busy)
        if getattr(self, "_regenerate_menu_button", None) is not None:
            self._regenerate_menu_button.set_sensitive(self._regenerate_profile_chip_sensitive())
        if hasattr(self, "_regenerate_profile_buttons"):
            for button in self._regenerate_profile_buttons:
                button.set_sensitive(self._regenerate_profile_chip_sensitive())
        if getattr(self, "_text_draft_regenerate_menu_button", None) is not None:
            self._text_draft_regenerate_menu_button.set_sensitive(self._text_draft_regenerate_profile_chip_sensitive())
        if hasattr(self, "_text_draft_regenerate_profile_buttons"):
            for button in self._text_draft_regenerate_profile_buttons:
                button.set_sensitive(self._text_draft_regenerate_profile_chip_sensitive())
        if hasattr(self, "_transform_action_buttons"):
            for button in self._transform_action_buttons:
                button.set_sensitive(not busy)
        if hasattr(self, "_text_draft_action_buttons"):
            for button in self._text_draft_action_buttons:
                button.set_sensitive(not busy)
        if hasattr(self, "_text_draft_external_action_buttons"):
            for button in self._text_draft_external_action_buttons:
                button.set_sensitive(not busy)
        if hasattr(self, "_text_draft_terminal_buttons"):
            for button in self._text_draft_terminal_buttons:
                button.set_sensitive(not busy)
        if hasattr(self, "_update_text_draft_template_controls"):
            self._update_text_draft_template_controls()

    # Proofreading pipeline ----------------------------------------------
    def _proofreading_scope_label(self, start: int | None, end: int | None) -> str:
        if start is None or end is None:
            return "entire document"
        return f"pages {start}-{end}"

    def _gather_and_request(
        self,
        doc: XTextDocument,  # type: ignore[type-arg]
        start: int | None,
        end: int | None,
        profile: ModelProfile,
    ) -> None:
        paragraphs = self._collect_proofreading_paragraphs(doc)
        if not paragraphs:
            GLib.idle_add(self._on_request_failed, "No text found in the open document.")
            return
        scope_label = self._proofreading_scope_label(start, end)
        scoped_paragraphs = self._filter_proofreading_paragraphs_by_page(paragraphs, start, end)
        if not scoped_paragraphs:
            GLib.idle_add(self._on_request_failed, f"No paragraphs found in {scope_label}.")
            return
        chunks = self._build_proofreading_chunks(scoped_paragraphs)
        if not chunks:
            GLib.idle_add(self._on_request_failed, f"Unable to build proofreading chunks for {scope_label}.")
            return

        merged: list[Suggestion] = []
        total_chunks = len(chunks)
        for chunk in chunks:
            GLib.idle_add(
                self._set_proofreading_progress,
                chunk.index,
                total_chunks,
                profile.display_name(),
                scope_label,
            )
            payload = self._compose_llm_payload(chunk, total_chunks, profile)
            try:
                suggestions = self._call_llm(
                    payload,
                    profile,
                    allow_empty=True,
                    request_title=f"Proof Reading Chunk {chunk.index} of {total_chunks}",
                )
            except Exception as exc:  # noqa: BLE001
                GLib.idle_add(
                    self._on_llm_failed,
                    f"Chunk {chunk.index} of {total_chunks} failed: {exc}",
                )
                return
            merged.extend(suggestions)

        deduped = self._dedupe_suggestions(merged)
        if not deduped:
            GLib.idle_add(self._on_request_failed, "No suggestions returned after reviewing the document.")
            return
        GLib.idle_add(self._on_request_finished, deduped, total_chunks)

    def _add_model_id(
        self,
        payload: dict[str, Any],
        model_id: str,
        disable_reasoning: bool = False,
    ) -> dict[str, Any]:
        cleaned = model_id.strip()
        if cleaned:
            payload["model"] = cleaned
        _apply_disable_reasoning_to_body(
            payload,
            model_id=cleaned,
            disable_reasoning=disable_reasoning,
        )
        return payload

    def _should_retry_without_reasoning_controls(
        self,
        payload: dict[str, Any],
        message: str,
        *,
        attempted_without_thinking: bool,
        attempted_without_reasoning_effort: bool,
    ) -> tuple[bool, bool, bool]:
        lowered = (message or "").lower()
        if (
            not attempted_without_thinking
            and "thinking" in payload
            and "thinking" in lowered
            and any(marker in lowered for marker in ("unsupported", "unknown", "invalid"))
        ):
            payload.pop("thinking", None)
            return True, True, attempted_without_reasoning_effort
        if (
            not attempted_without_reasoning_effort
            and "reasoning_effort" in payload
            and "reasoning_effort" in lowered
            and any(marker in lowered for marker in ("unsupported", "unknown", "invalid"))
        ):
            payload.pop("reasoning_effort", None)
            return True, attempted_without_thinking, True
        return False, attempted_without_thinking, attempted_without_reasoning_effort

    def _clear_model_response_audit(self) -> None:
        self._last_model_response_raw = None
        self._last_model_response_title = None
        self._last_model_response_api_url = None
        self._last_model_response_timestamp = None
        self._last_model_response_diagnostic = None
        GLib.idle_add(self._set_model_response_view_available, False)

    def _remember_editor_request(self, title: str, api_url: str, payload: Any) -> None:
        self._clear_model_response_audit()
        try:
            raw = json.dumps(payload, ensure_ascii=False)
        except Exception:
            raw = str(payload)
        self._last_editor_request_raw = raw
        self._last_editor_request_title = title
        self._last_editor_request_api_url = api_url
        self._last_editor_request_timestamp = datetime.now().strftime("%B %d, %Y %I:%M:%S %p")
        GLib.idle_add(self._set_editor_prompt_view_available, True)

    def _set_editor_prompt_view_available(self, available: bool) -> bool:
        action = self.get_application().lookup_action("view-last-editor-prompt")
        if action is not None:
            action.set_enabled(available)
        return False

    def _remember_model_response(
        self,
        title: str,
        api_url: str,
        payload: Any,
        raw_response: str,
        response_obj: Any | None = None,
    ) -> None:
        self._last_model_response_raw = raw_response
        self._last_model_response_title = title
        self._last_model_response_api_url = api_url
        self._last_model_response_timestamp = datetime.now().strftime("%B %d, %Y %I:%M:%S %p")
        self._last_model_response_diagnostic = self._build_model_response_diagnostic(
            payload,
            raw_response,
            response_obj=response_obj,
        )
        GLib.idle_add(self._set_model_response_view_available, True)

    def _set_model_response_view_available(self, available: bool) -> bool:
        action = self.get_application().lookup_action("view-last-model-response")
        if action is not None:
            action.set_enabled(available)
        return False

    def _build_model_response_diagnostic(
        self,
        payload: Any,
        raw_response: str,
        *,
        response_obj: Any | None = None,
    ) -> str:
        if response_obj is None:
            try:
                response_obj = json.loads(raw_response)
            except Exception:
                response_obj = raw_response
        usage = self._find_last_dict_value(response_obj, "usage")
        perf_metrics = self._find_last_dict_value(response_obj, "perf_metrics")
        return "\n".join(
            [
                f"Model: {self._audit_payload_model(payload)}",
                f"Stream: {self._audit_payload_stream(payload)}",
                f"Reasoning control sent: {self._audit_reasoning_control(payload)}",
                f"reasoning_content seen: {self._yes_no(self._response_has_reasoning_content(response_obj))}",
                f"<think> tag seen: {self._yes_no(self._response_has_think_tag(response_obj))}",
                f"Usage: {self._format_usage_for_audit(usage)}",
                f"Performance metrics: {self._format_perf_metrics_for_audit(perf_metrics)}",
            ]
        )

    def _audit_payload_model(self, payload: Any) -> str:
        if isinstance(payload, dict) and payload.get("model"):
            return str(payload["model"])
        return "not sent"

    def _audit_payload_stream(self, payload: Any) -> str:
        if isinstance(payload, dict) and "stream" in payload:
            return str(bool(payload.get("stream"))).lower()
        return "not sent"

    def _audit_reasoning_control(self, payload: Any) -> str:
        if not isinstance(payload, dict):
            return "not sent"
        controls: list[str] = []
        if "reasoning_effort" in payload:
            controls.append(f"reasoning_effort={self._json_audit_value(payload.get('reasoning_effort'))}")
        if "thinking" in payload:
            controls.append(f"thinking={self._json_audit_value(payload.get('thinking'))}")
        return ", ".join(controls) if controls else "not sent"

    def _json_audit_value(self, value: Any) -> str:
        try:
            return json.dumps(value, ensure_ascii=False)
        except Exception:
            return str(value)

    def _json_value_has_content(self, value: Any) -> bool:
        if value is None:
            return False
        if isinstance(value, str):
            return bool(value.strip())
        if isinstance(value, dict):
            return any(self._json_value_has_content(item) for item in value.values())
        if isinstance(value, list):
            return any(self._json_value_has_content(item) for item in value)
        return bool(value)

    def _response_has_reasoning_content(self, value: Any) -> bool:
        if isinstance(value, dict):
            for key, item in value.items():
                if key == "reasoning_content" and self._json_value_has_content(item):
                    return True
                if self._response_has_reasoning_content(item):
                    return True
        elif isinstance(value, list):
            return any(self._response_has_reasoning_content(item) for item in value)
        return False

    def _response_has_think_tag(self, value: Any) -> bool:
        if isinstance(value, str):
            lowered = value.lower()
            return "<think" in lowered or "</think" in lowered
        if isinstance(value, dict):
            return any(self._response_has_think_tag(item) for item in value.values())
        if isinstance(value, list):
            return any(self._response_has_think_tag(item) for item in value)
        return False

    def _find_last_dict_value(self, value: Any, key: str) -> dict[str, Any] | None:
        found: dict[str, Any] | None = None
        if isinstance(value, dict):
            item = value.get(key)
            if isinstance(item, dict) and item:
                found = item
            for child in value.values():
                child_found = self._find_last_dict_value(child, key)
                if child_found is not None:
                    found = child_found
        elif isinstance(value, list):
            for child in value:
                child_found = self._find_last_dict_value(child, key)
                if child_found is not None:
                    found = child_found
        return found

    def _format_usage_for_audit(self, usage: dict[str, Any] | None) -> str:
        if not usage:
            return "not returned"
        parts: list[str] = []
        for label, key in (
            ("prompt", "prompt_tokens"),
            ("completion", "completion_tokens"),
            ("total", "total_tokens"),
            ("reasoning", "reasoning_tokens"),
        ):
            if usage.get(key) is not None:
                parts.append(f"{label}={usage[key]}")
        details = usage.get("prompt_tokens_details")
        if isinstance(details, dict) and details.get("cached_tokens") is not None:
            parts.append(f"cached_prompt={details['cached_tokens']}")
        completion_details = usage.get("completion_tokens_details")
        if isinstance(completion_details, dict) and completion_details.get("reasoning_tokens") is not None:
            parts.append(f"reasoning={completion_details['reasoning_tokens']}")
        return ", ".join(parts) if parts else self._json_audit_value(usage)

    def _format_perf_metrics_for_audit(self, perf_metrics: dict[str, Any] | None) -> str:
        if not perf_metrics:
            return "not returned"
        return self._json_audit_value(perf_metrics)

    def _yes_no(self, value: bool) -> str:
        return "yes" if value else "no"

    def _post_json_and_read(
        self,
        payload: dict[str, Any],
        api_url: str,
        api_key: str,
        *,
        accept: str | None = None,
        request_title: str | None = None,
    ) -> str:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Prose/1.0",
        }
        if accept:
            headers["Accept"] = accept

        attempted_without_thinking = False
        attempted_without_reasoning_effort = False
        while True:
            if request_title:
                self._remember_editor_request(request_title, api_url, payload)
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(api_url, data=data, headers=headers, method="POST")
            try:
                with urllib.request.urlopen(req) as resp:
                    raw = resp.read().decode("utf-8", errors="ignore")
                    if request_title:
                        self._remember_model_response(request_title, api_url, payload, raw)
                    return raw
            except urllib.error.HTTPError as exc:
                detail = ""
                try:
                    detail = exc.read().decode("utf-8", errors="ignore").strip()
                except Exception:
                    pass
                should_retry, attempted_without_thinking, attempted_without_reasoning_effort = (
                    self._should_retry_without_reasoning_controls(
                        payload,
                        detail or str(exc.reason or "request failed"),
                        attempted_without_thinking=attempted_without_thinking,
                        attempted_without_reasoning_effort=attempted_without_reasoning_effort,
                    )
                )
                if should_retry:
                    continue
                if detail:
                    raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'} - {detail}") from exc
                raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'}") from exc

    def _is_gemini_generate_content_url(self, api_url: str) -> bool:
        lowered = api_url.lower()
        return "generativelanguage.googleapis.com" in lowered and ":generatecontent" in lowered

    def _call_gemini_generate_content(
        self,
        api_url: str,
        api_key: str,
        text: str,
        request_title: str | None = None,
    ) -> str:
        headers = {
            "Content-Type": "application/json",
            "x-goog-api-key": api_key,
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Prose/1.0",
        }
        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": text},
                    ]
                }
            ]
        }
        if request_title:
            self._remember_editor_request(request_title, api_url, payload)
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(api_url, data=data, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req) as resp:
                raw = resp.read().decode("utf-8", errors="ignore")
                if request_title:
                    self._remember_model_response(request_title, api_url, payload, raw)
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8", errors="ignore").strip()
            except Exception:
                pass
            if detail:
                raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'} - {detail}") from exc
            raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'}") from exc
        try:
            data_obj = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Gemini response was not valid JSON: {exc}") from exc
        parts: list[str] = []
        for candidate in data_obj.get("candidates", []):
            if not isinstance(candidate, dict):
                continue
            content = candidate.get("content") or {}
            if not isinstance(content, dict):
                continue
            for part in content.get("parts", []) or []:
                if isinstance(part, dict) and part.get("text"):
                    parts.append(str(part["text"]))
        output = self._normalize_generated_output_text("".join(parts))
        if not output:
            raise ValueError("Gemini returned empty output.")
        return output

    def _filter_proofreading_paragraphs_by_page(
        self,
        paragraphs: list[ProofreadParagraph],
        start: int | None,
        end: int | None,
    ) -> list[str]:
        if start is None or end is None:
            return [paragraph.text for paragraph in paragraphs if paragraph.text]
        scoped = [
            paragraph.text
            for paragraph in paragraphs
            if paragraph.text and (paragraph.page is None or start <= paragraph.page <= end)
        ]
        if scoped:
            return scoped
        # Fallback for environments where page numbers cannot be read reliably.
        if any(paragraph.page is None for paragraph in paragraphs):
            return [paragraph.text for paragraph in paragraphs if paragraph.text]
        return []

    def _build_proofreading_chunks(self, paragraphs: list[str]) -> list[ProofreadChunk]:
        chunks: list[ProofreadChunk] = []
        current: list[str] = []
        current_chars = 0
        for paragraph in paragraphs:
            paragraph = paragraph.strip()
            if not paragraph:
                continue
            separator_len = 2 if current else 0
            would_exceed_count = len(current) >= PROOFREAD_CHUNK_PARAGRAPH_LIMIT
            would_exceed_chars = bool(current) and (
                current_chars + separator_len + len(paragraph) > PROOFREAD_CHUNK_CHAR_LIMIT
            )
            if would_exceed_count or would_exceed_chars:
                chunks.append(ProofreadChunk(index=len(chunks) + 1, paragraphs=current))
                current = []
                current_chars = 0
            if current:
                current_chars += 2
            current.append(paragraph)
            current_chars += len(paragraph)
        if current:
            chunks.append(ProofreadChunk(index=len(chunks) + 1, paragraphs=current))
        return chunks

    def _compose_llm_payload(
        self,
        chunk: ProofreadChunk,
        total_chunks: int,
        profile: ModelProfile,
    ) -> dict[str, Any]:
        system_prompt = _expand_shared_prompt_parts(self._proof_settings.prompt or DEFAULT_PROMPT)
        instructions = (
            "Return a JSON array of suggested edits for ONLY the provided Writer paragraph batch. "
            "Each item must look like: {"
            '"title": short summary, '
            '"snippet": exact text that should be replaced, '
            '"replacement": improved text, '
            '"reasoning": short explanation'
            "}. "
            "Use standard JSON syntax (double quotes for keys and string values). "
            "Keep snippets short but unique. Only propose changes that truly improve clarity or correctness. "
            "For snippet values, copy exact text from the paragraph batch and preserve punctuation and spacing when possible. "
            'Do not include a "page" field. '
            "Do not return comments, summaries, or unchanged text."
        )
        paragraph_blocks = [
            f"[Paragraph {offset}]\n{text}"
            for offset, text in enumerate(chunk.paragraphs, start=1)
            if text.strip()
        ]
        content = (
            f"Chunk {chunk.index} of {total_chunks}\n"
            f"Paragraph count: {len(paragraph_blocks)}\n\n"
            + "\n\n".join(paragraph_blocks)
        )
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": instructions},
                {"role": "user", "content": content},
            ],
            "stream": False,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _call_llm(
        self,
        payload: dict[str, Any],
        profile: ModelProfile,
        allow_empty: bool = False,
        request_title: str = "Proof Reading",
    ) -> list[Suggestion]:
        raw = self._post_json_and_read(
            payload,
            profile.api_url,
            profile.api_key,
            request_title=request_title,
        )
        suggestions = self._parse_suggestions(raw, allow_empty=allow_empty)
        return suggestions

    def _sort_suggestions(self, suggestions: list[Suggestion]) -> list[Suggestion]:
        return list(suggestions)

    def _normalize_suggestion_key_part(self, value: str) -> str:
        return re.sub(r"\s+", " ", str(value or "").strip()).casefold()

    def _dedupe_suggestions(self, suggestions: list[Suggestion]) -> list[Suggestion]:
        deduped: list[Suggestion] = []
        seen: set[tuple[str, str]] = set()
        for suggestion in suggestions:
            key = (
                self._normalize_suggestion_key_part(suggestion.snippet),
                self._normalize_suggestion_key_part(suggestion.replacement),
            )
            if not all(key) or key in seen:
                continue
            seen.add(key)
            deduped.append(suggestion)
        return deduped

    def _parse_suggestions(self, raw: str, allow_empty: bool = False) -> list[Suggestion]:
        def _normalize_json_punctuation(text: str) -> str:
            # Some models return JSON-like text with curly quotes; normalize for JSON parsing.
            return text.translate(
                {
                    ord("“"): '"',
                    ord("”"): '"',
                    ord("‘"): "'",
                    ord("’"): "'",
                }
            )

        def _extract_message_content(message: Any) -> str:
            if not isinstance(message, dict):
                return ""
            content = message.get("content")
            if isinstance(content, str):
                return content
            if isinstance(content, dict):
                text = content.get("text")
                if isinstance(text, str):
                    return text
                if isinstance(text, dict):
                    value = text.get("value")
                    if isinstance(value, str):
                        return value
            if isinstance(content, list):
                parts: list[str] = []
                for block in content:
                    if not isinstance(block, dict):
                        continue
                    block_type = str(block.get("type") or "").lower()
                    if block_type and ("reason" in block_type or "think" in block_type):
                        continue
                    text = block.get("text")
                    if isinstance(text, str) and text:
                        parts.append(text)
                    elif isinstance(text, dict):
                        value = text.get("value")
                        if isinstance(value, str) and value:
                            parts.append(value)
                    elif block.get("output_text"):
                        parts.append(str(block.get("output_text")))
                return "".join(parts)
            return ""

        def _extract_function_arguments(message: Any) -> str:
            if not isinstance(message, dict):
                return ""
            tool_calls = message.get("tool_calls")
            if not isinstance(tool_calls, list):
                return ""
            for tool_call in tool_calls:
                if not isinstance(tool_call, dict):
                    continue
                function_obj = tool_call.get("function")
                if not isinstance(function_obj, dict):
                    continue
                arguments = function_obj.get("arguments")
                if isinstance(arguments, str) and arguments.strip():
                    return arguments
            return ""

        def _load_json(text: str):
            cleaned = text.strip()
            if not cleaned:
                raise ValueError("Empty response from model.")
            candidate = cleaned
            if cleaned.startswith("```"):
                fence_match = re.match(r"^```(?:json)?\s*([\s\S]*?)\s*```$", cleaned, flags=re.IGNORECASE)
                if fence_match:
                    candidate = fence_match.group(1).strip()
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                return json.loads(_normalize_json_punctuation(candidate))

        def _coerce_list(obj: Any) -> list[dict[str, Any]]:
            if isinstance(obj, list):
                return [item for item in obj if isinstance(item, dict)]
            if isinstance(obj, dict):
                if any(key in obj for key in ("snippet", "replacement", "source", "original", "suggestion", "corrected")):
                    return [obj]
                for key in ("suggestions", "edits", "data", "changes", "corrections"):
                    if isinstance(obj.get(key), list):
                        return [item for item in obj[key] if isinstance(item, dict)]
            return []

        def _try_extract_embedded_json(text: str) -> Any:
            # Try fenced JSON first.
            fenced = re.findall(r"```(?:json)?\s*([\s\S]*?)```", text, flags=re.IGNORECASE)
            for candidate in fenced:
                candidate = candidate.strip()
                if not candidate:
                    continue
                try:
                    return _load_json(candidate)
                except Exception:
                    continue

            # Then try non-greedy array/object candidates.
            for pattern in (r"\[[\s\S]*?\]", r"\{[\s\S]*?\}"):
                for match in re.finditer(pattern, text):
                    candidate = match.group(0).strip()
                    if not candidate:
                        continue
                    try:
                        return _load_json(candidate)
                    except Exception:
                        continue
            raise ValueError("LLM response was not valid JSON.")

        try:
            data = _load_json(raw)
        except json.JSONDecodeError as exc:
            try:
                data = _try_extract_embedded_json(raw)
            except ValueError:
                raise ValueError(f"LLM response was not valid JSON: {exc}") from exc

        # OpenAI-style shape: {choices: [{message: {content: "...json..."}}]}
        if isinstance(data, dict) and "choices" in data:
            content_fragments: list[str] = []
            for choice in data.get("choices", []):
                if not isinstance(choice, dict):
                    continue
                msg = choice.get("message") or choice.get("delta") or {}
                content = _extract_message_content(msg)
                if content:
                    content_fragments.append(content)
                    continue
                arguments = _extract_function_arguments(msg)
                if arguments:
                    content_fragments.append(arguments)
            content = "".join(content_fragments).strip()
            if content:
                try:
                    data = _load_json(content)
                except json.JSONDecodeError as exc:
                    try:
                        data = _try_extract_embedded_json(content)
                    except ValueError:
                        raise ValueError(f"LLM content was not valid JSON: {exc}") from exc

        candidates = _coerce_list(data)
        suggestions: list[Suggestion] = []
        for item in candidates:
            snippet = str(
                item.get("snippet")
                or item.get("source")
                or item.get("original")
                or item.get("from")
                or item.get("text")
                or ""
            ).strip()
            replacement = str(
                item.get("replacement")
                or item.get("suggestion")
                or item.get("corrected")
                or item.get("to")
                or item.get("edit")
                or ""
            ).strip()
            if not snippet or not replacement:
                continue  # skip unusable entries
            page_value = (
                item.get("page")
                or item.get("page_number")
                or item.get("pageNum")
                or item.get("pagenum")
            )
            page_match = re.search(r"\d+", str(page_value or ""))
            page = int(page_match.group(0)) if page_match else None
            suggestions.append(
                Suggestion(
                    title=str(item.get("title") or "Suggested change"),
                    page=page,
                    snippet=snippet,
                    replacement=replacement,
                    reasoning=str(item.get("reasoning") or item.get("why") or item.get("explanation") or ""),
                )
            )
        if not suggestions and not allow_empty:
            raise ValueError("No suggestions returned by the model.")
        return suggestions

    def _on_request_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        return False

    def _set_proofreading_progress(
        self,
        chunk_index: int,
        total_chunks: int,
        profile_name: str,
        scope_label: str,
    ) -> bool:
        self._status_label.set_label(
            f"Proofreading {scope_label}: chunk {chunk_index} of {total_chunks} with {profile_name}…"
        )
        return False

    def _on_llm_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._notify_llm_error(message)
        return False

    def _on_request_finished(self, suggestions: list[Suggestion], total_chunks: int = 1) -> bool:
        self._set_busy(False)
        chunk_text = "chunk" if total_chunks == 1 else "chunks"
        self._status_label.set_label(
            f"Received {len(suggestions)} suggestions from {total_chunks} {chunk_text}."
        )
        self._empty_label.set_label("Proof reading suggestions will appear here after Prose finishes batching the document.")
        self._suggestions = self._sort_suggestions(suggestions)
        self._render_suggestions()
        self._focus_first_suggestion(notify_success=False)
        return False

    def _run_spellingstyle(
        self,
        doc: XTextDocument,
        source_text: str,
        profile: ModelProfile,
    ) -> None:  # type: ignore[type-arg]
        try:
            payload = self._compose_spellingstyle_payload(source_text, profile)
            stream = self._stream_spellingstyle(payload, profile)
            for chunk in stream:
                GLib.idle_add(self._append_editor_text, chunk)
                GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_spellingstyle_finished, "SpellingStyle complete.")

    def _run_text_draft_spellingstyle(
        self,
        source_text: str,
        profile: ModelProfile,
        route_to_terminal: bool = False,
        terminal_pid: int | None = None,
    ) -> None:
        generated_parts: list[str] = []
        try:
            payload = self._compose_text_draft_spellingstyle_payload(source_text, profile)
            for chunk in self._stream_spellingstyle(payload, profile):
                if route_to_terminal:
                    if not self._text_draft_terminal_child_is_running(terminal_pid):
                        raise RuntimeError("Embedded Codex is no longer running.")
                    generated_parts.append(chunk)
                    GLib.idle_add(self._append_text_draft_terminal_text, chunk)
                else:
                    GLib.idle_add(self._append_text_draft_inserted_text, chunk)
                    GLib.idle_add(self._append_text_draft_original_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_text_draft_failed, str(exc))
            return
        if route_to_terminal:
            GLib.idle_add(self._set_text_draft_terminal_generated_output, "".join(generated_parts))
        GLib.idle_add(
            self._on_text_draft_spellingstyle_finished,
            "Text Draft SpellingStyle streamed to embedded Codex." if route_to_terminal else "Text Draft SpellingStyle complete.",
            route_to_terminal,
        )

    def _run_improve(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_improve_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Improve"):
                GLib.idle_add(self._append_improve1_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_improve_finished, f"Improve Generated complete with {profile.display_name()}.")

    def _run_text_draft_improve(
        self,
        source_text: str,
        profile: ModelProfile,
        route_to_terminal: bool = False,
        terminal_pid: int | None = None,
    ) -> None:
        payload = self._compose_improve_payload(source_text, profile)
        generated_parts: list[str] = []
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Improve"):
                if route_to_terminal:
                    if not self._text_draft_terminal_child_is_running(terminal_pid):
                        raise RuntimeError("Embedded Codex is no longer running.")
                    generated_parts.append(chunk)
                    GLib.idle_add(self._append_text_draft_terminal_text, chunk)
                else:
                    GLib.idle_add(self._append_text_draft_inserted_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_text_draft_failed, str(exc))
            return
        if route_to_terminal:
            GLib.idle_add(self._set_text_draft_terminal_generated_output, "".join(generated_parts))
        GLib.idle_add(
            self._on_text_draft_improve_finished,
            (
                f"Improve Generated streamed to embedded Codex with {profile.display_name()}."
                if route_to_terminal
                else f"Improve Generated complete with {profile.display_name()}."
            ),
            route_to_terminal,
        )

    def _run_rephrase_generated(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_rephrase_generated_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Rephrase"):
                GLib.idle_add(self._append_improve1_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_improve_finished, f"Rephrase Generated complete with {profile.display_name()}.")

    def _run_text_draft_rephrase_generated(
        self,
        source_text: str,
        profile: ModelProfile,
        route_to_terminal: bool = False,
        terminal_pid: int | None = None,
    ) -> None:
        payload = self._compose_rephrase_generated_payload(source_text, profile)
        generated_parts: list[str] = []
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Rephrase"):
                if route_to_terminal:
                    if not self._text_draft_terminal_child_is_running(terminal_pid):
                        raise RuntimeError("Embedded Codex is no longer running.")
                    generated_parts.append(chunk)
                    GLib.idle_add(self._append_text_draft_terminal_text, chunk)
                else:
                    GLib.idle_add(self._append_text_draft_inserted_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_text_draft_failed, str(exc))
            return
        if route_to_terminal:
            GLib.idle_add(self._set_text_draft_terminal_generated_output, "".join(generated_parts))
        GLib.idle_add(
            self._on_text_draft_improve_finished,
            (
                f"Rephrase Generated streamed to embedded Codex with {profile.display_name()}."
                if route_to_terminal
                else f"Rephrase Generated complete with {profile.display_name()}."
            ),
            route_to_terminal,
        )

    def _run_improve_selected(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_improve_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Improve"):
                GLib.idle_add(self._append_improve1_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(
            self._on_improve_selected_finished,
            f"Improve Selected complete with {profile.display_name()}.",
        )

    def _run_text_draft_improve_selected(
        self,
        source_text: str,
        profile: ModelProfile,
        route_to_terminal: bool = False,
        terminal_pid: int | None = None,
    ) -> None:
        payload = self._compose_improve_payload(source_text, profile)
        generated_parts: list[str] = []
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Improve"):
                if route_to_terminal:
                    if not self._text_draft_terminal_child_is_running(terminal_pid):
                        raise RuntimeError("Embedded Codex is no longer running.")
                    generated_parts.append(chunk)
                    GLib.idle_add(self._append_text_draft_terminal_text, chunk)
                else:
                    GLib.idle_add(self._append_text_draft_inserted_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_text_draft_failed, str(exc))
            return
        if route_to_terminal:
            GLib.idle_add(self._set_text_draft_terminal_generated_output, "".join(generated_parts))
        GLib.idle_add(
            self._on_text_draft_improve_finished,
            (
                f"Improve Selected streamed to embedded Codex with {profile.display_name()}."
                if route_to_terminal
                else f"Improve Selected complete with {profile.display_name()}."
            ),
            route_to_terminal,
        )

    def _run_shorten(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_shorten_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(payload, profile.api_url, profile.api_key, request_title="Shorten"):
                GLib.idle_add(self._append_improve1_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_shorten_finished, f"Shorten complete with {profile.display_name()}.")

    def _run_topic_sentence(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_topic_sentence_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(
                payload,
                profile.api_url,
                profile.api_key,
                request_title="Topic Sentence",
            ):
                GLib.idle_add(self._append_editor_text, chunk)
                GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(
            self._on_topic_sentence_finished,
            f"Topic sentence complete with {profile.display_name()}.",
        )

    def _run_introduction(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_introduction_payload(source_text, profile)
        try:
            if self._is_gemini_generate_content_url(profile.api_url):
                prompt = _expand_shared_prompt_parts(self._introduction_settings.prompt or DEFAULT_INTRO_PROMPT)
                combined = f"{prompt}\n\n{source_text}" if prompt else source_text
                output = self._call_gemini_generate_content(
                    profile.api_url,
                    profile.api_key,
                    combined,
                    request_title="Introduction",
                )
                GLib.idle_add(self._append_editor_text, output)
                GLib.idle_add(self._append_spelling_output_text, output)
            else:
                for chunk in self._stream_custom(
                    payload,
                    profile.api_url,
                    profile.api_key,
                    request_title="Introduction",
                ):
                    GLib.idle_add(self._append_editor_text, chunk)
                    GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_introduction_finished, f"Introduction complete with {profile.display_name()}.")

    def _run_introduction_reply(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_introduction_reply_payload(source_text, profile)
        try:
            if self._is_gemini_generate_content_url(profile.api_url):
                prompt = _expand_shared_prompt_parts(self._introduction_reply_settings.prompt or DEFAULT_INTRO_REPLY_PROMPT)
                combined = f"{prompt}\n\n{source_text}" if prompt else source_text
                output = self._call_gemini_generate_content(
                    profile.api_url,
                    profile.api_key,
                    combined,
                    request_title="Introduction Reply",
                )
                GLib.idle_add(self._append_editor_text, output)
                GLib.idle_add(self._append_spelling_output_text, output)
            else:
                for chunk in self._stream_custom(
                    payload,
                    profile.api_url,
                    profile.api_key,
                    request_title="Introduction Reply",
                ):
                    GLib.idle_add(self._append_editor_text, chunk)
                    GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(
            self._on_introduction_finished,
            f"Introduction for reply complete with {profile.display_name()}.",
        )

    def _run_conclusion(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_conclusion_payload(source_text, profile)
        try:
            if self._is_gemini_generate_content_url(profile.api_url):
                prompt = _expand_shared_prompt_parts(self._conclusion_settings.prompt or DEFAULT_CONCLUSION_PROMPT)
                combined = f"{prompt}\n\n{source_text}" if prompt else source_text
                output = self._call_gemini_generate_content(
                    profile.api_url,
                    profile.api_key,
                    combined,
                    request_title="Conclusion",
                )
                GLib.idle_add(self._append_editor_text, output)
                GLib.idle_add(self._append_spelling_output_text, output)
            else:
                for chunk in self._stream_custom(
                    payload,
                    profile.api_url,
                    profile.api_key,
                    request_title="Conclusion",
                ):
                    GLib.idle_add(self._append_editor_text, chunk)
                    GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(self._on_conclusion_finished, f"Conclusion complete with {profile.display_name()}.")

    def _run_conclusion_no_issues(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_conclusion_no_issues_payload(source_text, profile)
        try:
            if self._is_gemini_generate_content_url(profile.api_url):
                prompt = _expand_shared_prompt_parts(self._concl_no_issues_settings.prompt or DEFAULT_CONCL_NO_ISSUES_PROMPT)
                combined = f"{prompt}\n\n{source_text}" if prompt else source_text
                output = self._call_gemini_generate_content(
                    profile.api_url,
                    profile.api_key,
                    combined,
                    request_title="Conclusion (No Issues)",
                )
                GLib.idle_add(self._append_editor_text, output)
                GLib.idle_add(self._append_spelling_output_text, output)
            else:
                for chunk in self._stream_custom(
                    payload,
                    profile.api_url,
                    profile.api_key,
                    request_title="Conclusion (No Issues)",
                ):
                    GLib.idle_add(self._append_editor_text, chunk)
                    GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(
            self._on_conclusion_no_issues_finished,
            f"Conclusion (no issues) complete with {profile.display_name()}.",
        )

    def _run_concl_section(self, source_text: str, profile: ModelProfile) -> None:
        payload = self._compose_concl_section_payload(source_text, profile)
        try:
            for chunk in self._stream_custom(
                payload,
                profile.api_url,
                profile.api_key,
                request_title="Section Conclusion",
            ):
                GLib.idle_add(self._append_editor_text, chunk)
                GLib.idle_add(self._append_spelling_output_text, chunk)
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))
            return
        GLib.idle_add(
            self._on_concl_section_finished,
            f"Section conclusion complete with {profile.display_name()}.",
        )

    def _run_translate_document(self, doc: XTextDocument, profile: ModelProfile) -> None:  # type: ignore[type-arg]
        try:
            paragraphs = self._collect_document_paragraph_texts(doc)
            if not paragraphs:
                GLib.idle_add(self._on_spellingstyle_failed, "Document has no translatable text.")
                return
            translated_paragraphs: list[str] = []
            for source_text in paragraphs:
                if not source_text.strip():
                    translated_paragraphs.append(source_text)
                    continue
                translated_text = self._translate_unit_text(source_text, strict=False, profile=profile)
                if not translated_text.strip():
                    translated_text = source_text
                translated_paragraphs.append(translated_text)
            normalized_paragraphs = [p.replace("\r\n", "\n").replace("\r", "\n").strip("\n") for p in translated_paragraphs]
            self._replace_document_text_plain(doc, "\n\n".join(normalized_paragraphs))
            GLib.idle_add(self._on_translate_finished, "Document translated to Spanish.")
        except Exception as exc:  # noqa: BLE001
            GLib.idle_add(self._on_spellingstyle_failed, str(exc))

    def _collect_document_paragraph_texts(self, doc: XTextDocument) -> list[str]:  # type: ignore[type-arg]
        try:
            doc_text = doc.getText()
            enum = doc_text.createEnumeration()
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Unable to read Writer document: {exc}") from exc
        paragraphs: list[str] = []
        while enum.hasMoreElements():
            paragraph = enum.nextElement()
            try:
                paragraph_text = str(paragraph.getString() or "")
            except Exception:
                paragraph_text = ""
            paragraphs.append(paragraph_text)
        return paragraphs

    def _replace_document_text_plain(self, doc: XTextDocument, text: str) -> None:  # type: ignore[type-arg]
        try:
            doc_text = doc.getText()
            doc_text.setString(text)
            cursor = doc_text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            slant_none = self._get_font_slant("NONE")
            for prop_name in ("CharPosture", "CharPostureAsian", "CharPostureComplex"):
                try:
                    cursor.setPropertyValue(prop_name, slant_none)
                except Exception:
                    continue
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Unable to replace document text: {exc}") from exc

    def _translate_paragraph_by_sentences(self, paragraph: str, profile: ModelProfile) -> str:
        parts = self._split_paragraph_for_translation_fallback(paragraph)
        translated_parts: list[str] = []
        for part in parts:
            if not part.strip():
                translated_parts.append(part)
                continue
            translated = self._translate_unit_text(part, strict=True, profile=profile)
            if self._translation_looks_truncated(part, translated):
                return paragraph
            translated_parts.append(translated)
        return "".join(translated_parts)

    def _split_paragraph_for_translation_fallback(self, paragraph: str) -> list[str]:
        chunks = re.split(r"(?<=[.!?:;])(\s+)", paragraph)
        if not chunks:
            return [paragraph]
        parts: list[str] = []
        idx = 0
        while idx < len(chunks):
            text = chunks[idx]
            ws = chunks[idx + 1] if idx + 1 < len(chunks) else ""
            parts.append(f"{text}{ws}")
            idx += 2
        return parts if parts else [paragraph]

    def _collect_translatable_paragraphs(
        self,
        doc: XTextDocument,  # type: ignore[type-arg]
    ) -> list[tuple[int, str, list[tuple[str, Any]]]]:
        paragraphs: list[tuple[int, str, list[tuple[str, Any]]]] = []
        index = 0
        try:
            doc_text = doc.getText()
            index = self._collect_translatable_paragraphs_in_text(doc_text, paragraphs, index)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Unable to read Writer document: {exc}") from exc
        for note_text in self._iter_note_texts(doc):
            index = self._collect_translatable_paragraphs_in_text(note_text, paragraphs, index)
        return paragraphs

    def _collect_translatable_paragraphs_in_text(
        self,
        text_obj: Any,
        target: list[tuple[int, str, list[tuple[str, Any]]]],
        start_index: int,
    ) -> int:
        try:
            enum = text_obj.createEnumeration()
        except Exception:
            return start_index
        index = start_index
        while enum.hasMoreElements():
            paragraph = enum.nextElement()
            try:
                portion_enum = paragraph.createEnumeration()
            except Exception:
                continue
            portion_entries: list[tuple[str, Any]] = []
            while portion_enum.hasMoreElements():
                portion = portion_enum.nextElement()
                try:
                    portion_type = str(portion.getPropertyValue("TextPortionType") or "")
                except Exception:
                    portion_type = ""
                if portion_type != "Text":
                    continue
                text = str(portion.getString() or "")
                if not text:
                    continue
                try:
                    cursor = text_obj.createTextCursorByRange(portion.getStart())
                    cursor.gotoRange(portion.getEnd(), True)
                except Exception:
                    continue
                portion_entries.append((text, cursor))
            if not portion_entries:
                continue
            paragraph_text = "".join(text for text, _cursor in portion_entries)
            if not paragraph_text.strip():
                continue
            target.append((index, paragraph_text, portion_entries))
            index += 1
        return index

    def _distribute_text_across_portions(self, text: str, source_portions: list[str]) -> list[str]:
        count = len(source_portions)
        if count <= 1:
            return [text]
        if not text:
            return [""] * count
        source_lengths = [max(1, len(part)) for part in source_portions]
        total_source = sum(source_lengths)
        total_target = len(text)
        boundaries: list[int] = []
        running_source = 0
        previous = 0
        for idx in range(count - 1):
            running_source += source_lengths[idx]
            target = int(round((running_source / total_source) * total_target))
            target = max(previous, min(target, total_target))
            target = self._snap_split_boundary(text, target, previous, total_target)
            boundaries.append(target)
            previous = target
        parts: list[str] = []
        start = 0
        for boundary in boundaries:
            parts.append(text[start:boundary])
            start = boundary
        parts.append(text[start:])
        if len(parts) < count:
            parts.extend([""] * (count - len(parts)))
        return parts[:count]

    def _snap_split_boundary(self, text: str, target: int, lower: int, upper: int, window: int = 20) -> int:
        if target <= lower:
            return lower
        if target >= upper:
            return upper
        best = target
        for distance in range(window + 1):
            left = target - distance
            right = target + distance
            candidates = []
            if lower < left < upper:
                candidates.append(left)
            if lower < right < upper and right != left:
                candidates.append(right)
            for candidate in candidates:
                if text[candidate - 1].isspace() or text[candidate].isspace():
                    return candidate
        return best

    def _collect_translatable_portions(self, doc: XTextDocument) -> list[tuple[int, str, Any]]:  # type: ignore[type-arg]
        collected: list[tuple[int, str, Any]] = []
        index = 0
        try:
            doc_text = doc.getText()
            index = self._collect_translatable_portions_in_text(doc_text, collected, index)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"Unable to read Writer document: {exc}") from exc
        for note_text in self._iter_note_texts(doc):
            index = self._collect_translatable_portions_in_text(note_text, collected, index)
        return collected

    def _collect_translatable_portions_in_text(
        self,
        text_obj: Any,
        target: list[tuple[int, str, Any]],
        start_index: int,
    ) -> int:
        try:
            enum = text_obj.createEnumeration()
        except Exception:
            return start_index
        index = start_index
        while enum.hasMoreElements():
            paragraph = enum.nextElement()
            try:
                portion_enum = paragraph.createEnumeration()
            except Exception:
                continue
            while portion_enum.hasMoreElements():
                portion = portion_enum.nextElement()
                try:
                    portion_type = str(portion.getPropertyValue("TextPortionType") or "")
                except Exception:
                    portion_type = ""
                if portion_type != "Text":
                    continue
                text = str(portion.getString() or "")
                if not text:
                    continue
                try:
                    cursor = text_obj.createTextCursorByRange(portion.getStart())
                    cursor.gotoRange(portion.getEnd(), True)
                except Exception:
                    continue
                target.append((index, text, cursor))
                index += 1
        return index

    def _build_translation_units(
        self,
        portions: list[tuple[int, str, Any]],
    ) -> list[tuple[int, int, int, str]]:
        units: list[tuple[int, int, int, str]] = []
        unit_index = 0
        for portion_index, text, _cursor in portions:
            chunks = self._split_text_for_translation(text)
            if not chunks:
                chunks = [text]
            for chunk_order, chunk_text in enumerate(chunks):
                units.append((unit_index, portion_index, chunk_order, chunk_text))
                unit_index += 1
        return units

    def _split_text_for_translation(self, text: str, max_chunk_chars: int = 700) -> list[str]:
        if not text:
            return []
        if len(text) <= max_chunk_chars:
            return [text]
        parts: list[str] = []
        start = 0
        text_len = len(text)
        while start < text_len:
            end = min(start + max_chunk_chars, text_len)
            if end < text_len:
                cut = text.rfind("\n", start + 1, end)
                if cut <= start:
                    cut = text.rfind(" ", start + 1, end)
                if cut > start:
                    end = cut + 1
            chunk = text[start:end]
            if not chunk:
                break
            parts.append(chunk)
            start = end
        return parts

    def _build_translation_batches(
        self,
        portions: list[tuple[int, str, Any]] | list[tuple[int, int, int, str]],
        max_chars: int = 3500,
        max_items: int = 24,
    ) -> list[list[tuple[int, str, Any]]]:
        normalized: list[tuple[int, str, Any]] = []
        for item in portions:
            if len(item) == 3:
                index, text, cursor = item  # type: ignore[misc]
                normalized.append((index, text, cursor))
            elif len(item) == 4:
                index, _portion_index, _chunk_order, text = item  # type: ignore[misc]
                normalized.append((index, text, None))
        batches: list[list[tuple[int, str, Any]]] = []
        current: list[tuple[int, str, Any]] = []
        current_chars = 0
        for item in normalized:
            item_chars = len(item[1])
            if current and (len(current) >= max_items or current_chars + item_chars > max_chars):
                batches.append(current)
                current = []
                current_chars = 0
            current.append(item)
            current_chars += item_chars
        if current:
            batches.append(current)
        return batches

    def _translate_unit_text(self, source_text: str, strict: bool, profile: ModelProfile) -> str:
        prompt = self._build_translate_instructions(strict, expect_json=False)
        if profile.api_url.rstrip("/").endswith("/responses"):
            payload = {
                "input": (
                    f"{prompt}\n\n"
                    "Return only the translated text.\n\n"
                    "SOURCE:\n"
                    f"{source_text}"
                ),
                "stream": False,
            }
            payload = self._add_model_id(
                payload,
                profile.model_id,
                disable_reasoning=profile.disable_reasoning,
            )
            return self._call_responses_text(
                payload,
                profile.api_url,
                profile.api_key,
                request_title="Translate",
            )
        payload = {
            "messages": [
                {"role": "system", "content": prompt},
                {
                    "role": "user",
                    "content": (
                        "Return only the translated text.\n\n"
                        f"{source_text}"
                    ),
                },
            ],
            "stream": False,
        }
        payload = self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )
        return self._call_chat_text(
            payload,
            profile.api_url,
            profile.api_key,
            request_title="Translate",
        )

    def _translate_batch(self, batch: list[tuple[int, str, Any]], profile: ModelProfile) -> dict[int, str]:
        translated = self._translate_batch_once(batch, strict=False, profile=profile)
        source_by_index = {index: text for index, text, _cursor in batch}
        suspect = self._find_suspect_translations(source_by_index, translated)
        if not suspect:
            return translated
        for index in sorted(suspect):
            source_text = source_by_index[index]
            single_item = [(index, source_text, None)]
            translated.update(self._translate_batch_once(single_item, strict=True, profile=profile))
        suspect_after_retry = self._find_suspect_translations(source_by_index, translated)
        if suspect_after_retry:
            for index in suspect_after_retry:
                translated[index] = source_by_index[index]
            current_count = int(getattr(self, "_translate_fallback_count", 0) or 0)
            self._translate_fallback_count = current_count + len(suspect_after_retry)
        return translated

    def _translate_batch_once(
        self,
        batch: list[tuple[int, str, Any]],
        strict: bool,
        profile: ModelProfile,
    ) -> dict[int, str]:
        source_payload = [{"index": index, "text": text} for index, text, _cursor in batch]
        instructions = self._build_translate_instructions(strict)
        source_json = json.dumps(source_payload, ensure_ascii=False)
        if profile.api_url.rstrip("/").endswith("/responses"):
            request_payload = {
                "input": f"{instructions}\n\nSOURCE:\n{source_json}",
                "stream": False,
            }
            request_payload = self._add_model_id(
                request_payload,
                profile.model_id,
                disable_reasoning=profile.disable_reasoning,
            )
            raw_output = self._call_responses_text(
                request_payload,
                profile.api_url,
                profile.api_key,
                request_title="Translate",
            )
        else:
            request_payload = {
                "messages": [
                    {"role": "system", "content": instructions},
                    {"role": "user", "content": source_json},
                ],
                "stream": False,
            }
            request_payload = self._add_model_id(
                request_payload,
                profile.model_id,
                disable_reasoning=profile.disable_reasoning,
            )
            raw_output = self._call_chat_text(
                request_payload,
                profile.api_url,
                profile.api_key,
                request_title="Translate",
            )
        return self._parse_translation_batch(raw_output, {index for index, _text, _cursor in batch})

    def _build_translate_instructions(self, strict: bool, expect_json: bool = True) -> str:
        prompt = _expand_shared_prompt_parts(self._translate_settings.prompt or DEFAULT_TRANSLATE_PROMPT)
        if expect_json:
            base = (
                f"{prompt}\n\n"
                "Translate each item's text into Spanish.\n"
                "Return JSON only using this exact shape:\n"
                '{"translations":[{"index":0,"text":"..."},{"index":1,"text":"..."}]}\n'
                "Keep the same index values from the source.\n"
                "Never omit lines, headings, case numbers, section numbers, dates, or citations.\n"
                "Preserve line breaks and paragraph breaks.\n"
                "If any source text is already Spanish, keep it in Spanish.\n"
                "Do not summarize."
            )
        else:
            base = (
                f"{prompt}\n\n"
                "Translate the provided text into Spanish.\n"
                "Never omit lines, headings, case numbers, section numbers, dates, or citations.\n"
                "Preserve line breaks and paragraph breaks.\n"
                "If source text is already Spanish, keep it in Spanish.\n"
                "Do not summarize."
            )
        if not strict:
            return base
        if expect_json:
            return (
                base
                + "\n"
                + "STRICT CHECK: the translated text for each item must fully cover the source item with no omissions."
            )
        return base + "\nSTRICT CHECK: the translated text must fully cover the source text with no omissions."

    def _find_suspect_translations(
        self,
        source_by_index: dict[int, str],
        translated_by_index: dict[int, str],
    ) -> set[int]:
        suspect: set[int] = set()
        for index, source_text in source_by_index.items():
            translated_text = translated_by_index.get(index, "")
            if self._translation_looks_truncated(source_text, translated_text):
                suspect.add(index)
        return suspect

    def _translation_looks_truncated(self, source: str, translated: str) -> bool:
        source_clean = source.strip()
        translated_clean = translated.strip()
        if not source_clean:
            return False
        if not translated_clean:
            return True
        if len(source_clean) >= 30 and len(translated_clean) < 8:
            return True
        if len(source_clean) >= 80 and len(translated_clean) < int(len(source_clean) * 0.35):
            return True
        if len(source_clean) >= 40 and re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]$", translated_clean):
            if translated_clean.endswith(" d") or translated_clean.endswith(" e"):
                return True
        source_numbers = set(re.findall(r"\d[\d./,-]*", source_clean))
        translated_numbers = set(re.findall(r"\d[\d./,-]*", translated_clean))
        if source_numbers and not source_numbers.issubset(translated_numbers):
            return True
        source_end_punct = bool(re.search(r"[.:;!?)]\s*$", source_clean))
        translated_end_punct = bool(re.search(r"[.:;!?)]\s*$", translated_clean))
        if source_end_punct and not translated_end_punct and len(translated_clean) < int(len(source_clean) * 0.7):
            return True
        source_lines = [line for line in source_clean.splitlines() if line.strip()]
        translated_lines = [line for line in translated_clean.splitlines() if line.strip()]
        if len(source_lines) >= 2 and len(translated_lines) == 0:
            return True
        return False

    def _call_chat_text(
        self,
        payload: dict[str, Any],
        api_url: str,
        api_key: str,
        request_title: str | None = None,
    ) -> str:
        raw = self._post_json_and_read(payload, api_url, api_key, request_title=request_title)
        text = self._normalize_generated_output_text("".join(self._extract_response_text(raw)))
        if not text:
            raise ValueError("Translate returned empty output.")
        return text

    def _call_responses_text(
        self,
        payload: dict[str, Any],
        api_url: str,
        api_key: str,
        request_title: str | None = None,
    ) -> str:
        raw = self._post_json_and_read(payload, api_url, api_key, request_title=request_title)
        text = self._normalize_generated_output_text("".join(self._extract_responses_text(raw)))
        if not text:
            raise ValueError("Translate returned empty output.")
        return text

    def _parse_translation_batch(self, raw_output: str, expected_indexes: set[int]) -> dict[int, str]:
        cleaned = raw_output.strip()
        if not cleaned:
            raise ValueError("Translate returned empty output.")
        if cleaned.startswith("```"):
            fence_match = re.match(r"^```(?:json)?\s*([\s\S]*?)\s*```$", cleaned, flags=re.IGNORECASE)
            if fence_match:
                cleaned = fence_match.group(1).strip()
        try:
            data = json.loads(cleaned)
        except json.JSONDecodeError as exc:
            raise ValueError(f"Translate output was not valid JSON: {exc}") from exc
        entries = []
        if isinstance(data, dict):
            entries = data.get("translations") or data.get("items") or []
        elif isinstance(data, list):
            entries = data
        if not isinstance(entries, list):
            raise ValueError("Translate output did not include a translations list.")
        translated: dict[int, str] = {}
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            index_raw = entry.get("index")
            text_raw = entry.get("text")
            try:
                index = int(index_raw)
            except (TypeError, ValueError):
                continue
            if index not in expected_indexes:
                continue
            translated[index] = str(text_raw or "")
        missing = expected_indexes - set(translated.keys())
        if missing:
            missing_preview = ", ".join(str(i) for i in sorted(missing)[:5])
            raise ValueError(f"Translate output missed indexes: {missing_preview}")
        return translated

    def _compose_spellingstyle_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        system_prompt = _expand_shared_prompt_parts(self._spelling_settings.prompt or DEFAULT_SPELLINGSTYLE_PROMPT)
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": source_text},
            ],
            "stream": True,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _compose_text_draft_spellingstyle_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        system_prompt = _expand_shared_prompt_parts(
            self._text_draft_spelling_prompt or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        )
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": source_text},
            ],
            "stream": True,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _compose_improve_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._improve1_settings.prompt,
            DEFAULT_IMPROVE_PROMPT,
            profile,
        )

    def _compose_rephrase_generated_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._improve2_settings.prompt,
            DEFAULT_IMPROVE2_PROMPT,
            profile,
        )

    def _compose_profile_prompt_payload(
        self,
        source_text: str,
        prompt_text: str,
        default_prompt: str,
        profile: ModelProfile,
    ) -> dict[str, Any]:
        system_prompt = _expand_shared_prompt_parts(prompt_text or default_prompt)
        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": source_text},
            ],
            "stream": True,
        }
        return self._add_model_id(
            payload,
            profile.model_id,
            disable_reasoning=profile.disable_reasoning,
        )

    def _compose_shorten_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._shorten_settings.prompt,
            DEFAULT_SHORTEN_PROMPT,
            profile,
        )

    def _compose_topic_sentence_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._topic_sentence_settings.prompt,
            DEFAULT_TOPIC_SENTENCE_PROMPT,
            profile,
        )

    def _compose_introduction_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._introduction_settings.prompt,
            DEFAULT_INTRO_PROMPT,
            profile,
        )

    def _compose_introduction_reply_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._introduction_reply_settings.prompt,
            DEFAULT_INTRO_REPLY_PROMPT,
            profile,
        )

    def _compose_conclusion_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._conclusion_settings.prompt,
            DEFAULT_CONCLUSION_PROMPT,
            profile,
        )

    def _compose_conclusion_no_issues_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._concl_no_issues_settings.prompt,
            DEFAULT_CONCL_NO_ISSUES_PROMPT,
            profile,
        )

    def _compose_concl_section_payload(self, source_text: str, profile: ModelProfile) -> dict[str, Any]:
        return self._compose_profile_prompt_payload(
            source_text,
            self._concl_section_settings.prompt,
            DEFAULT_CONCL_SECTION_PROMPT,
            profile,
        )

    def _stream_spellingstyle(self, payload: dict[str, Any], profile: ModelProfile) -> Iterable[str]:
        yield from self._stream_custom(
            payload,
            profile.api_url,
            profile.api_key,
            request_title="SpellingStyle",
        )

    def _stream_custom(
        self,
        payload: dict[str, Any],
        api_url: str,
        api_key: str,
        request_title: str | None = None,
    ) -> Iterable[str]:
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Prose/1.0",
        }
        attempted_without_thinking = False
        attempted_without_reasoning_effort = False
        saw_output = False
        while True:
            if request_title:
                self._remember_editor_request(request_title, api_url, payload)
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(api_url, data=data, headers=headers, method="POST")
            stream_events: list[Any] = []
            try:
                with urllib.request.urlopen(req) as resp:
                    content_type = resp.headers.get("Content-Type", "")
                    if "text/event-stream" not in content_type:
                        raw = resp.read().decode("utf-8", errors="ignore")
                        if request_title:
                            self._remember_model_response(request_title, api_url, payload, raw)
                        for chunk in self._extract_response_text(raw):
                            cleaned = self._normalize_generated_output_chunk(chunk, is_first_chunk=not saw_output)
                            if not cleaned:
                                continue
                            saw_output = True
                            yield cleaned
                        return
                    for raw_line in resp:
                        line = raw_line.decode("utf-8", errors="ignore").strip()
                        if not line or not line.startswith("data:"):
                            continue
                        payload_str = line[5:].strip()
                        if payload_str == "[DONE]":
                            break
                        try:
                            data_obj = json.loads(payload_str)
                        except json.JSONDecodeError:
                            continue
                        stream_events.append(data_obj)
                        chunk = self._extract_stream_delta(data_obj)
                        if chunk:
                            cleaned = self._normalize_generated_output_chunk(chunk, is_first_chunk=not saw_output)
                            if not cleaned:
                                continue
                            saw_output = True
                            yield cleaned
                    if request_title:
                        raw = json.dumps(stream_events, ensure_ascii=False)
                        self._remember_model_response(request_title, api_url, payload, raw, response_obj=stream_events)
                return
            except urllib.error.HTTPError as exc:
                detail = ""
                try:
                    detail = exc.read().decode("utf-8", errors="ignore").strip()
                except Exception:
                    pass
                should_retry, attempted_without_thinking, attempted_without_reasoning_effort = (
                    self._should_retry_without_reasoning_controls(
                        payload,
                        detail or str(exc.reason or "request failed"),
                        attempted_without_thinking=attempted_without_thinking,
                        attempted_without_reasoning_effort=attempted_without_reasoning_effort,
                    )
                )
                if should_retry:
                    continue
                if detail:
                    raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'} - {detail}") from exc
                raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'}") from exc

    def _stream_responses(
        self,
        payload: dict[str, Any],
        api_url: str,
        api_key: str,
        request_title: str | None = None,
    ) -> Iterable[str]:
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Prose/1.0",
        }
        attempted_without_thinking = False
        attempted_without_reasoning_effort = False
        saw_output = False
        while True:
            if request_title:
                self._remember_editor_request(request_title, api_url, payload)
            data = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(api_url, data=data, headers=headers, method="POST")
            stream_events: list[Any] = []
            try:
                with urllib.request.urlopen(req) as resp:
                    content_type = resp.headers.get("Content-Type", "")
                    if "text/event-stream" not in content_type:
                        raw = resp.read().decode("utf-8", errors="ignore")
                        if request_title:
                            self._remember_model_response(request_title, api_url, payload, raw)
                        for chunk in self._extract_responses_text(raw):
                            cleaned = self._normalize_generated_output_chunk(chunk, is_first_chunk=not saw_output)
                            if not cleaned:
                                continue
                            saw_output = True
                            yield cleaned
                        return
                    for raw_line in resp:
                        line = raw_line.decode("utf-8", errors="ignore").strip()
                        if not line or not line.startswith("data:"):
                            continue
                        payload_str = line[5:].strip()
                        if payload_str == "[DONE]":
                            break
                        try:
                            data_obj = json.loads(payload_str)
                        except json.JSONDecodeError:
                            continue
                        stream_events.append(data_obj)
                        chunk = self._extract_responses_delta(data_obj)
                        if chunk:
                            cleaned = self._normalize_generated_output_chunk(chunk, is_first_chunk=not saw_output)
                            if not cleaned:
                                continue
                            saw_output = True
                            yield cleaned
                    if request_title:
                        raw = json.dumps(stream_events, ensure_ascii=False)
                        self._remember_model_response(request_title, api_url, payload, raw, response_obj=stream_events)
                return
            except urllib.error.HTTPError as exc:
                detail = ""
                try:
                    detail = exc.read().decode("utf-8", errors="ignore").strip()
                except Exception:
                    pass
                should_retry, attempted_without_thinking, attempted_without_reasoning_effort = (
                    self._should_retry_without_reasoning_controls(
                        payload,
                        detail or str(exc.reason or "request failed"),
                        attempted_without_thinking=attempted_without_thinking,
                        attempted_without_reasoning_effort=attempted_without_reasoning_effort,
                    )
                )
                if should_retry:
                    continue
                if detail:
                    raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'} - {detail}") from exc
                raise ValueError(f"HTTP {exc.code}: {exc.reason or 'request failed'}") from exc

    def _stream_reference_responses(self, payload: dict[str, Any]) -> Iterable[str]:
        return self._stream_responses(payload, self._reference_settings.api_url, self._reference_settings.api_key)

    def _extract_responses_delta(self, data: dict[str, Any]) -> str:
        if not isinstance(data, dict):
            return ""
        event_type = str(data.get("type") or "")
        if event_type in ("response.output_text.delta", "response.output_text"):
            for key in ("delta", "text", "output_text"):
                if data.get(key):
                    return str(data[key])
        if event_type == "response.completed":
            response_obj = data.get("response")
            if isinstance(response_obj, dict) and response_obj.get("output_text"):
                return str(response_obj["output_text"])
        return ""

    def _extract_responses_text(self, raw: str) -> Iterable[str]:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"LLM response was not valid JSON: {exc}") from exc
        if isinstance(data, dict):
            output_text = data.get("output_text")
            if output_text:
                return [str(output_text)]
            output = data.get("output")
            if isinstance(output, list):
                parts: list[str] = []
                for item in output:
                    if not isinstance(item, dict):
                        continue
                    if item.get("type") != "message":
                        continue
                    content = item.get("content", [])
                    if not isinstance(content, list):
                        continue
                    for block in content:
                        if not isinstance(block, dict):
                            continue
                        if block.get("type") == "output_text" and block.get("text"):
                            parts.append(str(block["text"]))
                if parts:
                    return ["".join(parts)]
        if isinstance(data, str):
            return [data]
        raise ValueError("LLM response did not include content.")

    def _extract_stream_delta(self, data: dict[str, Any]) -> str:
        if not isinstance(data, dict):
            return ""
        choices = data.get("choices", [])
        if not isinstance(choices, list):
            return ""
        parts: list[str] = []
        for choice in choices:
            if not isinstance(choice, dict):
                continue
            delta = choice.get("delta") or choice.get("message") or {}
            if isinstance(delta, dict) and delta.get("content"):
                parts.append(str(delta["content"]))
        return "".join(parts)

    def _extract_response_text(self, raw: str) -> Iterable[str]:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(f"LLM response was not valid JSON: {exc}") from exc
        if isinstance(data, dict) and "choices" in data:
            parts: list[str] = []
            for choice in data.get("choices", []):
                if not isinstance(choice, dict):
                    continue
                message = choice.get("message", {})
                if isinstance(message, dict) and message.get("content"):
                    parts.append(str(message["content"]))
            if parts:
                return ["".join(parts)]
        if isinstance(data, dict) and data.get("content"):
            return [str(data["content"])]
        if isinstance(data, str):
            return [data]
        raise ValueError("LLM response did not include content.")

    def _on_spellingstyle_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._clear_pending_regenerate_context()
        self._clear_pending_text_draft_regenerate_context()
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._flush_pending_newlines(
                self._improve_insert_doc,
                self._improve_insert_cursor,
                "_improve_pending_newlines",
            )
            self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
        self._notify_llm_error(message)
        return False

    def _on_text_draft_failed(self, message: str) -> bool:
        self._set_busy(False)
        self._clear_pending_text_draft_regenerate_context()
        self._flush_text_draft_pending_newlines(is_original_output=False)
        self._flush_text_draft_pending_newlines(is_original_output=True)
        self._notify_llm_error(message)
        return False

    def _on_spellingstyle_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._capture_spellingstyle_range_end()
        return False

    def _on_text_draft_spellingstyle_finished(self, message: str, route_to_terminal: bool = False) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if route_to_terminal:
            return False
        self._flush_text_draft_pending_newlines(is_original_output=False)
        self._flush_text_draft_pending_newlines(is_original_output=True)
        self._ensure_single_text_draft_trailing_space()
        self._focus_text_draft_insert_end()
        self._trim_text_draft_original_output_edges()
        self._text_draft_temp_dirty = True
        return False

    def _on_improve_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._flush_pending_newlines(
                self._improve_insert_doc,
                self._improve_insert_cursor,
                "_improve_pending_newlines",
            )
            self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._improve_insert_doc,
                self._improve_insert_cursor,
            )
        self._commit_pending_regenerate_context()
        self._capture_improve1_range_end()
        return False

    def _on_text_draft_improve_finished(self, message: str, route_to_terminal: bool = False) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if route_to_terminal:
            self._commit_pending_text_draft_regenerate_context()
            return False
        self._flush_text_draft_pending_newlines(is_original_output=False)
        self._ensure_single_text_draft_trailing_space()
        self._focus_text_draft_insert_end()
        self._commit_pending_text_draft_regenerate_context()
        self._text_draft_temp_dirty = True
        return False

    def _on_improve_selected_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._flush_pending_newlines(
                self._improve_insert_doc,
                self._improve_insert_cursor,
                "_improve_pending_newlines",
            )
            self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._improve_insert_doc,
                self._improve_insert_cursor,
            )
        self._commit_pending_regenerate_context()
        self._capture_improve1_range_end()
        return False

    def _on_shorten_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._improve_insert_doc and self._improve_insert_cursor:
            self._flush_pending_newlines(
                self._improve_insert_doc,
                self._improve_insert_cursor,
                "_improve_pending_newlines",
            )
            self._ensure_single_trailing_space(self._improve_insert_doc, self._improve_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._improve_insert_doc,
                self._improve_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_improve1_range_end()
        return False

    def _on_topic_sentence_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_spellingstyle_range_end()
        return False

    def _on_introduction_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_spellingstyle_range_end()
        return False

    def _on_conclusion_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_spellingstyle_range_end()
        return False

    def _on_conclusion_no_issues_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_spellingstyle_range_end()
        return False

    def _on_concl_section_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        if self._editor_insert_doc and self._editor_insert_cursor:
            self._flush_pending_newlines(
                self._editor_insert_doc,
                self._editor_insert_cursor,
                "_editor_pending_newlines",
            )
            self._ensure_single_trailing_space(self._editor_insert_doc, self._editor_insert_cursor)
            self._apply_case_citation_italics_to_recent_writer_insert(
                self._editor_insert_doc,
                self._editor_insert_cursor,
            )
        self._trim_spelling_output_edges()
        self._commit_pending_regenerate_context()
        self._capture_spellingstyle_range_end()
        return False

    def _on_translate_finished(self, message: str) -> bool:
        self._set_busy(False)
        self._status_label.set_label(message)
        return False

    def _render_suggestions(self) -> None:
        for row in self._suggestion_rows:
            self._list_box.remove(row)
        self._suggestion_rows.clear()
        if not self._suggestions:
            self._empty_label.set_visible(True)
            return
        self._empty_label.set_visible(False)
        for idx, suggestion in enumerate(self._suggestions):
            row = self._build_suggestion_row(idx, suggestion)
            self._list_box.append(row)
            self._suggestion_rows.append(row)
        self._list_box.set_visible(True)

    def _build_suggestion_row(self, idx: int, suggestion: Suggestion) -> Gtk.Box:
        card = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        card.set_margin_start(6)
        card.set_margin_end(6)
        card.set_margin_top(6)
        card.set_margin_bottom(6)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        content.set_margin_start(8)
        content.set_margin_end(8)
        content.set_margin_top(8)
        content.set_margin_bottom(8)
        label = Gtk.Label(label=suggestion.title, wrap=True, xalign=0)
        label.add_css_class("title-3")
        content.append(label)

        reason = Gtk.Label(label=suggestion.reasoning or "", wrap=True, xalign=0)
        reason.add_css_class("dim-label")
        content.append(reason)

        snippet = self._build_suggestion_original_view(suggestion)
        content.append(snippet)

        replacement = self._build_suggestion_replacement_view(suggestion)
        content.append(replacement)

        btn_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        goto = Gtk.Button(label="Go To", icon_name="go-jump-symbolic")
        goto.add_css_class("flat")
        goto.set_label("Go To")
        goto.set_action_name("app.suggestion-goto")
        goto.set_action_target_value(GLib.Variant("i", idx))

        btn_box.append(goto)
        accept = Gtk.Button(label="Accept", icon_name="emblem-default-symbolic")
        accept.add_css_class("suggested-action")
        accept.add_css_class("flat")
        accept.set_label("Accept")
        accept.set_action_name("app.suggestion-accept")
        accept.set_action_target_value(GLib.Variant("i", idx))

        reject = Gtk.Button(label="Reject", icon_name="edit-delete-symbolic")
        reject.add_css_class("flat")
        reject.set_label("Reject")
        reject.set_action_name("app.suggestion-reject")
        reject.set_action_target_value(GLib.Variant("i", idx))

        btn_box.append(accept)
        btn_box.append(reject)
        content.append(btn_box)

        card.append(content)
        card.add_css_class("card")
        return card

    def _build_suggestion_diff_row(
        self,
        marker_text: str,
        marker_css_class: str,
        child: Gtk.Widget,
        tooltip: str,
    ) -> Gtk.Box:
        row = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        row.add_css_class("proof-suggestion-diff-row")
        row.set_hexpand(True)

        marker = Gtk.Label(label=marker_text, xalign=0.5)
        marker.add_css_class("monospace")
        marker.add_css_class("proof-suggestion-diff-marker")
        marker.add_css_class(marker_css_class)
        marker.set_size_request(24, -1)
        marker.set_tooltip_text(tooltip)
        marker.set_valign(Gtk.Align.START)

        child.set_hexpand(True)
        child.set_valign(Gtk.Align.START)
        row.append(marker)
        row.append(child)
        return row

    def _build_suggestion_original_view(self, suggestion: Suggestion) -> Gtk.Box:
        snippet = Gtk.Label(
            label=suggestion.snippet,
            wrap=True,
            xalign=0,
        )
        snippet.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        snippet.set_selectable(True)
        snippet.add_css_class("monospace")
        snippet.add_css_class("proof-suggestion-original-text")
        return self._build_suggestion_diff_row(
            "-",
            "proof-suggestion-diff-marker-minus",
            snippet,
            "Original text",
        )

    def _build_suggestion_replacement_view(self, suggestion: Suggestion) -> Gtk.Box:
        replacement = Gtk.Label(label=suggestion.replacement, wrap=True, xalign=0)
        replacement.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        replacement.set_selectable(True)
        replacement.add_css_class("proof-suggestion-replacement")
        replacement.add_css_class("monospace")
        attributes = Pango.AttrList()
        for start_offset, end_offset in self._changed_text_word_ranges(
            suggestion.snippet,
            suggestion.replacement,
        ):
            if start_offset >= end_offset:
                continue
            byte_start = len(suggestion.replacement[:start_offset].encode("utf-8"))
            byte_end = len(suggestion.replacement[:end_offset].encode("utf-8"))
            for attr in (
                Pango.attr_weight_new(Pango.Weight.BOLD),
                Pango.attr_foreground_new(38 * 257, 162 * 257, 105 * 257),
            ):
                attr.start_index = byte_start
                attr.end_index = byte_end
                attributes.insert(attr)
        replacement.set_attributes(attributes)
        return self._build_suggestion_diff_row(
            "+",
            "proof-suggestion-diff-marker-plus",
            replacement,
            "Suggested replacement",
        )

    # Suggestion actions --------------------------------------------------
    def _on_accept_clicked(self, _button: Gtk.Button, idx: int) -> None:
        suggestion = self._get_suggestion(idx)
        if not suggestion:
            return
        desktop = self._get_desktop()
        if not desktop:
            self._show_toast("LibreOffice is not reachable.")
            return
        doc = self._get_active_writer(desktop)
        if not doc:
            self._show_toast("Open a Writer document first.")
            return
        found = self._find_range(doc, suggestion.snippet)
        if not found:
            self._show_toast("Snippet not found in document.")
            return
        try:
            replacement_range = self._replace_writer_range_text(doc, found, suggestion.replacement)
            self._apply_case_citation_italics_to_range(doc, replacement_range, suggestion.replacement)
        except Exception:
            try:
                found.String = suggestion.replacement  # type: ignore[attr-defined]
            except Exception as exc:  # noqa: BLE001
                self._show_toast(f"Unable to apply change: {exc}")
                return
        self._save_document(doc)
        self._remove_suggestion(idx)
        self._show_toast("Change applied.")
        self._focus_first_suggestion(notify_success=False)

    def _on_reject_clicked(self, _button: Gtk.Button, idx: int) -> None:
        self._remove_suggestion(idx)
        self._focus_first_suggestion(notify_success=False)

    def _on_goto_clicked(self, _button: Gtk.Button, idx: int) -> None:
        self._focus_suggestion(idx, notify_success=True)

    def _get_suggestion(self, idx: int) -> Suggestion | None:
        if idx < 0 or idx >= len(self._suggestions):
            return None
        return self._suggestions[idx]

    def _remove_suggestion(self, idx: int) -> None:
        if idx < 0 or idx >= len(self._suggestions):
            return
        self._suggestions.pop(idx)
        self._render_suggestions()

    def _save_document(self, doc: XTextDocument) -> None:  # type: ignore[type-arg]
        try:
            if hasattr(doc, "store"):
                doc.store()
        except Exception as exc:  # noqa: BLE001
            self._show_toast(f"Unable to save document: {exc}")

    def _focus_suggestion(
        self, idx: int, notify_success: bool = True, notify_failure: bool = True
    ) -> bool:
        suggestion = self._get_suggestion(idx)
        if not suggestion:
            return False
        desktop = self._get_desktop()
        if not desktop:
            if notify_failure:
                self._show_toast("LibreOffice is not reachable.")
            return False
        doc = self._get_active_writer(desktop)
        if not doc:
            if notify_failure:
                self._show_toast("Open a Writer document first.")
            return False
        found = self._find_range(doc, suggestion.snippet)
        if not found:
            if notify_failure:
                self._show_toast("Snippet not found in document.")
            return False
        try:
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            view_cursor.gotoRange(found, False)
            controller.select(found)
            if notify_success:
                self._show_toast("Moved to suggested change.")
            return True
        except Exception as exc:  # noqa: BLE001
            if notify_failure:
                self._show_toast(f"Unable to move cursor: {exc}")
            return False

    def _focus_first_suggestion(self, notify_success: bool = False) -> None:
        if not self._suggestions:
            return
        self._focus_suggestion(0, notify_success=notify_success, notify_failure=True)

    def _on_view_model_response_clicked(self, _button: Gtk.Button | None) -> None:
        if not self._last_model_response_raw:
            self._show_toast("No model response available yet.")
            return
        text = self._format_last_model_response()
        if self._model_response_window:
            try:
                self._model_response_window.destroy()
            except Exception:
                pass
        window = Adw.ApplicationWindow(application=self.get_application(), title="Model Response Audit")
        window.set_default_size(900, 720)
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.set_title_widget(Gtk.Label(label="Model Response Audit"))
        header.set_show_end_title_buttons(True)
        view.add_top_bar(header)
        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scroller.set_vexpand(True)
        scroller.set_hexpand(True)
        textview = Gtk.TextView()
        textview.set_monospace(True)
        textview.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        textview.set_editable(False)
        textview.set_cursor_visible(False)
        textview.get_buffer().set_text(text)
        scroller.set_child(textview)
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        box.set_margin_top(12)
        box.set_margin_bottom(12)
        box.set_margin_start(12)
        box.set_margin_end(12)
        title_label = Gtk.Label(label=self._last_model_response_title or "Latest Model Response", xalign=0)
        title_label.add_css_class("title-4")
        title_label.set_wrap(True)
        box.append(title_label)
        box.append(scroller)
        view.set_content(box)
        window.set_content(view)
        window.connect("close-request", self._on_model_response_window_closed)
        self._model_response_window = window
        window.present()

    def _format_last_model_response(self) -> str:
        raw = self._last_model_response_raw or ""
        try:
            loaded = json.loads(raw)
            formatted_response = json.dumps(loaded, indent=2, ensure_ascii=False)
        except Exception:
            formatted_response = raw
        lines = [
            f"Action: {self._last_model_response_title or 'Unknown'}",
            f"Timestamp: {self._last_model_response_timestamp or 'Unknown'}",
            f"API URL: {self._last_model_response_api_url or 'Unknown'}",
            "",
            "Reasoning diagnostics:",
            self._last_model_response_diagnostic or "not available",
            "",
            "Response:",
            formatted_response,
        ]
        return "\n".join(lines)

    def _on_model_response_window_closed(self, window: Adw.ApplicationWindow, *_args: object) -> bool:
        self._model_response_window = None
        window.destroy()
        return True

    def _on_view_editor_prompt_clicked(self, _button: Gtk.Button | None) -> None:
        if not self._last_editor_request_raw:
            self._show_toast("No Editor prompt recorded yet.")
            return
        text = self._format_last_editor_prompt()
        if self._editor_prompt_window:
            try:
                self._editor_prompt_window.destroy()
            except Exception:
                pass
        title = self._last_editor_request_title or "Latest LLM Prompt"
        window = Adw.ApplicationWindow(application=self.get_application(), title="Latest LLM Prompt")
        window.set_default_size(900, 720)
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.set_title_widget(Gtk.Label(label="Latest LLM Prompt"))
        header.set_show_end_title_buttons(True)
        view.add_top_bar(header)
        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scroller.set_vexpand(True)
        scroller.set_hexpand(True)
        textview = Gtk.TextView()
        textview.set_monospace(True)
        textview.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        textview.set_editable(False)
        textview.set_cursor_visible(False)
        textview.get_buffer().set_text(text)
        scroller.set_child(textview)
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        box.set_margin_top(12)
        box.set_margin_bottom(12)
        box.set_margin_start(12)
        box.set_margin_end(12)
        title_label = Gtk.Label(label=title, xalign=0)
        title_label.add_css_class("title-4")
        title_label.set_wrap(True)
        box.append(title_label)
        box.append(scroller)
        view.set_content(box)
        window.set_content(view)
        window.connect("close-request", self._on_editor_prompt_window_closed)
        self._editor_prompt_window = window
        window.present()

    def _format_last_editor_prompt(self) -> str:
        raw = self._last_editor_request_raw or ""
        try:
            formatted_payload = json.dumps(json.loads(raw), indent=2, ensure_ascii=False)
        except Exception:
            formatted_payload = raw
        lines = [
            f"Action: {self._last_editor_request_title or 'Unknown'}",
            f"Timestamp: {self._last_editor_request_timestamp or 'Unknown'}",
            f"API URL: {self._last_editor_request_api_url or 'Unknown'}",
            "",
            "Payload:",
            formatted_payload,
        ]
        return "\n".join(lines)

    def _on_editor_prompt_window_closed(self, window: Adw.ApplicationWindow, *_args: object) -> bool:
        self._editor_prompt_window = None
        window.destroy()
        return True

    # Helpers -------------------------------------------------------------
    def _start_listener(self) -> None:
        if self._listener_proc and self._listener_proc.poll() is None:
            return
        try:
            LIBREOFFICE_PROFILE.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        soffice_bin = SOFFICE_PATH if SOFFICE_PATH.exists() else None
        cmd = [
            str(soffice_bin) if soffice_bin else "libreoffice",
            "--writer",
            "--nodefault",  # avoid opening a blank document on launch
            "--nologo",
            "--nofirststartwizard",
            f"-env:UserInstallation=file://{LIBREOFFICE_PROFILE.as_posix()}",
            '--accept=socket,host=127.0.0.1,port=2004;urp;',
        ]
        try:
            self._listener_proc = subprocess.Popen(cmd)
            self._status_label.set_label("Starting dedicated LibreOffice listener…")
        except Exception as exc:  # noqa: BLE001
            self._show_toast(f"Unable to start LibreOffice: {exc}")

    def _connect_desktop(self, retries: int = 5, delay: float = 0.6) -> bool:
        try:
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )
        except Exception:
            return False
        for _ in range(retries):
            try:
                ctx = resolver.resolve(UNO_BRIDGE_URL)
                self._ctx = ctx
                self._desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
                self._status_label.set_label("LibreOffice UNO connected.")
                return True
            except NoConnectException:
                time.sleep(delay)
            except Exception:
                break
        return False

    def _show_toast(self, message: str) -> None:
        toast = Adw.Toast.new(message)
        if hasattr(self, "_overlay") and self._overlay:
            self._overlay.add_toast(toast)

    def _extract_http_error_summary(self, message: str) -> str:
        cleaned = str(message or "").strip()
        if not cleaned.lower().startswith("http "):
            return ""
        detail = cleaned
        separator = " - "
        if separator in cleaned:
            detail = cleaned.split(separator, 1)[1].strip()
        if not detail:
            return ""
        try:
            data = json.loads(detail)
        except json.JSONDecodeError:
            data = None
        if isinstance(data, dict):
            candidates: list[str] = []
            error_obj = data.get("error")
            if isinstance(error_obj, dict):
                for key in ("message", "detail", "type", "code"):
                    value = error_obj.get(key)
                    if isinstance(value, str) and value.strip():
                        candidates.append(value.strip())
            for key in ("message", "detail", "error"):
                value = data.get(key)
                if isinstance(value, str) and value.strip():
                    candidates.append(value.strip())
            if candidates:
                detail = ". ".join(dict.fromkeys(candidates))
        detail = re.sub(r"\s+", " ", detail).strip().strip(".")
        if not detail:
            return ""
        if len(detail) > 160:
            detail = f"{detail[:157].rstrip()}..."
        return detail

    def _llm_toast_message(self, message: str) -> str:
        cleaned = str(message or "").strip()
        if not cleaned:
            return "Model error. Try again."
        lowered = cleaned.lower()
        if "503" in lowered or "over capacity" in lowered:
            return "Model is busy. Try again."
        if lowered.startswith("http "):
            summary = self._extract_http_error_summary(cleaned)
            if summary:
                return summary
            return "Model error. Try again."
        if lowered.startswith("llm error"):
            return "Model error. Try again."
        if "empty output" in lowered:
            return "Model returned empty output."
        if "no suggestions" in lowered:
            return "No suggestions returned."
        if "not valid json" in lowered or "did not include content" in lowered:
            return "Model returned unreadable output."
        if len(cleaned) > 80:
            return "Model error. Try again."
        return cleaned

    def _notify_llm_error(self, message: str) -> None:
        self._status_label.set_label("Ready.")
        self._show_toast(self._llm_toast_message(message))

    def _update_uno_status(self) -> None:
        if uno is None:
            self._status_label.set_label(_uno_status_message(self._libreoffice_python_path))
            return
        self._get_desktop()


class SettingsWindow(Adw.ApplicationWindow):
    def __init__(
        self,
        parent: ProseWindow,
        model_profiles: list[ModelProfile],
        proof_settings: ProofreadSettings,
        spelling_settings: SpellingStyleSettings,
        text_draft_spelling_prompt: str,
        citation_validator_settings: CitationValidatorSettings,
        improve1_settings: Improve1Settings,
        improve2_settings: Improve2Settings,
        combine_cites_settings: CombineCitesSettings,
        thesaurus_settings: ThesaurusSettings,
        reference_settings: ReferenceSettings,
        ask_settings: AskSettings,
        shorten_settings: ShortenSettings,
        introduction_settings: IntroductionSettings,
        introduction_reply_settings: IntroductionSettings,
        conclusion_settings: ConclusionSettings,
        concl_no_issues_settings: ConclNoIssuesSettings,
        topic_sentence_settings: TopicSentenceSettings,
        concl_section_settings: ConclSectionSettings,
        translate_settings: TranslateSettings,
        shared_style_rules: str,
        editor_action_profile_defaults: dict[str, str | None],
        editor_pinned_action_ids: list[str],
        text_draft_pinned_action_ids: list[str],
        text_draft_template_dir: Path | None,
        text_draft_external_actions: list[TextDraftExternalAction],
        libreoffice_python_path: Path | None,
        concordance_file_path: Path | None,
        editor_source_file: Path | None,
        on_source_change: Callable[[Path | None], None],
        on_copy_normal_profile: Callable[[], tuple[bool, str]],
        on_save: Callable[
            [
                list[ModelProfile],
                ProofreadSettings,
                SpellingStyleSettings,
                str,
                CitationValidatorSettings,
                Improve1Settings,
                Improve2Settings,
                CombineCitesSettings,
                ThesaurusSettings,
                ReferenceSettings,
                AskSettings,
                ShortenSettings,
                IntroductionSettings,
                IntroductionSettings,
                ConclusionSettings,
                ConclNoIssuesSettings,
                TopicSentenceSettings,
                ConclSectionSettings,
                TranslateSettings,
                str,
                dict[str, str | None],
                list[str],
                list[str],
                Path | None,
                list[TextDraftExternalAction],
                Path | None,
                Path | None,
            ],
            None,
        ],
    ) -> None:
        super().__init__(application=parent.get_application(), title="Settings")
        self._parent_window = parent
        self._on_save = on_save
        self._on_source_change = on_source_change
        self._on_copy_normal_profile = on_copy_normal_profile
        self._model_profiles = model_profiles
        self._proof_settings = proof_settings
        self._spelling_settings = spelling_settings
        self._text_draft_spelling_prompt = text_draft_spelling_prompt.strip() or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        self._citation_validator_settings = citation_validator_settings
        self._improve1_settings = improve1_settings
        self._improve2_settings = improve2_settings
        self._combine_cites_settings = combine_cites_settings
        self._thesaurus_settings = thesaurus_settings
        self._reference_settings = reference_settings
        self._ask_settings = ask_settings
        self._shorten_settings = shorten_settings
        self._introduction_settings = introduction_settings
        self._introduction_reply_settings = introduction_reply_settings
        self._conclusion_settings = conclusion_settings
        self._concl_no_issues_settings = concl_no_issues_settings
        self._topic_sentence_settings = topic_sentence_settings
        self._concl_section_settings = concl_section_settings
        self._translate_settings = translate_settings
        self._shared_style_rules_text = str(shared_style_rules or "").strip()
        self._editor_action_profile_defaults = _sanitize_editor_action_profile_defaults(editor_action_profile_defaults)
        self._editor_action_order = _ordered_editor_quick_action_keys(editor_pinned_action_ids)
        pinned_action_set = set(_sanitize_editor_pinned_actions(editor_pinned_action_ids))
        self._quick_action_enabled = {
            key: key in pinned_action_set
            for key in self._editor_action_order
        }
        self._text_draft_action_order = _ordered_text_draft_quick_action_keys(text_draft_pinned_action_ids)
        text_draft_pinned_action_set = set(_sanitize_text_draft_pinned_actions(text_draft_pinned_action_ids))
        self._text_draft_quick_action_enabled = {
            key: key in text_draft_pinned_action_set
            for key in self._text_draft_action_order
        }
        self._text_draft_template_dir = (
            text_draft_template_dir.expanduser().resolve(strict=False) if text_draft_template_dir else None
        )
        self._text_draft_external_actions = list(text_draft_external_actions)
        self._libreoffice_python_path = libreoffice_python_path
        self._concordance_file_path = concordance_file_path
        self._editor_source_file = editor_source_file
        self._proof_suggestions_json_path = proof_settings.suggestions_json_path
        self._text_draft_spelling_prompt_buffer: Gtk.TextBuffer | None = None
        self._shared_style_rules_buffer: Gtk.TextBuffer | None = None
        self._model_profile_editors: dict[str, ModelProfileEditorWidgets] = {}
        self._prompt_editors: dict[str, PromptEditorWidgets] = {}
        self._choices_profile_dropdowns: dict[str, Gtk.DropDown] = {}
        self._prompt_row_keys: dict[Gtk.ListBoxRow, str] = {}
        self._source_row_guard = False
        self._text_draft_template_dir_row_guard = False
        self._text_draft_external_actions_box: Gtk.Box | None = None
        self._text_draft_external_action_widgets: list[TextDraftExternalActionEditorWidgets] = []
        self._libreoffice_path_row_guard = False
        self._concordance_path_row_guard = False
        self._proof_suggestions_json_path_row_guard = False
        self._quick_action_toggle_guard = False
        self._text_draft_quick_action_toggle_guard = False
        self.set_default_size(900, 720)
        self.set_resizable(True)
        self._build_ui()

    def _build_ui(self) -> None:
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.add_css_class("flat")
        header.set_title_widget(Gtk.Label(label="Settings", xalign=0))
        view.add_top_bar(header)

        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        box.set_margin_top(18)
        box.set_margin_bottom(12)
        box.set_margin_start(18)
        box.set_margin_end(18)

        split = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        split.set_hexpand(True)
        split.set_vexpand(False)

        prompt_list = Gtk.ListBox()
        prompt_list.set_selection_mode(Gtk.SelectionMode.SINGLE)
        prompt_list.add_css_class("navigation-sidebar")
        prompt_list.connect("row-selected", self._on_prompt_row_selected)

        prompt_list_scroller = Gtk.ScrolledWindow()
        prompt_list_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        prompt_list_scroller.set_min_content_width(220)
        prompt_list_scroller.set_size_request(220, 720)
        prompt_list_scroller.set_child(prompt_list)

        prompt_stack = Gtk.Stack()
        prompt_stack.set_hexpand(True)
        prompt_stack.set_vexpand(True)
        prompt_stack.set_transition_type(Gtk.StackTransitionType.SLIDE_LEFT_RIGHT)
        self._prompt_stack = prompt_stack
        self._prompt_list = prompt_list

        first_row: Gtk.ListBoxRow | None = None

        source_row = Gtk.ListBoxRow()
        source_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        source_row_box.set_margin_top(8)
        source_row_box.set_margin_bottom(8)
        source_row_box.set_margin_start(12)
        source_row_box.set_margin_end(12)
        source_row_box.append(Gtk.Label(label="Source Text", xalign=0))
        source_row.set_child(source_row_box)
        prompt_list.append(source_row)
        self._prompt_row_keys[source_row] = "source-text"
        prompt_stack.add_named(self._build_source_text_page(), "source-text")
        first_row = source_row

        external_action_row = Gtk.ListBoxRow()
        external_action_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        external_action_row_box.set_margin_top(8)
        external_action_row_box.set_margin_bottom(8)
        external_action_row_box.set_margin_start(12)
        external_action_row_box.set_margin_end(12)
        external_action_row_box.append(Gtk.Label(label="Text Draft External Action", xalign=0))
        external_action_row.set_child(external_action_row_box)
        prompt_list.append(external_action_row)
        self._prompt_row_keys[external_action_row] = "text-draft-external-action"
        prompt_stack.add_named(self._build_text_draft_external_action_page(), "text-draft-external-action")

        profiles_row = Gtk.ListBoxRow()
        profiles_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        profiles_row_box.set_margin_top(8)
        profiles_row_box.set_margin_bottom(8)
        profiles_row_box.set_margin_start(12)
        profiles_row_box.set_margin_end(12)
        profiles_row_box.append(Gtk.Label(label="Model Profiles", xalign=0))
        profiles_row.set_child(profiles_row_box)
        prompt_list.append(profiles_row)
        self._prompt_row_keys[profiles_row] = "model-profiles"
        prompt_stack.add_named(self._build_model_profiles_page(), "model-profiles")

        choices_models_row = Gtk.ListBoxRow()
        choices_models_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        choices_models_row_box.set_margin_top(8)
        choices_models_row_box.set_margin_bottom(8)
        choices_models_row_box.set_margin_start(12)
        choices_models_row_box.set_margin_end(12)
        choices_models_row_box.append(Gtk.Label(label="Choices Models", xalign=0))
        choices_models_row.set_child(choices_models_row_box)
        prompt_list.append(choices_models_row)
        self._prompt_row_keys[choices_models_row] = "choices-models"
        prompt_stack.add_named(self._build_choices_models_page(), "choices-models")

        style_rules_row = Gtk.ListBoxRow()
        style_rules_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        style_rules_row_box.set_margin_top(8)
        style_rules_row_box.set_margin_bottom(8)
        style_rules_row_box.set_margin_start(12)
        style_rules_row_box.set_margin_end(12)
        style_rules_row_box.append(Gtk.Label(label="Style Rules", xalign=0))
        style_rules_row.set_child(style_rules_row_box)
        prompt_list.append(style_rules_row)
        self._prompt_row_keys[style_rules_row] = "style-rules"
        prompt_stack.add_named(self._build_style_rules_page(), "style-rules")

        prompt_definitions = [
            ("proof", "Proof Reading", self._proof_settings, DEFAULT_PROMPT),
            ("spelling", "SpellingStyle", self._spelling_settings, DEFAULT_SPELLINGSTYLE_PROMPT),
            (
                "citation-validator",
                "Citation Number Validator",
                self._citation_validator_settings,
                DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT,
            ),
            ("thesaurus", "Thesaurus", self._thesaurus_settings, DEFAULT_THESAURUS_PROMPT),
            ("reference", "Reference", self._reference_settings, DEFAULT_REFERENCE_PROMPT),
            ("ask", "Ask Field", self._ask_settings, DEFAULT_ASK_PROMPT),
            ("improve-generated", "Improve Generated", self._improve1_settings, DEFAULT_IMPROVE_PROMPT),
            ("rephrase-generated", "Rephrase Generated", self._improve2_settings, DEFAULT_IMPROVE2_PROMPT),
            ("combine", "Combine Cites", self._combine_cites_settings, DEFAULT_COMBINE_CITES_PROMPT),
            ("shorten", "Shorten Selected", self._shorten_settings, DEFAULT_SHORTEN_PROMPT),
            ("intro", "Introduction", self._introduction_settings, DEFAULT_INTRO_PROMPT),
            ("intro-reply", "Introduction Reply", self._introduction_reply_settings, DEFAULT_INTRO_REPLY_PROMPT),
            ("conclusion", "Conclusion", self._conclusion_settings, DEFAULT_CONCLUSION_PROMPT),
            ("concl-no-issues", "Conclusion No Issues", self._concl_no_issues_settings, DEFAULT_CONCL_NO_ISSUES_PROMPT),
            ("topic-sentence", "Topic Sentence", self._topic_sentence_settings, DEFAULT_TOPIC_SENTENCE_PROMPT),
            ("concl-section", "Section Conclusion", self._concl_section_settings, DEFAULT_CONCL_SECTION_PROMPT),
            ("translate", "Translate", self._translate_settings, DEFAULT_TRANSLATE_PROMPT),
        ]
        for key, title, settings, default_prompt in prompt_definitions:
            row = Gtk.ListBoxRow()
            row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
            row_box.set_margin_top(8)
            row_box.set_margin_bottom(8)
            row_box.set_margin_start(12)
            row_box.set_margin_end(12)
            label = Gtk.Label(label=title, xalign=0)
            row_box.append(label)
            row.set_child(row_box)
            prompt_list.append(row)
            self._prompt_row_keys[row] = key
            if first_row is None:
                first_row = row

            page = self._build_prompt_page(key, title, settings, default_prompt)
            prompt_stack.add_named(page, key)

        text_draft_spelling_row = Gtk.ListBoxRow()
        text_draft_spelling_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_spelling_row_box.set_margin_top(8)
        text_draft_spelling_row_box.set_margin_bottom(8)
        text_draft_spelling_row_box.set_margin_start(12)
        text_draft_spelling_row_box.set_margin_end(12)
        text_draft_spelling_row_box.append(Gtk.Label(label="Text Draft SpellingStyle", xalign=0))
        text_draft_spelling_row.set_child(text_draft_spelling_row_box)
        prompt_list.append(text_draft_spelling_row)
        self._prompt_row_keys[text_draft_spelling_row] = "text-draft-spelling"
        prompt_stack.add_named(self._build_text_draft_spellingstyle_page(), "text-draft-spelling")

        quick_actions_row = Gtk.ListBoxRow()
        quick_actions_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        quick_actions_row_box.set_margin_top(8)
        quick_actions_row_box.set_margin_bottom(8)
        quick_actions_row_box.set_margin_start(12)
        quick_actions_row_box.set_margin_end(12)
        quick_actions_row_box.append(Gtk.Label(label="Quick Actions", xalign=0))
        quick_actions_row.set_child(quick_actions_row_box)
        prompt_list.append(quick_actions_row)
        self._prompt_row_keys[quick_actions_row] = "quick-actions"
        prompt_stack.add_named(self._build_quick_actions_page(), "quick-actions")

        text_draft_quick_actions_row = Gtk.ListBoxRow()
        text_draft_quick_actions_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_quick_actions_row_box.set_margin_top(8)
        text_draft_quick_actions_row_box.set_margin_bottom(8)
        text_draft_quick_actions_row_box.set_margin_start(12)
        text_draft_quick_actions_row_box.set_margin_end(12)
        text_draft_quick_actions_row_box.append(Gtk.Label(label="Text Draft Quick Actions", xalign=0))
        text_draft_quick_actions_row.set_child(text_draft_quick_actions_row_box)
        prompt_list.append(text_draft_quick_actions_row)
        self._prompt_row_keys[text_draft_quick_actions_row] = "text-draft-quick-actions"
        prompt_stack.add_named(self._build_text_draft_quick_actions_page(), "text-draft-quick-actions")

        editor_commands_row = Gtk.ListBoxRow()
        editor_commands_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        editor_commands_row_box.set_margin_top(8)
        editor_commands_row_box.set_margin_bottom(8)
        editor_commands_row_box.set_margin_start(12)
        editor_commands_row_box.set_margin_end(12)
        editor_commands_row_box.append(Gtk.Label(label="Editor Commands", xalign=0))
        editor_commands_row.set_child(editor_commands_row_box)
        prompt_list.append(editor_commands_row)
        self._prompt_row_keys[editor_commands_row] = "editor-commands"
        prompt_stack.add_named(self._build_editor_commands_page(), "editor-commands")

        text_draft_commands_row = Gtk.ListBoxRow()
        text_draft_commands_row_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        text_draft_commands_row_box.set_margin_top(8)
        text_draft_commands_row_box.set_margin_bottom(8)
        text_draft_commands_row_box.set_margin_start(12)
        text_draft_commands_row_box.set_margin_end(12)
        text_draft_commands_row_box.append(Gtk.Label(label="Text Draft Commands", xalign=0))
        text_draft_commands_row.set_child(text_draft_commands_row_box)
        prompt_list.append(text_draft_commands_row)
        self._prompt_row_keys[text_draft_commands_row] = "text-draft-commands"
        prompt_stack.add_named(self._build_text_draft_commands_page(), "text-draft-commands")

        if first_row is not None:
            prompt_list.select_row(first_row)
            prompt_stack.set_visible_child_name(self._prompt_row_keys[first_row])

        split.append(prompt_list_scroller)
        split.append(prompt_stack)
        box.append(split)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        scrolled = Gtk.ScrolledWindow()
        scrolled.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scrolled.set_hexpand(True)
        scrolled.set_vexpand(True)
        scrolled.set_child(box)
        content.append(scrolled)

        buttons = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        buttons.set_margin_top(6)
        buttons.set_margin_bottom(12)
        buttons.set_margin_start(12)
        buttons.set_margin_end(12)
        buttons.set_halign(Gtk.Align.END)
        close_btn = Gtk.Button(label="Close")
        close_btn.add_css_class("flat")
        close_btn.connect("clicked", self._on_close_clicked)
        buttons.append(close_btn)
        save_btn = Gtk.Button(label="Save Settings")
        save_btn.add_css_class("suggested-action")
        save_btn.add_css_class("flat")
        save_btn.set_action_name("app.save-settings")
        buttons.append(save_btn)
        content.append(buttons)

        view.set_content(content)
        self.set_content(view)

    def set_source_file(self, path: Path | None) -> None:
        self._set_editor_source_file(path, notify=False)

    def _build_source_text_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Source Text", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        source_group = Adw.PreferencesGroup(title="Source Text")
        source_group.add_css_class("list-stack")
        source_group.set_hexpand(True)
        page_box.append(source_group)

        source_row = Adw.EntryRow(title="Source text file")
        source_row.set_hexpand(True)
        if self._editor_source_file:
            source_row.set_text(str(self._editor_source_file))
        source_row.connect("changed", self._on_source_row_changed)
        choose_btn = Gtk.Button(label="Choose")
        choose_btn.add_css_class("flat")
        choose_btn.connect("clicked", self._on_choose_source_file)
        source_row.add_suffix(choose_btn)
        source_group.add(source_row)
        self._source_row = source_row

        text_draft_template_row, text_draft_template_entry = self._build_path_setting_row(
            title="Text Draft template directory",
            value=str(self._text_draft_template_dir or ""),
            info_text=(
                "Optional. Prose treats immediate subfolders here as template categories, and root-level .txt files "
                "appear under Uncategorized. "
                "Add leading lines like [[email: Label <address@example.com>]] to create copy buttons."
            ),
            on_changed=self._on_text_draft_template_dir_row_changed,
            on_choose=self._on_choose_text_draft_template_dir,
            on_clear=self._on_clear_text_draft_template_dir,
        )
        source_group.add(text_draft_template_row)
        self._text_draft_template_dir_row = text_draft_template_entry

        proof_suggestions_row, proof_suggestions_entry = self._build_path_setting_row(
            title="Proofreading suggestions JSON",
            value=str(self._proof_suggestions_json_path or ""),
            info_text=(
                "Optional. When this .json file exists, Prose auto-loads it when the Proof Reader tab is shown. "
                "Clear this path to disable auto-load."
            ),
            on_changed=self._on_proof_suggestions_json_path_row_changed,
            on_choose=self._on_choose_proof_suggestions_json_path,
            on_clear=self._on_clear_proof_suggestions_json_path,
        )
        source_group.add(proof_suggestions_row)
        self._proof_suggestions_json_path_row = proof_suggestions_entry

        libreoffice_row, libreoffice_entry = self._build_path_setting_row(
            title="LibreOffice Python path",
            value=str(self._libreoffice_python_path or ""),
            info_text="Required. Point this at LibreOffice's Python bridge directory, usually .../libreoffice/program.",
            on_changed=self._on_libreoffice_path_row_changed,
            on_choose=self._on_choose_libreoffice_path,
            on_clear=self._on_clear_libreoffice_path,
        )
        source_group.add(libreoffice_row)
        self._libreoffice_path_row = libreoffice_entry

        concordance_row, concordance_entry = self._build_path_setting_row(
            title="Concordance file",
            value=str(self._concordance_file_path or ""),
            info_text=(
                "Optional. Used by Add Case when appending citations to the concordance file "
                "and by the Text Draft case picker."
            ),
            on_changed=self._on_concordance_path_row_changed,
            on_choose=self._on_choose_concordance_file,
            on_clear=self._on_clear_concordance_file,
        )
        source_group.add(concordance_row)
        self._concordance_path_row = concordance_entry

        profile_row = Adw.ActionRow(
            title="LibreOffice Profile Import",
            subtitle=(
                f"Copy {_normal_libreoffice_profile_path()} into {LIBREOFFICE_PROFILE}.\n"
                "Overwrites the Prose LibreOffice profile after confirmation. Close any Prose Writer window first."
            ),
        )
        profile_row.set_subtitle_lines(4)
        profile_copy_btn = Gtk.Button(label="Copy Normal Profile to Prose")
        profile_copy_btn.add_css_class("flat")
        profile_copy_btn.connect("clicked", self._on_copy_normal_profile_clicked)
        profile_row.add_suffix(profile_copy_btn)
        source_group.add(profile_row)

        return page_box

    def _build_text_draft_external_action_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft External Action", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        actions_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        actions_box.set_hexpand(True)
        page_box.append(actions_box)
        self._text_draft_external_actions_box = actions_box
        self._rebuild_text_draft_external_action_editor_rows()

        add_button = Gtk.Button(label="Add Action", icon_name="list-add-symbolic")
        add_button.add_css_class("flat")
        add_button.add_css_class("suggested-action")
        add_button.set_halign(Gtk.Align.START)
        add_button.connect("clicked", self._on_add_text_draft_external_action_clicked)
        page_box.append(add_button)

        return page_box

    def _clear_settings_box(self, box: Gtk.Box) -> None:
        child = box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            box.remove(child)
            child = next_child

    def _rebuild_text_draft_external_action_editor_rows(self) -> None:
        box = self._text_draft_external_actions_box
        if box is None:
            return
        self._clear_settings_box(box)
        self._text_draft_external_action_widgets = []
        if not self._text_draft_external_actions:
            empty_label = Gtk.Label(label="No external Text Draft actions are configured.", xalign=0)
            empty_label.add_css_class("dim-label")
            box.append(empty_label)
            return

        for index, action in enumerate(self._text_draft_external_actions):
            group = Adw.PreferencesGroup(title=f"Action {index + 1}")
            group.add_css_class("list-stack")
            group.set_hexpand(True)
            box.append(group)

            enabled_row = Adw.SwitchRow(
                title="Enable action",
                subtitle="Show this action as a Text Draft button.",
            )
            enabled_row.set_active(action.enabled)
            group.add(enabled_row)

            label_row = Adw.EntryRow(title="Button label")
            label_row.set_text(action.label or "External")
            group.add(label_row)

            icon_row = Adw.EntryRow(title="Icon name")
            icon_row.set_text(action.icon_name or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON)
            group.add(icon_row)

            tooltip_row = Adw.EntryRow(title="Tooltip")
            tooltip_row.set_text(action.tooltip)
            group.add(tooltip_row)

            success_message_row = Adw.EntryRow(title="Success message")
            success_message_row.set_text(action.success_message)
            group.add(success_message_row)

            command_row = Adw.EntryRow(title="Command")
            command_row.set_text(shlex.join(action.command))
            command_row.set_show_apply_button(False)
            command_row.set_tooltip_text(
                f"Use {TEXT_DRAFT_EXTERNAL_ACTION_DRAFT_FILE_TOKEN} where the Draft temp-file path should go."
            )
            group.add(command_row)

            cwd_row = Adw.EntryRow(title="Working directory")
            cwd_row.set_text(str(action.cwd or ""))
            cwd_row.set_show_apply_button(False)
            group.add(cwd_row)

            codex_reasoning_dropdown = None
            env_values = dict(action.env)
            if self._parent_window._is_text_draft_codex_action(action):
                env_values.pop("CODEX_REASONING_EFFORT", None)

                codex_reasoning_row = Adw.ActionRow(
                    title="Codex reasoning effort",
                    subtitle="Reasoning level used when this Codex action starts.",
                )
                codex_reasoning_row.set_activatable(False)
                codex_reasoning_dropdown = Gtk.DropDown(
                    model=Gtk.StringList.new(list(TEXT_DRAFT_CODEX_REASONING_EFFORTS))
                )
                codex_reasoning_dropdown.set_selected(
                    TEXT_DRAFT_CODEX_REASONING_EFFORTS.index(
                        _sanitize_text_draft_codex_reasoning_effort(action.codex_reasoning_effort)
                    )
                )
                codex_reasoning_row.add_suffix(codex_reasoning_dropdown)
                group.add(codex_reasoning_row)

            env_buffer = Gtk.TextBuffer()
            env_buffer.set_text(_format_text_draft_external_action_env(env_values))
            env_row = Adw.PreferencesRow()
            env_row.set_selectable(False)
            env_row.set_activatable(False)
            env_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
            env_box.set_margin_top(10)
            env_box.set_margin_bottom(10)
            env_box.set_margin_start(12)
            env_box.set_margin_end(12)

            env_title = Gtk.Label(label="Environment", xalign=0)
            env_title.add_css_class("heading")
            env_box.append(env_title)

            env_hint = Gtk.Label(
                label="Optional. Enter one KEY=value pair per line.",
                xalign=0,
            )
            env_hint.add_css_class("dim-label")
            env_hint.set_wrap(True)
            env_box.append(env_hint)

            env_view = Gtk.TextView.new_with_buffer(env_buffer)
            env_view.set_monospace(True)
            env_view.set_wrap_mode(Gtk.WrapMode.NONE)
            env_view.set_hexpand(True)
            env_view.set_top_margin(8)
            env_view.set_bottom_margin(8)
            env_view.set_left_margin(8)
            env_view.set_right_margin(8)

            env_scroller = Gtk.ScrolledWindow()
            env_scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
            env_scroller.set_hexpand(True)
            env_scroller.set_min_content_height(82)
            env_scroller.set_max_content_height(150)
            env_scroller.set_child(env_view)
            env_box.append(env_scroller)

            env_row.set_child(env_box)
            group.add(env_row)

            remove_row = Adw.ActionRow(title="Remove action")
            remove_button = Gtk.Button(icon_name="user-trash-symbolic")
            remove_button.add_css_class("flat")
            remove_button.add_css_class("destructive-action")
            remove_button.set_tooltip_text("Remove this external Text Draft action.")
            remove_button.connect(
                "clicked",
                lambda _button, action_index=index: self._on_remove_text_draft_external_action_clicked(
                    action_index,
                ),
            )
            remove_row.add_suffix(remove_button)
            group.add(remove_row)

            self._text_draft_external_action_widgets.append(
                TextDraftExternalActionEditorWidgets(
                    enabled_row=enabled_row,
                    label_row=label_row,
                    icon_row=icon_row,
                    tooltip_row=tooltip_row,
                    success_message_row=success_message_row,
                    command_row=command_row,
                    cwd_row=cwd_row,
                    codex_reasoning_dropdown=codex_reasoning_dropdown,
                    env_buffer=env_buffer,
                )
            )

    def _collect_text_draft_external_actions_from_settings(self) -> list[TextDraftExternalAction] | None:
        actions: list[TextDraftExternalAction] = []
        for index, widgets in enumerate(self._text_draft_external_action_widgets):
            command: list[str] = []
            command_text = widgets.command_row.get_text().strip()
            if command_text:
                try:
                    command = shlex.split(command_text)
                except ValueError as exc:
                    self._parent_window._show_toast(f"External action {index + 1} command is invalid: {exc}")
                    return None
            env, env_error = _parse_text_draft_external_action_env(self._prompt_text(widgets.env_buffer))
            if env_error is not None or env is None:
                self._parent_window._show_toast(f"External action {index + 1} environment {env_error}")
                return None
            codex_reasoning_effort = DEFAULT_TEXT_DRAFT_CODEX_REASONING_EFFORT
            if widgets.codex_reasoning_dropdown is not None:
                env.pop("CODEX_REASONING_EFFORT", None)
                selected_reasoning = int(widgets.codex_reasoning_dropdown.get_selected())
                if 0 <= selected_reasoning < len(TEXT_DRAFT_CODEX_REASONING_EFFORTS):
                    codex_reasoning_effort = TEXT_DRAFT_CODEX_REASONING_EFFORTS[selected_reasoning]
            cwd_text = widgets.cwd_row.get_text().strip()
            actions.append(
                TextDraftExternalAction(
                    enabled=widgets.enabled_row.get_active(),
                    label=widgets.label_row.get_text().strip() or "External",
                    command=command,
                    cwd=Path(cwd_text).expanduser().resolve(strict=False) if cwd_text else None,
                    env=env,
                    icon_name=widgets.icon_row.get_text().strip() or DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON,
                    tooltip=widgets.tooltip_row.get_text().strip(),
                    success_message=widgets.success_message_row.get_text().strip(),
                    codex_reasoning_effort=codex_reasoning_effort,
                )
            )
        return actions

    def _on_add_text_draft_external_action_clicked(self, _button: Gtk.Button) -> None:
        actions = self._collect_text_draft_external_actions_from_settings()
        if actions is None:
            return
        actions.append(
            TextDraftExternalAction(
                enabled=True,
                label="External",
                command=[],
                cwd=None,
                env={},
                icon_name=DEFAULT_TEXT_DRAFT_EXTERNAL_ACTION_ICON,
                tooltip="",
                success_message="",
            )
        )
        self._text_draft_external_actions = actions
        self._rebuild_text_draft_external_action_editor_rows()

    def _on_remove_text_draft_external_action_clicked(self, index: int) -> None:
        actions = self._collect_text_draft_external_actions_from_settings()
        if actions is None:
            return
        if 0 <= index < len(actions):
            actions.pop(index)
        self._text_draft_external_actions = actions
        self._rebuild_text_draft_external_action_editor_rows()

    def _profile_dropdown_model(self, include_unset: bool = False) -> Gtk.StringList:
        labels = [profile.display_name() for profile in self._model_profiles]
        if include_unset:
            labels = [UNSET_PROFILE_LABEL, *labels]
        return Gtk.StringList.new(labels)

    def _build_model_profiles_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Model Profiles", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Set up up to four shared model profiles here. Proof Reading now keeps its own credentials, "
                "while eligible editor actions reuse these shared profiles and keep only their own prompts."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        for profile in self._model_profiles:
            group = Adw.PreferencesGroup(title=profile.display_name())
            group.add_css_class("list-stack")

            nickname_row = Adw.EntryRow(title="Nickname")
            nickname_row.set_text(profile.display_name())
            group.add(nickname_row)

            abbreviation_row = Adw.EntryRow(title="Abbreviation (optional)")
            abbreviation_row.set_text(profile.abbreviation)
            group.add(abbreviation_row)

            api_url_row = Adw.EntryRow(title="API URL")
            api_url_row.set_text(profile.api_url)
            group.add(api_url_row)

            model_row = Adw.EntryRow(title="Model ID (optional)")
            model_row.set_text(profile.model_id)
            group.add(model_row)

            api_key_row = Adw.PasswordEntryRow(title="API Key")
            api_key_row.set_text(profile.api_key)
            group.add(api_key_row)

            disable_reasoning_row = Adw.SwitchRow(title="Disable reasoning")
            disable_reasoning_row.set_active(bool(profile.disable_reasoning))
            group.add(disable_reasoning_row)

            self._model_profile_editors[profile.key] = ModelProfileEditorWidgets(
                nickname_row=nickname_row,
                abbreviation_row=abbreviation_row,
                api_url_row=api_url_row,
                model_row=model_row,
                api_key_row=api_key_row,
                disable_reasoning_row=disable_reasoning_row,
            )
            page_box.append(group)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_choices_models_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Choices Models", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Generated Choices and Selected Choices use these same three model assignments. "
                "Improve uses the Improve Generated prompt; Rephrase 1 and Rephrase 2 use the Rephrase Generated prompt. "
                "Draft introduction and conclusion Choices, including Section Conclusion Choices, use the action prompt "
                "for all three assignments."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        group = Adw.PreferencesGroup(title="Choices Model Profiles")
        group.add_css_class("list-stack")
        group.set_hexpand(True)
        page_box.append(group)

        self._choices_profile_dropdowns = {}
        rows = (
            ("choices-improve", "Improve", "Default Choices model assignment."),
            ("choices-rephrase-1", "Rephrase 1", "First Choices variation model assignment."),
            ("choices-rephrase-2", "Rephrase 2", "Second Choices variation model assignment."),
        )
        for profile_key, title, subtitle in rows:
            row = Adw.ActionRow(title=title, subtitle=subtitle)
            row.set_activatable(False)
            dropdown = Gtk.DropDown(model=self._profile_dropdown_model(include_unset=True))
            selected_profile_key = self._editor_action_profile_defaults.get(profile_key)
            selected_index = (
                MODEL_PROFILE_IDS.index(selected_profile_key) + 1
                if selected_profile_key in MODEL_PROFILE_IDS
                else 0
            )
            dropdown.set_selected(selected_index)
            row.add_suffix(dropdown)
            group.add(row)
            self._choices_profile_dropdowns[profile_key] = dropdown

        return page_box

    def _build_quick_actions_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Quick Actions", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} actions for the Editor toolbar. "
                "Use the arrows to change their order."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        quick_actions_group = Adw.PreferencesGroup(title="Pinned Editor Actions")
        quick_actions_group.add_css_class("list-stack")
        quick_actions_group.set_hexpand(True)
        page_box.append(quick_actions_group)

        quick_actions_row = Adw.PreferencesRow()
        quick_actions_row.set_selectable(False)
        quick_actions_row.set_activatable(False)

        quick_actions_content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        quick_actions_content.set_margin_top(10)
        quick_actions_content.set_margin_bottom(10)
        quick_actions_content.set_margin_start(12)
        quick_actions_content.set_margin_end(12)

        quick_actions_hint = Gtk.Label(xalign=0)
        quick_actions_hint.add_css_class("caption")
        quick_actions_hint.add_css_class("dim-label")
        quick_actions_hint.set_wrap(True)
        quick_actions_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        quick_actions_content.append(quick_actions_hint)
        self._quick_actions_hint = quick_actions_hint

        quick_actions_list = Gtk.ListBox()
        quick_actions_list.set_selection_mode(Gtk.SelectionMode.NONE)
        quick_actions_list.add_css_class("boxed-list")
        quick_actions_content.append(quick_actions_list)
        self._quick_actions_list = quick_actions_list

        quick_actions_row.set_child(quick_actions_content)
        quick_actions_group.add(quick_actions_row)
        self._rebuild_quick_action_rows()
        return page_box

    def _build_text_draft_quick_actions_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft Quick Actions", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} actions for the Text Draft toolbar. "
                "Unpinned actions stay under More."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        quick_actions_group = Adw.PreferencesGroup(title="Pinned Text Draft Actions")
        quick_actions_group.add_css_class("list-stack")
        quick_actions_group.set_hexpand(True)
        page_box.append(quick_actions_group)

        quick_actions_row = Adw.PreferencesRow()
        quick_actions_row.set_selectable(False)
        quick_actions_row.set_activatable(False)

        quick_actions_content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        quick_actions_content.set_margin_top(10)
        quick_actions_content.set_margin_bottom(10)
        quick_actions_content.set_margin_start(12)
        quick_actions_content.set_margin_end(12)

        quick_actions_hint = Gtk.Label(xalign=0)
        quick_actions_hint.add_css_class("caption")
        quick_actions_hint.add_css_class("dim-label")
        quick_actions_hint.set_wrap(True)
        quick_actions_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        quick_actions_content.append(quick_actions_hint)
        self._text_draft_quick_actions_hint = quick_actions_hint

        quick_actions_list = Gtk.ListBox()
        quick_actions_list.set_selection_mode(Gtk.SelectionMode.NONE)
        quick_actions_list.add_css_class("boxed-list")
        quick_actions_content.append(quick_actions_list)
        self._text_draft_quick_actions_list = quick_actions_list

        quick_actions_row.set_child(quick_actions_content)
        quick_actions_group.add(quick_actions_row)
        self._rebuild_text_draft_quick_action_rows()
        return page_box

    def _build_editor_commands_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Editor Commands", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Use Run to trigger actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        actions_group = Adw.PreferencesGroup(title="Editor Actions")
        actions_group.add_css_class("list-stack")
        page_box.append(actions_group)

        for title, action_name, param, desc, supports_profiles in _editor_command_items():
            row = self._build_settings_editor_command_row(
                title,
                action_name,
                param,
                desc,
                supports_profiles=supports_profiles,
            )
            actions_group.add(row)

        return page_box

    def _build_text_draft_commands_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_valign(Gtk.Align.START)

        title_label = Gtk.Label(label="Text Draft Commands", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        info_label = Gtk.Label(
            label=(
                "Use Run to trigger Text Draft actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard for use in other programs."
            ),
            xalign=0,
        )
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(info_label)

        actions_group = Adw.PreferencesGroup(title="Text Draft Actions")
        actions_group.add_css_class("list-stack")
        page_box.append(actions_group)

        for title, action_name, param, desc, supports_profiles in _text_draft_command_items():
            row = self._build_settings_editor_command_row(
                title,
                action_name,
                param,
                desc,
                supports_profiles=supports_profiles,
            )
            actions_group.add(row)

        return page_box

    def _build_settings_editor_command_row(
        self,
        title: str,
        action_name: str,
        param: str | None,
        desc: str,
        supports_profiles: bool = False,
    ) -> Adw.ActionRow:
        row = Adw.ActionRow(title=title, subtitle=desc)
        row.set_activatable(False)

        suffix = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        profile_dropdown = None
        if supports_profiles:
            profile_dropdown = Gtk.DropDown(model=self._profile_dropdown_model())
            lookup_key = _profile_default_lookup_key_for_action_name(action_name)
            selected_profile_key = self._editor_action_profile_defaults.get(lookup_key) if lookup_key else None
            selected_index = MODEL_PROFILE_IDS.index(selected_profile_key) if selected_profile_key in MODEL_PROFILE_IDS else 0
            profile_dropdown.set_selected(selected_index)
            suffix.append(profile_dropdown)

        run_btn = Gtk.Button(label="Run")
        run_btn.add_css_class("suggested-action")
        run_btn.add_css_class("flat")
        run_btn.connect("clicked", self._on_settings_editor_command_run_clicked, action_name, profile_dropdown)
        suffix.append(run_btn)

        copy_btn = Gtk.Button(label="Copy Command")
        copy_btn.add_css_class("flat")
        copy_btn.add_css_class("link")
        copy_btn.connect("clicked", self._on_settings_editor_command_copy_clicked, action_name, param, profile_dropdown)
        suffix.append(copy_btn)

        row.add_suffix(suffix)
        return row

    def _on_settings_editor_command_run_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        app = self.get_application()
        if profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._model_profiles]
            if 0 <= selected < len(nicknames):
                app.activate_action(action_name, GLib.Variant("s", nicknames[selected]))
            return
        app.activate_action(action_name, None)

    def _on_settings_editor_command_copy_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        param: str | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        if profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._model_profiles]
            if 0 <= selected < len(nicknames):
                param = _format_action_param(GLib.Variant("s", nicknames[selected]))
        app = self.get_application()
        object_path = ACTION_OBJECT_PATH
        if isinstance(app, Gio.Application):
            app_path = app.get_dbus_object_path()
            if app_path:
                object_path = app_path
        command = _action_command(action_name, param, object_path)
        display = Gdk.Display.get_default()
        if display:
            display.get_clipboard().set(command)

    def _on_choose_source_file(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose a source text file")
        dialog.open(self, None, self._on_source_file_chosen)

    def _on_source_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_editor_source_file(path, notify=True)

    def _on_source_row_changed(self, row: Adw.EntryRow) -> None:
        if self._source_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_editor_source_file(None, notify=True)
            return
        self._set_editor_source_file(Path(raw), notify=True)

    def _build_path_setting_row(
        self,
        title: str,
        value: str,
        info_text: str,
        on_changed: Callable[[Gtk.Editable], None],
        on_choose: Callable[[Gtk.Button], None],
        on_clear: Callable[[Gtk.Button], None],
    ) -> tuple[Adw.PreferencesRow, Gtk.Entry]:
        row = Adw.PreferencesRow()
        row.set_selectable(False)
        row.set_activatable(False)

        content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
        content.set_margin_top(10)
        content.set_margin_bottom(10)
        content.set_margin_start(12)
        content.set_margin_end(12)

        title_label = Gtk.Label(label=title, xalign=0)
        title_label.add_css_class("caption")
        title_label.add_css_class("dim-label")
        content.append(title_label)

        entry_box = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        entry = Gtk.Entry()
        entry.set_hexpand(True)
        entry.set_text(value)
        entry.connect("changed", on_changed)
        entry_box.append(entry)

        choose_btn = Gtk.Button(label="Choose")
        choose_btn.add_css_class("flat")
        choose_btn.connect("clicked", on_choose)
        entry_box.append(choose_btn)

        clear_btn = Gtk.Button(label="Clear")
        clear_btn.add_css_class("flat")
        clear_btn.connect("clicked", on_clear)
        entry_box.append(clear_btn)

        content.append(entry_box)

        info_label = Gtk.Label(label=info_text, xalign=0)
        info_label.add_css_class("caption")
        info_label.add_css_class("dim-label")
        info_label.set_wrap(True)
        info_label.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        content.append(info_label)

        row.set_child(content)
        return row, entry

    def _set_editor_source_file(self, path: Path | None, notify: bool) -> None:
        self._editor_source_file = path.expanduser().resolve(strict=False) if path else None
        self._source_row_guard = True
        self._source_row.set_text(str(self._editor_source_file or ""))
        self._source_row_guard = False
        if notify:
            self._on_source_change(self._editor_source_file)

    def _on_choose_text_draft_template_dir(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose Text Draft template directory")
        dialog.select_folder(self, None, self._on_text_draft_template_dir_chosen)

    def _on_text_draft_template_dir_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.select_folder_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_text_draft_template_dir(path)

    def _on_clear_text_draft_template_dir(self, _button: Gtk.Button) -> None:
        self._set_text_draft_template_dir(None)

    def _on_text_draft_template_dir_row_changed(self, row: Gtk.Editable) -> None:
        if self._text_draft_template_dir_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_text_draft_template_dir(None)
            return
        self._set_text_draft_template_dir(Path(raw))

    def _set_text_draft_template_dir(self, path: Path | None) -> None:
        self._text_draft_template_dir = path.expanduser().resolve(strict=False) if path else None
        self._text_draft_template_dir_row_guard = True
        self._text_draft_template_dir_row.set_text(str(self._text_draft_template_dir or ""))
        self._text_draft_template_dir_row_guard = False

    def _on_choose_proof_suggestions_json_path(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose proofreading suggestions JSON")
        dialog.open(self, None, self._on_proof_suggestions_json_path_chosen)

    def _on_proof_suggestions_json_path_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_proof_suggestions_json_path(path)

    def _on_clear_proof_suggestions_json_path(self, _button: Gtk.Button) -> None:
        self._set_proof_suggestions_json_path(None)

    def _on_proof_suggestions_json_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._proof_suggestions_json_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_proof_suggestions_json_path(None)
            return
        self._set_proof_suggestions_json_path(Path(raw))

    def _set_proof_suggestions_json_path(self, path: Path | None) -> None:
        self._proof_suggestions_json_path = path.expanduser().resolve(strict=False) if path else None
        self._proof_suggestions_json_path_row_guard = True
        self._proof_suggestions_json_path_row.set_text(str(self._proof_suggestions_json_path or ""))
        self._proof_suggestions_json_path_row_guard = False

    def _on_choose_libreoffice_path(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose LibreOffice Python directory")
        dialog.select_folder(self, None, self._on_libreoffice_path_chosen)

    def _on_libreoffice_path_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.select_folder_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_libreoffice_python_path(path)

    def _on_clear_libreoffice_path(self, _button: Gtk.Button) -> None:
        self._set_libreoffice_python_path(None)

    def _on_libreoffice_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._libreoffice_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_libreoffice_python_path(None)
            return
        self._set_libreoffice_python_path(Path(raw))

    def _set_libreoffice_python_path(self, path: Path | None) -> None:
        self._libreoffice_python_path = path.expanduser().resolve(strict=False) if path else None
        self._libreoffice_path_row_guard = True
        self._libreoffice_path_row.set_text(str(self._libreoffice_python_path or ""))
        self._libreoffice_path_row_guard = False

    def _on_choose_concordance_file(self, _button: Gtk.Button) -> None:
        dialog = Gtk.FileDialog(title="Choose concordance file")
        dialog.open(self, None, self._on_concordance_file_chosen)

    def _on_concordance_file_chosen(self, dialog: Gtk.FileDialog, result: Gio.AsyncResult) -> None:  # noqa: D401
        try:
            file = dialog.open_finish(result)
            path = Path(file.get_path() or "")
        except Exception:
            return
        if not path:
            return
        self._set_concordance_file_path(path)

    def _on_clear_concordance_file(self, _button: Gtk.Button) -> None:
        self._set_concordance_file_path(None)

    def _on_concordance_path_row_changed(self, row: Gtk.Editable) -> None:
        if self._concordance_path_row_guard:
            return
        raw = row.get_text().strip()
        if not raw:
            self._set_concordance_file_path(None)
            return
        self._set_concordance_file_path(Path(raw))

    def _set_concordance_file_path(self, path: Path | None) -> None:
        self._concordance_file_path = path.expanduser().resolve(strict=False) if path else None
        self._concordance_path_row_guard = True
        self._concordance_path_row.set_text(str(self._concordance_file_path or ""))
        self._concordance_path_row_guard = False

    def _on_copy_normal_profile_clicked(self, _button: Gtk.Button) -> None:
        source = _normal_libreoffice_profile_path()
        target = LIBREOFFICE_PROFILE
        dialog = Adw.MessageDialog.new(
            self,
            "Copy normal LibreOffice profile?",
            (
                f"This replaces the Prose LibreOffice profile.\n\n"
                f"Source: {source}\n"
                f"Target: {target}\n\n"
                "Any Prose-specific LibreOffice settings in the target profile will be lost."
            ),
        )
        dialog.add_response("cancel", "Cancel")
        dialog.add_response("copy", "Copy Profile")
        dialog.set_response_appearance("copy", Adw.ResponseAppearance.DESTRUCTIVE)
        dialog.set_default_response("cancel")
        dialog.set_close_response("cancel")
        dialog.connect("response", self._on_copy_normal_profile_response)
        dialog.present()

    def _on_copy_normal_profile_response(self, _dialog: Adw.MessageDialog, response: str) -> None:
        if response != "copy":
            return
        ok, message = self._on_copy_normal_profile()
        self._parent_window._show_toast(message)
        if ok:
            self._parent_window._status_label.set_label("LibreOffice profile imported. Launch Writer to use it.")

    def _clear_list_box(self, list_box: Gtk.ListBox) -> None:
        child = list_box.get_first_child()
        while child is not None:
            next_child = child.get_next_sibling()
            list_box.remove(child)
            child = next_child

    def _current_editor_pinned_actions(self) -> list[str]:
        return [key for key in self._editor_action_order if self._quick_action_enabled.get(key, False)]

    def _update_quick_actions_hint(self) -> None:
        count = len(self._current_editor_pinned_actions())
        self._quick_actions_hint.set_label(
            f"{count} of {MAX_PINNED_EDITOR_ACTIONS} pinned. Unpinned actions stay under More."
        )

    def _rebuild_quick_action_rows(self) -> None:
        self._clear_list_box(self._quick_actions_list)
        self._update_quick_actions_hint()

        last_index = len(self._editor_action_order) - 1
        for index, key in enumerate(self._editor_action_order):
            definition = EDITOR_QUICK_ACTION_BY_KEY[key]
            row = Adw.ActionRow(title=definition.title, subtitle=definition.description)
            row.set_activatable(False)

            toggle = Gtk.CheckButton()
            toggle.set_valign(Gtk.Align.CENTER)
            toggle.set_active(self._quick_action_enabled.get(key, False))
            toggle.set_tooltip_text("Show this action in the Editor quick-actions row")
            toggle.connect("toggled", self._on_quick_action_toggled, key)
            row.add_prefix(toggle)

            up_btn = Gtk.Button(icon_name="go-up-symbolic")
            up_btn.add_css_class("flat")
            up_btn.set_tooltip_text("Move up")
            up_btn.set_sensitive(index > 0)
            up_btn.connect("clicked", self._on_move_quick_action_clicked, key, -1)
            row.add_suffix(up_btn)

            down_btn = Gtk.Button(icon_name="go-down-symbolic")
            down_btn.add_css_class("flat")
            down_btn.set_tooltip_text("Move down")
            down_btn.set_sensitive(index < last_index)
            down_btn.connect("clicked", self._on_move_quick_action_clicked, key, 1)
            row.add_suffix(down_btn)

            self._quick_actions_list.append(row)

    def _current_text_draft_pinned_actions(self) -> list[str]:
        return [key for key in self._text_draft_action_order if self._text_draft_quick_action_enabled.get(key, False)]

    def _update_text_draft_quick_actions_hint(self) -> None:
        count = len(self._current_text_draft_pinned_actions())
        self._text_draft_quick_actions_hint.set_label(
            f"{count} of {MAX_PINNED_EDITOR_ACTIONS} pinned. Unpinned actions stay under More."
        )

    def _rebuild_text_draft_quick_action_rows(self) -> None:
        self._clear_list_box(self._text_draft_quick_actions_list)
        self._update_text_draft_quick_actions_hint()

        last_index = len(self._text_draft_action_order) - 1
        for index, key in enumerate(self._text_draft_action_order):
            definition = TEXT_DRAFT_QUICK_ACTION_BY_KEY[key]
            row = Adw.ActionRow(title=definition.title, subtitle=definition.description)
            row.set_activatable(False)

            toggle = Gtk.CheckButton()
            toggle.set_valign(Gtk.Align.CENTER)
            toggle.set_active(self._text_draft_quick_action_enabled.get(key, False))
            toggle.set_tooltip_text("Show this action in the Text Draft quick-actions row")
            toggle.connect("toggled", self._on_text_draft_quick_action_toggled, key)
            row.add_prefix(toggle)

            up_btn = Gtk.Button(icon_name="go-up-symbolic")
            up_btn.add_css_class("flat")
            up_btn.set_tooltip_text("Move up")
            up_btn.set_sensitive(index > 0)
            up_btn.connect("clicked", self._on_move_text_draft_quick_action_clicked, key, -1)
            row.add_suffix(up_btn)

            down_btn = Gtk.Button(icon_name="go-down-symbolic")
            down_btn.add_css_class("flat")
            down_btn.set_tooltip_text("Move down")
            down_btn.set_sensitive(index < last_index)
            down_btn.connect("clicked", self._on_move_text_draft_quick_action_clicked, key, 1)
            row.add_suffix(down_btn)

            self._text_draft_quick_actions_list.append(row)

    def _on_quick_action_toggled(self, button: Gtk.CheckButton, key: str) -> None:
        if self._quick_action_toggle_guard:
            return

        active = button.get_active()
        was_active = self._quick_action_enabled.get(key, False)
        if active == was_active:
            return

        if active and len(self._current_editor_pinned_actions()) >= MAX_PINNED_EDITOR_ACTIONS:
            self._quick_action_toggle_guard = True
            button.set_active(False)
            self._quick_action_toggle_guard = False
            self._parent_window._show_toast(f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} quick actions.")
            return

        self._quick_action_enabled[key] = active
        self._update_quick_actions_hint()

    def _on_move_quick_action_clicked(self, _button: Gtk.Button, key: str, delta: int) -> None:
        index = self._editor_action_order.index(key)
        target_index = index + delta
        if target_index < 0 or target_index >= len(self._editor_action_order):
            return
        self._editor_action_order[index], self._editor_action_order[target_index] = (
            self._editor_action_order[target_index],
            self._editor_action_order[index],
        )
        self._rebuild_quick_action_rows()

    def _on_text_draft_quick_action_toggled(self, button: Gtk.CheckButton, key: str) -> None:
        if self._text_draft_quick_action_toggle_guard:
            return

        active = button.get_active()
        was_active = self._text_draft_quick_action_enabled.get(key, False)
        if active == was_active:
            return

        if active and len(self._current_text_draft_pinned_actions()) >= MAX_PINNED_EDITOR_ACTIONS:
            self._text_draft_quick_action_toggle_guard = True
            button.set_active(False)
            self._text_draft_quick_action_toggle_guard = False
            self._parent_window._show_toast(f"Pin up to {MAX_PINNED_EDITOR_ACTIONS} Text Draft quick actions.")
            return

        self._text_draft_quick_action_enabled[key] = active
        self._update_text_draft_quick_actions_hint()

    def _on_move_text_draft_quick_action_clicked(self, _button: Gtk.Button, key: str, delta: int) -> None:
        index = self._text_draft_action_order.index(key)
        target_index = index + delta
        if target_index < 0 or target_index >= len(self._text_draft_action_order):
            return
        self._text_draft_action_order[index], self._text_draft_action_order[target_index] = (
            self._text_draft_action_order[target_index],
            self._text_draft_action_order[index],
        )
        self._rebuild_text_draft_quick_action_rows()

    def trigger_save(self) -> None:
        self._on_save_clicked(None)

    def _on_close_clicked(self, _button: Gtk.Button) -> None:
        self.close()

    def _on_save_clicked(self, _button: Gtk.Button) -> None:
        model_profiles: list[ModelProfile] = []
        for profile_key in MODEL_PROFILE_IDS:
            widgets = self._model_profile_editors.get(profile_key)
            if widgets is None:
                continue
            model_profiles.append(
                ModelProfile(
                    key=profile_key,
                    nickname=widgets.nickname_row.get_text().strip() or DEFAULT_MODEL_PROFILE_NICKNAMES[profile_key],
                    abbreviation=widgets.abbreviation_row.get_text().strip(),
                    api_url=widgets.api_url_row.get_text().strip(),
                    model_id=widgets.model_row.get_text().strip(),
                    api_key=widgets.api_key_row.get_text().strip(),
                    disable_reasoning=widgets.disable_reasoning_row.get_active(),
                )
            )

        proof_widgets = self._prompt_editors.get("proof")
        spelling_widgets = self._prompt_editors.get("spelling")
        citation_validator_widgets = self._prompt_editors.get("citation-validator")
        thesaurus_widgets = self._prompt_editors.get("thesaurus")
        reference_widgets = self._prompt_editors.get("reference")
        ask_widgets = self._prompt_editors.get("ask")
        improve_widgets = self._prompt_editors.get("improve-generated")
        rephrase_widgets = self._prompt_editors.get("rephrase-generated")
        combine_widgets = self._prompt_editors.get("combine")
        shorten_widgets = self._prompt_editors.get("shorten")
        intro_widgets = self._prompt_editors.get("intro")
        intro_reply_widgets = self._prompt_editors.get("intro-reply")
        conclusion_widgets = self._prompt_editors.get("conclusion")
        concl_no_issues_widgets = self._prompt_editors.get("concl-no-issues")
        topic_sentence_widgets = self._prompt_editors.get("topic-sentence")
        concl_section_widgets = self._prompt_editors.get("concl-section")
        translate_widgets = self._prompt_editors.get("translate")
        text_draft_spelling_prompt_buffer = self._text_draft_spelling_prompt_buffer
        if not all(
            (
                proof_widgets,
                spelling_widgets,
                citation_validator_widgets,
                thesaurus_widgets,
                reference_widgets,
                ask_widgets,
                improve_widgets,
                rephrase_widgets,
                combine_widgets,
                shorten_widgets,
                intro_widgets,
                intro_reply_widgets,
                conclusion_widgets,
                concl_no_issues_widgets,
                topic_sentence_widgets,
                concl_section_widgets,
                translate_widgets,
                text_draft_spelling_prompt_buffer,
            )
        ):
            return

        proof_prompt_text = self._prompt_text(proof_widgets.prompt_buffer)
        spelling_prompt_text = self._prompt_text(spelling_widgets.prompt_buffer)
        citation_validator_prompt_text = self._prompt_text(citation_validator_widgets.prompt_buffer)
        text_draft_spelling_prompt_text = self._prompt_text(text_draft_spelling_prompt_buffer)
        thesaurus_prompt_text = self._prompt_text(thesaurus_widgets.prompt_buffer)
        reference_prompt_text = self._prompt_text(reference_widgets.prompt_buffer)
        ask_prompt_text = self._prompt_text(ask_widgets.prompt_buffer)
        improve_prompt_text = self._prompt_text(improve_widgets.prompt_buffer)
        rephrase_prompt_text = self._prompt_text(rephrase_widgets.prompt_buffer)
        combine_prompt_text = self._prompt_text(combine_widgets.prompt_buffer)
        shorten_prompt_text = self._prompt_text(shorten_widgets.prompt_buffer)
        intro_prompt_text = self._prompt_text(intro_widgets.prompt_buffer)
        intro_reply_prompt_text = self._prompt_text(intro_reply_widgets.prompt_buffer)
        conclusion_prompt_text = self._prompt_text(conclusion_widgets.prompt_buffer)
        concl_no_issues_prompt_text = self._prompt_text(concl_no_issues_widgets.prompt_buffer)
        topic_sentence_prompt_text = self._prompt_text(topic_sentence_widgets.prompt_buffer)
        concl_section_prompt_text = self._prompt_text(concl_section_widgets.prompt_buffer)
        translate_prompt_text = self._prompt_text(translate_widgets.prompt_buffer)
        shared_style_rules_text = (
            self._prompt_text(self._shared_style_rules_buffer)
            if self._shared_style_rules_buffer is not None
            else self._shared_style_rules_text
        )
        proof_settings = ProofreadSettings(
            api_url=proof_widgets.api_url_row.get_text().strip(),
            model_id=proof_widgets.model_row.get_text().strip(),
            api_key=proof_widgets.api_key_row.get_text().strip(),
            prompt=proof_prompt_text.strip() or DEFAULT_PROMPT,
            disable_reasoning=proof_widgets.disable_reasoning_row.get_active(),
            suggestions_json_path=self._proof_suggestions_json_path,
        )
        spelling_settings = SpellingStyleSettings(
            api_url=spelling_widgets.api_url_row.get_text().strip(),
            model_id=spelling_widgets.model_row.get_text().strip(),
            api_key=spelling_widgets.api_key_row.get_text().strip(),
            prompt=spelling_prompt_text.strip() or DEFAULT_SPELLINGSTYLE_PROMPT,
            disable_reasoning=spelling_widgets.disable_reasoning_row.get_active(),
        )
        citation_validator_settings = CitationValidatorSettings(
            api_url=citation_validator_widgets.api_url_row.get_text().strip(),
            model_id=citation_validator_widgets.model_row.get_text().strip(),
            api_key=citation_validator_widgets.api_key_row.get_text().strip(),
            prompt=citation_validator_prompt_text.strip() or DEFAULT_CITATION_NUMBER_VALIDATOR_PROMPT,
            disable_reasoning=citation_validator_widgets.disable_reasoning_row.get_active(),
        )
        thesaurus_settings = ThesaurusSettings(
            api_url=thesaurus_widgets.api_url_row.get_text().strip(),
            model_id=thesaurus_widgets.model_row.get_text().strip(),
            api_key=thesaurus_widgets.api_key_row.get_text().strip(),
            prompt=thesaurus_prompt_text.strip() or DEFAULT_THESAURUS_PROMPT,
            disable_reasoning=thesaurus_widgets.disable_reasoning_row.get_active(),
        )
        reference_settings = ReferenceSettings(
            api_url=reference_widgets.api_url_row.get_text().strip(),
            model_id=reference_widgets.model_row.get_text().strip(),
            api_key=reference_widgets.api_key_row.get_text().strip(),
            tavily_api_key=(reference_widgets.tavily_api_key_row.get_text().strip() if reference_widgets.tavily_api_key_row else ""),
            prompt=reference_prompt_text.strip() or DEFAULT_REFERENCE_PROMPT,
            disable_reasoning=reference_widgets.disable_reasoning_row.get_active(),
        )
        ask_settings = AskSettings(
            api_url=ask_widgets.api_url_row.get_text().strip(),
            model_id=ask_widgets.model_row.get_text().strip(),
            api_key=ask_widgets.api_key_row.get_text().strip(),
            tavily_api_key=(ask_widgets.tavily_api_key_row.get_text().strip() if ask_widgets.tavily_api_key_row else ""),
            prompt=ask_prompt_text.strip() or DEFAULT_ASK_PROMPT,
            disable_reasoning=ask_widgets.disable_reasoning_row.get_active(),
        )
        improve1_settings = Improve1Settings(
            api_url=improve_widgets.api_url_row.get_text().strip(),
            model_id=improve_widgets.model_row.get_text().strip(),
            api_key=improve_widgets.api_key_row.get_text().strip(),
            prompt=improve_prompt_text.strip() or DEFAULT_IMPROVE_PROMPT,
            disable_reasoning=improve_widgets.disable_reasoning_row.get_active(),
        )
        improve2_settings = Improve2Settings(
            api_url=rephrase_widgets.api_url_row.get_text().strip(),
            model_id=rephrase_widgets.model_row.get_text().strip(),
            api_key=rephrase_widgets.api_key_row.get_text().strip(),
            prompt=rephrase_prompt_text.strip() or DEFAULT_IMPROVE2_PROMPT,
            disable_reasoning=rephrase_widgets.disable_reasoning_row.get_active(),
        )
        combine_cites_settings = CombineCitesSettings(
            api_url=combine_widgets.api_url_row.get_text().strip(),
            model_id=combine_widgets.model_row.get_text().strip(),
            api_key=combine_widgets.api_key_row.get_text().strip(),
            prompt=combine_prompt_text.strip() or DEFAULT_COMBINE_CITES_PROMPT,
            disable_reasoning=combine_widgets.disable_reasoning_row.get_active(),
        )
        shorten_settings = ShortenSettings(
            api_url=shorten_widgets.api_url_row.get_text().strip(),
            model_id=shorten_widgets.model_row.get_text().strip(),
            api_key=shorten_widgets.api_key_row.get_text().strip(),
            prompt=shorten_prompt_text.strip() or DEFAULT_SHORTEN_PROMPT,
            disable_reasoning=shorten_widgets.disable_reasoning_row.get_active(),
        )
        introduction_settings = IntroductionSettings(
            api_url=intro_widgets.api_url_row.get_text().strip(),
            model_id=intro_widgets.model_row.get_text().strip(),
            api_key=intro_widgets.api_key_row.get_text().strip(),
            prompt=intro_prompt_text.strip() or DEFAULT_INTRO_PROMPT,
            disable_reasoning=intro_widgets.disable_reasoning_row.get_active(),
        )
        introduction_reply_settings = IntroductionSettings(
            api_url=intro_reply_widgets.api_url_row.get_text().strip(),
            model_id=intro_reply_widgets.model_row.get_text().strip(),
            api_key=intro_reply_widgets.api_key_row.get_text().strip(),
            prompt=intro_reply_prompt_text.strip() or DEFAULT_INTRO_REPLY_PROMPT,
            disable_reasoning=intro_reply_widgets.disable_reasoning_row.get_active(),
        )
        conclusion_settings = ConclusionSettings(
            api_url=conclusion_widgets.api_url_row.get_text().strip(),
            model_id=conclusion_widgets.model_row.get_text().strip(),
            api_key=conclusion_widgets.api_key_row.get_text().strip(),
            prompt=conclusion_prompt_text.strip() or DEFAULT_CONCLUSION_PROMPT,
            disable_reasoning=conclusion_widgets.disable_reasoning_row.get_active(),
        )
        concl_no_issues_settings = ConclNoIssuesSettings(
            api_url=concl_no_issues_widgets.api_url_row.get_text().strip(),
            model_id=concl_no_issues_widgets.model_row.get_text().strip(),
            api_key=concl_no_issues_widgets.api_key_row.get_text().strip(),
            prompt=concl_no_issues_prompt_text.strip() or DEFAULT_CONCL_NO_ISSUES_PROMPT,
            disable_reasoning=concl_no_issues_widgets.disable_reasoning_row.get_active(),
        )
        topic_sentence_settings = TopicSentenceSettings(
            api_url=topic_sentence_widgets.api_url_row.get_text().strip(),
            model_id=topic_sentence_widgets.model_row.get_text().strip(),
            api_key=topic_sentence_widgets.api_key_row.get_text().strip(),
            prompt=topic_sentence_prompt_text.strip() or DEFAULT_TOPIC_SENTENCE_PROMPT,
            disable_reasoning=topic_sentence_widgets.disable_reasoning_row.get_active(),
        )
        concl_section_settings = ConclSectionSettings(
            api_url=concl_section_widgets.api_url_row.get_text().strip(),
            model_id=concl_section_widgets.model_row.get_text().strip(),
            api_key=concl_section_widgets.api_key_row.get_text().strip(),
            prompt=concl_section_prompt_text.strip() or DEFAULT_CONCL_SECTION_PROMPT,
            disable_reasoning=concl_section_widgets.disable_reasoning_row.get_active(),
        )
        translate_settings = TranslateSettings(
            api_url=translate_widgets.api_url_row.get_text().strip(),
            model_id=translate_widgets.model_row.get_text().strip(),
            api_key=translate_widgets.api_key_row.get_text().strip(),
            prompt=translate_prompt_text.strip() or DEFAULT_TRANSLATE_PROMPT,
            disable_reasoning=translate_widgets.disable_reasoning_row.get_active(),
        )
        editor_action_profile_defaults = dict(self._editor_action_profile_defaults)
        dropdowns_by_action_key: dict[str, Gtk.DropDown] = {}
        for widgets in self._prompt_editors.values():
            if not widgets.default_profile_dropdowns:
                continue
            dropdowns_by_action_key.update(widgets.default_profile_dropdowns)
        dropdowns_by_action_key.update(self._choices_profile_dropdowns)
        for key in PROFILE_BACKED_COMMAND_KEYS:
            dropdown = dropdowns_by_action_key.get(key)
            if dropdown is None:
                continue
            selected = int(dropdown.get_selected())
            if selected <= 0:
                editor_action_profile_defaults[key] = None
                continue
            selected_index = selected - 1
            if selected_index < 0 or selected_index >= len(MODEL_PROFILE_IDS):
                editor_action_profile_defaults[key] = None
                continue
            editor_action_profile_defaults[key] = MODEL_PROFILE_IDS[selected_index]
        editor_pinned_action_ids = self._current_editor_pinned_actions()
        text_draft_pinned_action_ids = self._current_text_draft_pinned_actions()
        text_draft_external_actions = self._collect_text_draft_external_actions_from_settings()
        if text_draft_external_actions is None:
            return
        self._on_save(
            model_profiles,
            proof_settings,
            spelling_settings,
            text_draft_spelling_prompt_text.strip() or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT,
            citation_validator_settings,
            improve1_settings,
            improve2_settings,
            combine_cites_settings,
            thesaurus_settings,
            reference_settings,
            ask_settings,
            shorten_settings,
            introduction_settings,
            introduction_reply_settings,
            conclusion_settings,
            concl_no_issues_settings,
            topic_sentence_settings,
            concl_section_settings,
            translate_settings,
            shared_style_rules_text,
            editor_action_profile_defaults,
            editor_pinned_action_ids,
            text_draft_pinned_action_ids,
            self._text_draft_template_dir,
            text_draft_external_actions,
            self._libreoffice_python_path,
            self._concordance_file_path,
        )
        self._parent_window._show_toast("Settings saved.")

    def _on_prompt_row_selected(self, _listbox: Gtk.ListBox, row: Gtk.ListBoxRow | None) -> None:
        if not row:
            return
        key = self._prompt_row_keys.get(row)
        if key:
            self._prompt_stack.set_visible_child_name(key)

    def _build_style_rules_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Style Rules", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        subtitle = Gtk.Label(
            label=(
                f"These rules are inserted anywhere a prompt contains {STYLE_RULES_TOKEN}. "
                "Edit them here to change all token-based prompts at once."
            ),
            xalign=0,
        )
        subtitle.add_css_class("caption")
        subtitle.add_css_class("dim-label")
        subtitle.set_wrap(True)
        subtitle.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        page_box.append(subtitle)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Shared Rules", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_scroller, buffer = self._build_prompt_editor(self._shared_style_rules_text or DEFAULT_SHARED_STYLE_RULES)
        self._shared_style_rules_buffer = buffer
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_prompt_page(
        self,
        key: str,
        title: str,
        settings: (
            ProofreadSettings
            | SpellingStyleSettings
            | CitationValidatorSettings
            | Improve1Settings
            | Improve2Settings
            | CombineCitesSettings
            | ThesaurusSettings
            | ReferenceSettings
            | AskSettings
            | ShortenSettings
            | IntroductionSettings
            | ConclusionSettings
            | ConclNoIssuesSettings
            | TopicSentenceSettings
            | ConclSectionSettings
            | TranslateSettings
        ),
        default_prompt: str,
    ) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label=title, xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        uses_shared_profiles = key in PROFILE_BACKED_COMMAND_KEYS

        api_url_row = Adw.EntryRow(title="API URL")
        api_url_row.set_text(settings.api_url)

        model_row = Adw.EntryRow(title="Model ID (optional)")
        model_row.set_text(settings.model_id)

        api_key_row = Adw.PasswordEntryRow(title="API Key")
        api_key_row.set_text(settings.api_key)

        tavily_api_key_row = None
        disable_reasoning_row = Adw.SwitchRow(title="Disable reasoning")
        disable_reasoning_row.set_active(bool(settings.disable_reasoning))
        default_profile_dropdowns: dict[str, Gtk.DropDown] = {}

        if uses_shared_profiles:
            profile_group = Adw.PreferencesGroup(title="Default Model Profile")
            profile_group.add_css_class("list-stack")
            profile_group.set_hexpand(True)

            profile_row = Adw.PreferencesRow()
            profile_row.set_selectable(False)
            profile_row.set_activatable(False)
            profile_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=8)
            profile_box.set_margin_top(10)
            profile_box.set_margin_bottom(10)
            profile_box.set_margin_start(12)
            profile_box.set_margin_end(12)

            caption_text = "Choose the default profile used by this command. Its reasoning setting will also be used."
            profile_keys = (key,)
            if key == "improve-generated":
                caption_text = (
                    "Choose the default profiles used by Improve Generated and Improve Selected. "
                    "Generated Choices and Selected Choices use the separate Choices Models page."
                )
                profile_keys = (
                    "improve-generated",
                    "improve-selected",
                )

            caption = Gtk.Label(label=caption_text, xalign=0)
            caption.add_css_class("caption")
            caption.add_css_class("dim-label")
            caption.set_wrap(True)
            caption.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
            profile_box.append(caption)

            for profile_key in profile_keys:
                if len(profile_keys) > 1:
                    profile_label = Gtk.Label(
                        label=PROFILE_BACKED_COMMAND_TITLES.get(profile_key, profile_key.replace("-", " ").title()),
                        xalign=0,
                    )
                    profile_label.add_css_class("caption")
                    profile_label.add_css_class("dim-label")
                    profile_box.append(profile_label)

                dropdown = Gtk.DropDown(model=self._profile_dropdown_model(include_unset=True))
                selected_profile_key = self._editor_action_profile_defaults.get(profile_key)
                selected_index = (
                    MODEL_PROFILE_IDS.index(selected_profile_key) + 1
                    if selected_profile_key in MODEL_PROFILE_IDS
                    else 0
                )
                dropdown.set_selected(selected_index)
                profile_box.append(dropdown)
                default_profile_dropdowns[profile_key] = dropdown

            profile_row.set_child(profile_box)
            profile_group.add(profile_row)
            page_box.append(profile_group)
        else:
            credentials_group = Adw.PreferencesGroup(title="Credentials")
            credentials_group.add_css_class("list-stack")
            credentials_group.set_hexpand(True)
            page_box.append(credentials_group)

            credentials_group.add(api_url_row)
            credentials_group.add(model_row)
            credentials_group.add(api_key_row)

            if isinstance(settings, (ReferenceSettings, AskSettings)):
                tavily_api_key_row = Adw.PasswordEntryRow(title="Tavily API Key")
                tavily_api_key_row.set_text(settings.tavily_api_key)
                credentials_group.add(tavily_api_key_row)

            credentials_group.add(disable_reasoning_row)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Prompt", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_hint = Gtk.Label(label=STYLE_RULES_HINT_TEXT, xalign=0)
        prompt_hint.add_css_class("caption")
        prompt_hint.add_css_class("dim-label")
        prompt_hint.set_wrap(True)
        prompt_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        prompt_section.append(prompt_hint)
        prompt_scroller, buffer = self._build_prompt_editor(settings.prompt or default_prompt)
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        ask_prompt_buffer = None

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)

        self._prompt_editors[key] = PromptEditorWidgets(
            api_url_row=api_url_row,
            model_row=model_row,
            api_key_row=api_key_row,
            tavily_api_key_row=tavily_api_key_row,
            disable_reasoning_row=disable_reasoning_row,
            prompt_buffer=buffer,
            ask_prompt_buffer=ask_prompt_buffer,
            default_profile_dropdowns=default_profile_dropdowns or None,
        )
        return page

    def _build_text_draft_spellingstyle_page(self) -> Gtk.Widget:
        page_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=12)
        page_box.set_margin_top(12)
        page_box.set_margin_bottom(12)
        page_box.set_margin_start(12)
        page_box.set_margin_end(12)
        page_box.set_vexpand(True)

        title_label = Gtk.Label(label="Text Draft SpellingStyle", xalign=0)
        title_label.add_css_class("title-3")
        page_box.append(title_label)

        details_group = Adw.PreferencesGroup(title="Shared Credentials")
        details_group.add_css_class("list-stack")
        details_group.set_hexpand(True)
        page_box.append(details_group)

        details_row = Adw.ActionRow(
            title="Uses the main SpellingStyle model profile",
            subtitle=(
                "Text Draft SpellingStyle shares the default model profile selected on the SpellingStyle page. "
                "Only the prompt below is separate."
            ),
        )
        details_row.set_activatable(False)
        details_group.add(details_row)

        prompt_section = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        prompt_section.set_hexpand(True)
        prompt_section.set_vexpand(True)
        prompt_label = Gtk.Label(label="Prompt", xalign=0)
        prompt_label.add_css_class("dim-label")
        prompt_section.append(prompt_label)
        prompt_hint = Gtk.Label(label=STYLE_RULES_HINT_TEXT, xalign=0)
        prompt_hint.add_css_class("caption")
        prompt_hint.add_css_class("dim-label")
        prompt_hint.set_wrap(True)
        prompt_hint.set_wrap_mode(Pango.WrapMode.WORD_CHAR)
        prompt_section.append(prompt_hint)
        prompt_scroller, buffer = self._build_prompt_editor(
            self._text_draft_spelling_prompt or DEFAULT_TEXT_DRAFT_SPELLINGSTYLE_PROMPT
        )
        self._text_draft_spelling_prompt_buffer = buffer
        prompt_section.append(prompt_scroller)
        page_box.append(prompt_section)

        page = Gtk.ScrolledWindow()
        page.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        page.set_hexpand(True)
        page.set_vexpand(True)
        page.set_child(page_box)
        return page

    def _build_prompt_editor(self, text: str) -> tuple[Gtk.ScrolledWindow, Gtk.TextBuffer]:
        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        scroller.set_hexpand(True)
        scroller.set_vexpand(True)
        scroller.set_has_frame(False)

        buffer = Gtk.TextBuffer()
        buffer.set_text(text)
        prompt_view = Gtk.TextView.new_with_buffer(buffer)
        prompt_view.set_wrap_mode(Gtk.WrapMode.WORD_CHAR)
        prompt_view.set_monospace(True)
        prompt_view.set_vexpand(True)
        prompt_view.set_hexpand(True)
        prompt_view.set_top_margin(12)
        prompt_view.set_bottom_margin(12)
        prompt_view.set_left_margin(12)
        prompt_view.set_right_margin(12)
        scroller.set_child(prompt_view)
        return scroller, buffer

    def _prompt_text(self, buffer: Gtk.TextBuffer) -> str:
        start, end = buffer.get_bounds()
        return buffer.get_text(start, end, True)


class EditorCommandsWindow(Adw.ApplicationWindow):
    def __init__(self, parent: ProseWindow) -> None:
        super().__init__(application=parent.get_application(), title="Editor Commands")
        self._parent = parent
        self.set_default_size(900, 720)
        self.set_resizable(True)
        self._build_ui()

    def _build_ui(self) -> None:
        view = Adw.ToolbarView()
        header = Adw.HeaderBar()
        header.add_css_class("flat")
        header.set_title_widget(Adw.WindowTitle(title="Editor Commands", subtitle="Run or copy actions"))
        view.add_top_bar(header)

        page = Adw.PreferencesPage()
        intro = Adw.PreferencesGroup(
            title="How to use",
            description=(
                "Use Run to trigger actions inside the open Prose window. "
                "Use Copy Command to place the GApplication call on your clipboard."
            ),
        )
        page.add(intro)

        actions_group = Adw.PreferencesGroup(title="Editor Actions")
        page.add(actions_group)
        for title, action_name, param, desc, supports_profiles in _editor_command_items():
            row = self._build_action_row(title, action_name, param, desc, supports_profiles=supports_profiles)
            actions_group.add(row)

        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        scroller.set_child(page)
        view.set_content(scroller)
        self.set_content(view)

    def _build_action_row(
        self,
        title: str,
        action_name: str,
        param: str | None,
        desc: str,
        with_index: bool = False,
        supports_profiles: bool = False,
    ) -> Adw.ActionRow:
        row = Adw.ActionRow(title=title, subtitle=desc)
        row.set_activatable(False)

        suffix = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=6)
        spin = None
        profile_dropdown = None
        if with_index:
            spin = Gtk.SpinButton.new_with_range(0, 9999, 1)
            spin.set_value(0)
            spin.set_width_chars(4)
            suffix.append(spin)
        elif supports_profiles:
            profile_dropdown = Gtk.DropDown(
                model=Gtk.StringList.new(
                    [profile.display_name() for profile in self._parent._model_profiles]
                )
            )
            lookup_key = _profile_default_lookup_key_for_action_name(action_name)
            selected_profile_key = self._parent._editor_action_profile_defaults.get(lookup_key) if lookup_key else None
            selected_index = MODEL_PROFILE_IDS.index(selected_profile_key) if selected_profile_key in MODEL_PROFILE_IDS else 0
            profile_dropdown.set_selected(selected_index)
            suffix.append(profile_dropdown)

        run_btn = Gtk.Button(label="Run")
        run_btn.add_css_class("suggested-action")
        run_btn.add_css_class("flat")
        run_btn.connect("clicked", self._on_run_clicked, action_name, spin, profile_dropdown)
        suffix.append(run_btn)

        copy_btn = Gtk.Button(label="Copy Command")
        copy_btn.add_css_class("flat")
        copy_btn.add_css_class("link")
        copy_btn.connect("clicked", self._on_copy_clicked, action_name, spin, param, profile_dropdown)
        suffix.append(copy_btn)

        row.add_suffix(suffix)
        return row

    def _on_run_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        spin: Gtk.SpinButton | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        app = self.get_application()
        if spin is None:
            if profile_dropdown is not None:
                selected = int(profile_dropdown.get_selected())
                nicknames = [profile.display_name() for profile in self._parent._model_profiles]
                if 0 <= selected < len(nicknames):
                    app.activate_action(action_name, GLib.Variant("s", nicknames[selected]))
                return
            app.activate_action(action_name, None)
            return
        value = int(spin.get_value())
        app.activate_action(action_name, GLib.Variant("i", value))

    def _on_copy_clicked(
        self,
        _button: Gtk.Button,
        action_name: str,
        spin: Gtk.SpinButton | None,
        param: str | None,
        profile_dropdown: Gtk.DropDown | None,
    ) -> None:
        if spin is not None:
            param = _format_action_param(GLib.Variant("i", int(spin.get_value())))
        elif profile_dropdown is not None:
            selected = int(profile_dropdown.get_selected())
            nicknames = [profile.display_name() for profile in self._parent._model_profiles]
            if 0 <= selected < len(nicknames):
                param = _format_action_param(GLib.Variant("s", nicknames[selected]))
        app = self.get_application()
        object_path = ACTION_OBJECT_PATH
        if isinstance(app, Gio.Application):
            app_path = app.get_dbus_object_path()
            if app_path:
                object_path = app_path
        command = _action_command(action_name, param, object_path)
        display = Gdk.Display.get_default()
        if display:
            display.get_clipboard().set(command)


def main() -> None:
    app = ProseApp()
    app.run(sys.argv)


if __name__ == "__main__":
    main()
