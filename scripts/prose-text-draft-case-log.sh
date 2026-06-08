#!/usr/bin/env bash
set -euo pipefail

helper="/home/jesse/Dropbox/MCGLAW/config_files/scripts/PROJECTS/CurrentCaseTui/helpers/case_log_actions.py"
draft_file="${1:-${PROSE_TEXT_DRAFT_FILE:-}}"

error() {
  local message="$1"
  if command -v zenity >/dev/null 2>&1; then
    zenity --error --title="CASE LOG" --text="$message"
  else
    printf 'CASE LOG: %s\n' "$message" >&2
  fi
  exit 1
}

if [[ -z "$draft_file" ]]; then
  error "No Text Draft file was provided."
fi

if [[ ! -f "$draft_file" ]]; then
  error "Draft file not found: $draft_file"
fi

if [[ ! -f "$helper" ]]; then
  error "Case log helper not found: $helper"
fi

output="$(python3 "$helper" append --draft-file "$draft_file" --quiet 2>&1)" || {
  error "${output:-Unable to write case log.}"
}
