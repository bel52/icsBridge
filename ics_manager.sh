#!/usr/bin/env bash
set -euo pipefail

# === Paths & venv ===
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_BIN="$ROOT_DIR/.venv/bin"
PY="$VENV_BIN/python3"

# === State files ===
CONF_FILE="$ROOT_DIR/.icsbridge_config"         # stores CAL_NAME & CAL_INDEX
TRACK="$ROOT_DIR/.tracked_sources.json"         # newline-delimited JSON: one per source
: > /dev/null # noop

# --- Create venv if missing ---
ensure_venv() {
  if [[ ! -x "$PY" ]]; then
    echo "No venv found at $VENV_BIN; creating one and installing deps..."
    python3 -m venv "$ROOT_DIR/.venv"
    "$PY" -m pip install --upgrade pip >/dev/null
    "$PY" -m pip install icalendar python-dateutil >/dev/null
  fi
}

# --- Config handling (no jq) ---
load_config() {
  # Default unset until file exists
  CAL_NAME=""
  CAL_INDEX=""
  if [[ -f "$CONF_FILE" ]]; then
    # shellcheck disable=SC1090
    source "$CONF_FILE" || true
  fi
}

save_config() {
  local name="$1"
  local index="$2"
  cat > "$CONF_FILE" <<CFG
# Persisted defaults for ICS Bridge
CAL_NAME='${name//\'/\'\\\'\'}'
CAL_INDEX='${index//\'/\'\\\'\'}'
CFG
  echo "✅ Saved defaults: calendar='$name', index=$index"
}

require_defaults_or_prompt() {
  load_config
  if [[ -z "${CAL_NAME:-}" || -z "${CAL_INDEX:-}" ]]; then
    echo
    echo "No default target calendar configured yet."
    read -rp "Enter target Outlook calendar name (e.g., Calendar): " CAL_NAME
    read -rp "Enter occurrence index for that name (e.g., 2): " CAL_INDEX
    CAL_INDEX="${CAL_INDEX:-2}"
    save_config "$CAL_NAME" "$CAL_INDEX"
  fi
}

set_defaults() {
  load_config
  echo
  echo "════════ Set Default Target Calendar ════════"
  echo "Current: calendar='${CAL_NAME:-<unset>}', index='${CAL_INDEX:-<unset>}'"
  read -rp "New target Outlook calendar name: " NEW_NAME
  read -rp "New occurrence index (number): " NEW_IDX
  NEW_IDX="${NEW_IDX:-2}"
  save_config "$NEW_NAME" "$NEW_IDX"
  read -rp "Press Enter to continue…" _
}

list_sources() {
  echo "═════════ Tracked Sources ═════════"
  if [[ -s "$TRACK" ]]; then
    nl -ba "$TRACK"
  else
    echo "(none)"
  fi
  echo "───────────────────────────────────"
}

add_calendar() {
  require_defaults_or_prompt
  load_config  # ensure CAL_NAME & CAL_INDEX present

  echo
  read -rp "Enter calendar source (URL): " SRC
  read -rp "Enter a short ID for this calendar (e.g., lions-2025): " ID

  TMP="/tmp/${ID}.ics"
  echo -e "\n🔄 Preparing tagged ICS file..."
  echo "Fetching and processing: $SRC"

  "$PY" "$ROOT_DIR/prepare_ics_for_import.py" "$SRC" "$ID" "$TMP"
  echo "✅ ICS file is ready at $TMP."

  echo
  echo "Opening Outlook…"
  osascript -e 'tell application "Microsoft Outlook" to activate' >/dev/null 2>&1 || true

  echo
  echo "➡ Import $TMP into \"${CAL_NAME}\" (#${CAL_INDEX})."
  echo "   (Change defaults via menu option: Set Default Target Calendar)"
  echo "{\"id\":\"$ID\",\"url\":\"$SRC\",\"calendar\":\"$CAL_NAME\",\"index\":${CAL_INDEX}}" >> "$TRACK"
  echo -e "\n✨ Import process initiated."
}

remove_calendar() {
  echo "════════ Remove Calendar Events ════════"
  if [[ ! -s "$TRACK" ]]; then
    echo "(none tracked)"; read -rp "Press Enter to continue…" _; return
  fi
  list_sources
  read -rp "Enter number to remove (or 'q' to cancel): " N
  [[ "$N" =~ ^[0-9]+$ ]] || { echo "Cancelled."; sleep 1; return; }

  # Extract JSON line
  LINE="$(sed -n "${N}p" "$TRACK")" || true
  if [[ -z "$LINE" ]]; then
    echo "No such entry."; sleep 1; return
  fi
  ID="$(echo "$LINE" | sed -E 's/.*"id":"([^"]+)".*/\1/')"
  CAL="$(echo "$LINE" | sed -E 's/.*"calendar":"([^"]+)".*/\1/')"
  IDX="$(echo "$LINE" | sed -E 's/.*"index":([0-9]+).*/\1/')"

  echo "🗑️  Removing tracked entry for '$ID' (Calendar=\"$CAL\" #$IDX)…"
  tmpf="$(mktemp)"; awk -v n="$N" 'NR!=n' "$TRACK" > "$tmpf" && mv "$tmpf" "$TRACK"
  echo "✅ Removed '$ID' from tracking."
  echo "(If you have a separate deletion script that cleans Outlook items, run it as usual.)"
  read -rp "Press Enter to continue…" _
}

menu() {
  clear
  cat <<MENU
╔═════════ ICS Bridge for Outlook ═════════╗
║ 1) ➕ Add Calendar via Outlook Import    ║
║ 2) 🗑️  Remove Imported Calendar         ║
║ 3) 📋 List Imported Calendars          ║
║ 4) 🛠️  Set Default Target Calendar      ║
║ 5) ❌ Quit                               ║
╚══════════════════════════════════════════╝
MENU
  read -rp "Choose [1-5]: " CHOICE
}

main() {
  ensure_venv
  touch "$TRACK"

  while true; do
    menu
    case "${CHOICE:-}" in
      1) add_calendar ;;
      2) remove_calendar ;;
      3) list_sources; read -rp "Press Enter to continue…" _ ;;
      4) set_defaults ;;
      5) exit 0 ;;
      *) echo "Invalid choice"; sleep 1 ;;
    esac
  done
}

main "$@"
