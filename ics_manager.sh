#!/usr/bin/env bash
set -euo pipefail

# ====== Settings ======
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_BIN="$ROOT_DIR/.venv/bin"
PY="$VENV_BIN/python3"

CONF_FILE="$ROOT_DIR/.icsbridge_config"   # stores CAL_NAME & CAL_INDEX
TRACK="$ROOT_DIR/.tracked_sources.json"   # newline-delimited JSON entries
MARK_DIR="$ROOT_DIR/.sources"             # per-source marker files (redundant, helpful)

# UI behavior (no auto-clear; always verbose)
: "${ICSBRIDGE_VERBOSE:=1}"

mkdir -p "$MARK_DIR"
touch "$TRACK"

ts() { date "+%Y-%m-%d %H:%M:%S"; }
log() { if [[ "${ICSBRIDGE_VERBOSE}" = "1" ]]; then echo "[$(ts)] $*"; fi; }

ensure_venv() {
  if [[ ! -x "$PY" ]]; then
    log "Creating venv at $VENV_BIN and installing deps…"
    python3 -m venv "$ROOT_DIR/.venv"
    "$PY" -m pip install --upgrade pip >/dev/null
    "$PY" -m pip install icalendar python-dateutil >/dev/null
  fi
}

load_config() {
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
CAL_NAME='${name//\'/\'\\\'\'}'
CAL_INDEX='${index//\'/\'\\\'\'}'
CFG
  log "Saved defaults: calendar='${name}', index=${index}"
}

require_defaults_or_prompt() {
  load_config
  if [[ -z "${CAL_NAME:-}" || -z "${CAL_INDEX:-}" ]]; then
    echo
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
  echo "═════════ Tracked Sources (file) ═════════"
  if [[ -s "$TRACK" ]]; then nl -ba "$TRACK"; else echo "(none)"; fi
  echo "───────────────────────────────────────────"
  echo "═════════ Tracked Sources (.sources) ═════"
  ls -1 "$MARK_DIR" 2>/dev/null | sed 's/.json$//' || echo "(none)"
  echo "───────────────────────────────────────────"
}

add_calendar() {
  ensure_venv
  require_defaults_or_prompt
  load_config

  echo
  read -rp "Enter calendar source (URL): " SRC
  read -rp "Enter a short ID for this calendar (e.g., lions-2025): " ID

  TMP_ICS="/tmp/${ID}.ics"

  log "Preparing ICS…"
  log "Fetch & normalize: $SRC  ->  $TMP_ICS"
  "$PY" "$ROOT_DIR/prepare_ics_for_import.py" "$SRC" "$ID" "$TMP_ICS"

  log "Opening Outlook import for: $TMP_ICS"
  open -a "Microsoft Outlook" "$TMP_ICS" || {
    echo "⚠️  Could not open Outlook with the ICS; try double-clicking $TMP_ICS manually."
  }

  echo
  echo "➡ In Outlook's import, choose: \"${CAL_NAME}\" (#${CAL_INDEX})."
  echo "   (Change defaults via menu option: Set Default Target Calendar)"

  echo "{\"id\":\"$ID\",\"url\":\"$SRC\",\"calendar\":\"$CAL_NAME\",\"index\":${CAL_INDEX}}" >> "$TRACK"
  echo "{\"id\":\"$ID\",\"url\":\"$SRC\",\"calendar\":\"$CAL_NAME\",\"index\":${CAL_INDEX}}" > "$MARK_DIR/${ID}.json"

  log "Done. Imported '$ID' (pending your confirmation in Outlook)."
  read -rp "Press Enter to continue…" _
}

remove_calendar() {
  ensure_venv
  require_defaults_or_prompt
  load_config

  echo "════════ Remove Calendar Events ════════"
  list_sources
  echo
  echo "Options:"
  echo "  a) Remove by choosing a tracked entry"
  echo "  b) Remove by entering an SRC ID manually (even if not tracked)"
  read -rp "Choose [a/b] (default a): " MODE
  MODE="${MODE:-a}"

  if [[ "$MODE" == "a" && ! -s "$TRACK" ]]; then
    echo "(No file-tracked entries. Switching to manual SRC ID.)"
    MODE="b"
  fi

  local ID CAL IDX
  if [[ "$MODE" == "a" ]]; then
    read -rp "Enter number to remove (or 'q' to cancel): " N
    [[ "$N" =~ ^[0-9]+$ ]] || { echo "Cancelled."; sleep 1; return; }
    LINE="$(sed -n "${N}p" "$TRACK" || true)"
    if [[ -z "$LINE" ]]; then echo "No such entry."; sleep 1; return; fi

    ID="$(echo "$LINE" | sed -E 's/.*\"id\":\"([^\"]+)\".*/\1/')"
    CAL="$(echo "$LINE" | sed -E 's/.*\"calendar\":\"([^\"]+)\".*/\1/')"
    IDX="$(echo "$LINE" | sed -E 's/.*\"index\":([0-9]+).*/\1/')"

    log "Deleting [SRC: $ID] from '$CAL' (#$IDX)…"
    osascript "$ROOT_DIR/outlook_delete_by_src.applescript" "$ID" "$CAL" "$IDX" \
      && log "Deleted events tagged [SRC: $ID]" \
      || echo "⚠️  Delete script reported an error."

    tmpf="$(mktemp)"; awk -v n="$N" 'NR!=n' "$TRACK" > "$tmpf" && mv "$tmpf" "$TRACK"
    rm -f "$MARK_DIR/${ID}.json"
    read -rp "Press Enter to continue…" _
  else
    read -rp "Enter SRC ID to remove (e.g., lions): " ID
    if [[ -f "$MARK_DIR/${ID}.json" ]]; then
      CAL="$(sed -E 's/.*\"calendar\":\"([^\"]+)\".*/\1/;t;d' "$MARK_DIR/${ID}.json" || true)"
      IDX="$(sed -E 's/.*\"index\":([0-9]+).*/\1/;t;d' "$MARK_DIR/${ID}.json" || true)"
      CAL="${CAL:-$CAL_NAME}"
      IDX="${IDX:-$CAL_INDEX}"
    else
      CAL="$CAL_NAME"
      IDX="$CAL_INDEX"
    fi
    log "Deleting [SRC: $ID] from '$CAL' (#$IDX)…"
    osascript "$ROOT_DIR/outlook_delete_by_src.applescript" "$ID" "$CAL" "$IDX" \
      && log "Deleted events tagged [SRC: $ID]" \
      || echo "⚠️  Delete script reported an error."

    tmpf="$(mktemp)"; grep -v "\"id\":\"$ID\"" "$TRACK" > "$tmpf" || true; mv "$tmpf" "$TRACK"
    rm -f "$MARK_DIR/${ID}.json"
    read -rp "Press Enter to continue…" _
  fi
}

menu() {
  cat <<MENU
╔═════════ ICS Bridge for Outlook ═════════╗
║ 1) ➕ Add Calendar via Outlook Import    ║
║ 2) 🗑️  Remove Imported Calendar         ║
║ 3) 📋 List Imported Calendars           ║
║ 4) 🛠️  Set Default Target Calendar      ║
║ 5) ❌ Quit                               ║
╚══════════════════════════════════════════╝
MENU
  read -rp "Choose [1-5]: " CHOICE
}

main() {
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
