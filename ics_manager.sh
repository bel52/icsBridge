#!/usr/bin/env bash
set -euo pipefail

# ===== Paths =====
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_BIN="$ROOT_DIR/.venv/bin"
PY="$VENV_BIN/python3"

# Persistent tracker (outside the repo) — survives codebase changes
PERSIST_DIR="${HOME}/.icsbridge"
P_TRACK="${PERSIST_DIR}/tracked.jsonl"     # JSON Lines: {"id":"...","url":"...","calendar":"...","index":N}
P_MARK_DIR="${PERSIST_DIR}/sources"        # per-source mirrors (optional)

# Back-compat (in-repo) — kept so older entries still show up
TRACK_LOCAL="${ROOT_DIR}/.tracked_sources.json"
MARK_LOCAL="${ROOT_DIR}/.sources"

CONF_FILE="${ROOT_DIR}/.icsbridge_config"  # defaults: CAL_NAME, CAL_INDEX

: "${ICSBRIDGE_VERBOSE:=1}"

mkdir -p "${PERSIST_DIR}" "${P_MARK_DIR}" "${MARK_LOCAL}"
touch "${P_TRACK}" "${TRACK_LOCAL}"

ts(){ date "+%Y-%m-%d %H:%M:%S"; }
log(){ [[ "$ICSBRIDGE_VERBOSE" = "1" ]] && echo "[$(ts)] $*"; }

ensure_venv(){
  if [[ ! -x "$PY" ]]; then
    log "Creating venv and installing deps…"
    python3 -m venv "$ROOT_DIR/.venv"
    "$PY" -m pip install --upgrade pip >/dev/null
    "$PY" -m pip install icalendar python-dateutil >/dev/null
  fi
}

load_config(){ CAL_NAME=""; CAL_INDEX="";
  [[ -f "$CONF_FILE" ]] && source "$CONF_FILE" || true; }

save_config(){ printf "CAL_NAME='%s'\nCAL_INDEX='%s'\n" "$1" "$2" > "$CONF_FILE"; log "Saved defaults: $1 #$2"; }

require_defaults_or_prompt(){
  load_config
  if [[ -z "${CAL_NAME:-}" || -z "${CAL_INDEX:-}" ]]; then
    echo
    read -rp "Enter target Outlook calendar name (e.g., Calendar): " CAL_NAME
    read -rp "Enter occurrence index for that name (e.g., 2): " CAL_INDEX
    CAL_INDEX="${CAL_INDEX:-2}"; save_config "$CAL_NAME" "$CAL_INDEX"
  fi
  # guard: mis-set names (URLs) silently break later steps
  if [[ "${CAL_NAME}" == *"://"* ]]; then
    echo "⚠️  Saved calendar looks like a URL (${CAL_NAME}). Resetting…"
    read -rp "Enter target Outlook calendar name (e.g., Calendar): " CAL_NAME
    read -rp "Enter occurrence index for that name (e.g., 2): " CAL_INDEX
    CAL_INDEX="${CAL_INDEX:-2}"; save_config "$CAL_NAME" "$CAL_INDEX"
  fi
}

# ---- Tracking helpers (write to persistent AND local for back-compat) ----
upsert_track(){
  local id="$1" url="$2" cal="$3" idx="$4"
  local line="{\"id\":\"$id\",\"url\":\"$url\",\"calendar\":\"$cal\",\"index\":${idx}}"

  # persistent
  tmp="$(mktemp)"; grep -v "\"id\":\"$id\"" "$P_TRACK" > "$tmp" || true; mv "$tmp" "$P_TRACK"
  echo "$line" >> "$P_TRACK"
  echo "$line" > "$P_MARK_DIR/${id}.json"

  # local (for older scripts you may still have lying around)
  tmp="$(mktemp)"; grep -v "\"id\":\"$id\"" "$TRACK_LOCAL" > "$tmp" || true; mv "$tmp" "$TRACK_LOCAL"
  echo "$line" >> "$TRACK_LOCAL"
  echo "$line" > "$MARK_LOCAL/${id}.json"
}

delete_track(){
  local id="$1"
  tmp="$(mktemp)"; grep -v "\"id\":\"$id\"" "$P_TRACK" > "$tmp" || true; mv "$tmp" "$P_TRACK"
  rm -f "$P_MARK_DIR/${id}.json"
  tmp="$(mktemp)"; grep -v "\"id\":\"$id\"" "$TRACK_LOCAL" > "$tmp" || true; mv "$tmp" "$TRACK_LOCAL"
  rm -f "$MARK_LOCAL/${id}.json"
}

get_all_ids(){
  { sed -n 's/.*"id":"\([^"]\+\)".*/\1/p' "$P_TRACK"
    sed -n 's/.*"id":"\([^"]\+\)".*/\1/p' "$TRACK_LOCAL"
    (ls -1 "$P_MARK_DIR" 2>/dev/null || true) | sed 's/\.json$//'
    (ls -1 "$MARK_LOCAL" 2>/dev/null || true) | sed 's/\.json$//'
  } | sort -u
}

set_defaults(){
  load_config
  echo "Current: '${CAL_NAME:-<unset>}' #${CAL_INDEX:-<unset>}"
  read -rp "New calendar name: " n
  read -rp "New index (number): " i
  i="${i:-2}"
  save_config "$n" "$i"
  read -rp "Press Enter to continue…" _
}

list_sources(){
  require_defaults_or_prompt
  echo "═════════ Tracked Calendars (persistent + local) ═════════"
  ids="$(get_all_ids)"
  if [[ -z "$ids" ]]; then
    echo "(none)"; echo "────────────────────────────────────────────"; return
  fi
  printf "%-16s %-10s %-20s %s\n" "ID" "Index" "Calendar" "URL"
  printf "%-16s %-10s %-20s %s\n" "--" "-----" "--------" "---"
  while IFS= read -r id; do
    [[ -z "$id" ]] && continue
    # prefer persistent; fall back to local; then to defaults
    line="$(grep "\"id\":\"$id\"" "$P_TRACK" | tail -1)"
    [[ -z "$line" ]] && line="$(grep "\"id\":\"$id\"" "$TRACK_LOCAL" | tail -1)"
    url="$(printf "%s" "$line" | sed -n 's/.*"url":"\([^"]*\)".*/\1/p')"
    cal="$(printf "%s" "$line" | sed -n 's/.*"calendar":"\([^"]*\)".*/\1/p')"
    idx="$(printf "%s" "$line" | sed -n 's/.*"index":\([0-9][0-9]*\).*/\1/p')"
    [[ -z "$cal" ]] && [[ -f "$P_MARK_DIR/${id}.json" ]] && cal="$(sed -n 's/.*"calendar":"\([^"]*\)".*/\1/p' "$P_MARK_DIR/${id}.json")"
    [[ -z "$idx" ]] && [[ -f "$P_MARK_DIR/${id}.json" ]] && idx="$(sed -n 's/.*"index":\([0-9][0-9]*\).*/\1/p' "$P_MARK_DIR/${id}.json")"
    [[ -z "$cal" ]] && [[ -f "$MARK_LOCAL/${id}.json" ]] && cal="$(sed -n 's/.*"calendar":"\([^"]*\)".*/\1/p' "$MARK_LOCAL/${id}.json")"
    [[ -z "$idx" ]] && [[ -f "$MARK_LOCAL/${id}.json" ]] && idx="$(sed -n 's/.*"index":\([0-9][0-9]*\).*/\1/p' "$MARK_LOCAL/${id}.json")"
    [[ -z "$cal" ]] && cal="$CAL_NAME"
    [[ -z "$idx" ]] && idx="$CAL_INDEX"
    [[ -z "$url" ]] && url=""
    printf "%-16s %-10s %-20s %s\n" "$id" "$idx" "$cal" "$url"
  done <<< "$ids"
  echo "────────────────────────────────────────────"
  echo "Persistent tracker: $P_TRACK"
}

add_calendar_url(){
  ensure_venv
  require_defaults_or_prompt
  load_config
  echo
  read -rp "Enter calendar source (URL): " SRC
  read -rp "Enter a short ID (e.g., lions): " ID
  TMP="/tmp/${ID}.ics"
  log "Fetch & normalize: $SRC -> $TMP"
  "$PY" "$ROOT_DIR/prepare_ics_for_import.py" "$SRC" "$ID" "$TMP"
  log "Opening Outlook import…"
  open -a "Microsoft Outlook" "$TMP" || echo "⚠️  If Outlook didn't open, double-click $TMP"
  echo "➡ In Outlook, choose: \"${CAL_NAME}\" (#${CAL_INDEX})."
  upsert_track "$ID" "$SRC" "$CAL_NAME" "$CAL_INDEX"
  read -rp "Press Enter to continue…" _
}

import_local_file(){
  ensure_venv
  require_defaults_or_prompt
  load_config
  echo
  read -rp "Enter FULL path to .ics file: " FP
  ABS="$("$PY" - <<PY
import os,sys; p=sys.argv[1]; print(os.path.abspath(os.path.expanduser(p)))
PY
"$FP")"
  if [[ ! -f "$ABS" ]]; then
    echo "❌ File not found: $ABS"; read -rp "Press Enter…" _; return
  fi
  read -rp "Enter a short ID (e.g., msu25): " ID
  TMP="/tmp/${ID}.ics"
  SRC="file://${ABS}"
  log "Normalize local: $ABS -> $TMP"
  "$PY" "$ROOT_DIR/prepare_ics_for_import.py" "$SRC" "$ID" "$TMP"
  log "Opening Outlook import…"
  open -a "Microsoft Outlook" "$TMP" || echo "⚠️  If Outlook didn't open, double-click $TMP"
  echo "➡ In Outlook, choose: \"${CAL_NAME}\" (#${CAL_INDEX})."
  upsert_track "$ID" "$SRC" "$CAL_NAME" "$CAL_INDEX"
  read -rp "Press Enter to continue…" _
}

remove_calendar(){
  ensure_venv
  require_defaults_or_prompt
  load_config
  echo "Enter the SRC ID to remove (e.g., lions). Known IDs:"
  get_all_ids | nl -ba
  read -rp "ID or number: " choice
  if [[ "$choice" =~ ^[0-9]+$ ]]; then
    choice="$(get_all_ids | sed -n "${choice}p")"
  fi
  ID="$choice"
  [[ -z "$ID" ]] && { echo "Cancelled."; return; }

  # Find calendar/index to target
  line="$(grep "\"id\":\"$ID\"" "$P_TRACK" | tail -1)"
  [[ -z "$line" ]] && line="$(grep "\"id\":\"$ID\"" "$TRACK_LOCAL" | tail -1)"
  CAL="$(printf "%s" "$line" | sed -n 's/.*"calendar":"\([^"]*\)".*/\1/p')"
  IDX="$(printf "%s" "$line" | sed -n 's/.*"index":\([0-9][0-9]*\).*/\1/p')"
  [[ -z "$CAL" ]] && CAL="$CAL_NAME"
  [[ -z "$IDX" ]] && IDX="$CAL_INDEX"

  log "Deleting [SRC: $ID] from \"$CAL\" (#$IDX)…"
  osascript "$ROOT_DIR/outlook_delete_by_src.applescript" "$ID" "$CAL" "$IDX" \
    && log "Deleted." || echo "⚠️  Delete script reported an error."
  delete_track "$ID"
  read -rp "Press Enter to continue…" _
}

rebuild_tracking(){
  ensure_venv
  require_defaults_or_prompt
  load_config
  echo "Scanning \"$CAL_NAME\" (#$CAL_INDEX) for [SRC: …] tags…"
  json="$(osascript "$ROOT_DIR/outlook_scan_srcs.applescript" "$CAL_NAME" "$CAL_INDEX" || true)"
  echo "$json" | grep -q '"sources"' || { echo "⚠️  No data returned."; read -rp "Enter…" _; return; }
  echo "$json" | sed -n 's/.*"id":"\([^"]\+\)".*"count":\([0-9]\+\).*/\1 \2/p' | while read -r ID CNT; do
    upsert_track "$ID" "" "$CAL_NAME" "$CAL_INDEX"
    echo "  • Rebuilt: $ID (events: $CNT)"
  done
  echo "✅ Rebuild complete. (Persistent file: ${P_TRACK})"
  read -rp "Press Enter to continue…" _
}

menu(){
cat <<MENU
╔═════════ ICS Bridge for Outlook ═════════╗
║ 1) ➕ Add Calendar via URL               ║
║ 2) 📁 Import from Local .ics File        ║
║ 3) 🗑️  Remove Imported Calendar          ║
║ 4) 📋 List Imported Calendars            ║
║ 5) 🔁 Rebuild Tracking (this calendar)   ║
║ 6) 🛠️  Set Default Target Calendar       ║
║ 7) ❌ Quit                                ║
╚══════════════════════════════════════════╝
MENU
  read -rp "Choose [1-7]: " CHOICE
}

main(){
  while true; do
    menu
    case "${CHOICE:-}" in
      1) add_calendar_url ;;
      2) import_local_file ;;
      3) remove_calendar ;;
      4) list_sources; read -rp "Press Enter to continue…" _ ;;
      5) rebuild_tracking ;;
      6) set_defaults ;;
      7) exit 0 ;;
      *) echo "Invalid choice"; sleep 1 ;;
    esac
  done
}

main "$@"
