#!/usr/bin/env bash
set -euo pipefail

PYTHON="$HOME/icsBridge/.venv/bin/python3"
SOURCES_FILE="$HOME/icsBridge/sources.json"
LOG_DIR="$HOME/icsBridge/logs"

mkdir -p "$(dirname "$SOURCES_FILE")" "$LOG_DIR"
if [[ ! -f "$SOURCES_FILE" ]]; then echo '{}' > "$SOURCES_FILE"; fi

# ===== JSON helpers =====
json_keys() { "$PYTHON" -c 'import json,sys; [print(k) for k in sorted(json.load(open(sys.argv[1])).keys())]' "$SOURCES_FILE"; }
json_get_field() { "$PYTHON" -c 'import json,sys; print(json.load(open(sys.argv[1])).get(sys.argv[2], {}).get(sys.argv[3], ""))' "$SOURCES_FILE" "$1" "$2"; }
json_add_or_update() {
  local key="$1" src="$2" cal="$3" idx="$4"
  "$PYTHON" -c 'import json,sys; p,key,src,cal,idx=sys.argv[1:6]; d=json.load(open(p)); d[key]={"source":src,"calendar":cal,"calendar_index":int(idx)}; json.dump(d,open(p,"w"),indent=2)' "$SOURCES_FILE" "$key" "$src" "$cal" "$idx"
}
json_delete_key() { "$PYTHON" -c 'import json,sys; p,key=sys.argv[1:3]; d=json.load(open(p)); d.pop(key, None); json.dump(d,open(p,"w"),indent=2)' "$SOURCES_FILE" "$1"; }

# List tracked sources
list_sources() {
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then echo "No tracked calendars yet."; return; fi
  echo "═════════ Tracked Sources ═════════"
  local i=1
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    echo "$i) ID: $k"
    echo "   Source: $(json_get_field "$k" "source")"
    echo "   Target Calendar: \"$(json_get_field "$k" "calendar")\" (#$(json_get_field "$k" "calendar_index"))"
    echo "───────────────────────────────────"
    i=$((i+1))
  done <<< "$keys"
}

# Add/Update calendar
add_calendar() {
  echo; echo "═════════ Add Calendar Source ═════════"
  printf "Enter calendar source (URL): "
  read SOURCE
  if [[ -z "${SOURCE:-}" ]]; then echo "❌ No source provided."; return; fi
  
  printf "Enter a short ID for this calendar (e.g., lions-2025): "
  read SRC_ID
  if [[ -z "${SRC_ID:-}" ]]; then echo "❌ No ID provided."; return; fi
  
  printf "Enter target Outlook calendar name (for tracking): "
  read CAL_NAME
  if [[ -z "${CAL_NAME:-}" ]]; then echo "❌ No calendar name provided."; return; fi

  printf "Enter occurrence index for that name [2]: "
  read CAL_INDEX
  CAL_INDEX=${CAL_INDEX:-2}

  local ics_out="/tmp/${SRC_ID}.ics"
  
  echo; echo "🔄 Preparing tagged ICS file..."
  
  set +e
  local result
  result=$("$PYTHON" "$HOME/icsBridge/prepare_ics_for_import.py" "$SOURCE" "$SRC_ID" "$ics_out" 2>&1)
  local py_rc=$?
  set -e
  echo "$result"

  if [[ $py_rc -ne 0 ]]; then echo "❌ Failed to process calendar."; return; fi
  
  json_add_or_update "$SRC_ID" "$SOURCE" "$CAL_NAME" "$CAL_INDEX"
  
  echo; echo "✅ ICS file is ready. Outlook's import dialog will now open."
  echo "Please select the calendar \"$CAL_NAME\" (#$CAL_INDEX) in the dialog."
  
  open -a "Microsoft Outlook" "$ics_out"
  echo; echo "✨ Import process initiated."
}

# Remove calendar events
remove_calendar() {
  echo; echo "════════ Remove Calendar Events ════════"
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then echo "No tracked calendars to remove."; return; fi
  
  local arr=(); while IFS= read -r k; do [[ -z "$k" ]] || arr+=("$k"); done <<< "$keys"
  list_sources
  printf "Enter number to remove (or 'q' to cancel): "; read choice
  if [[ "$choice" =~ ^[Qq]$ ]] || [[ -z "$choice" ]]; then echo "Cancelled."; return; fi
  
  local chosen_id="${arr[$((choice-1))]}"
  local cal_name="$(json_get_field "$chosen_id" "calendar")"
  local cal_idx="$(json_get_field "$chosen_id" "calendar_index")"
  
  echo; echo "🗑️  Removing events for '$chosen_id' from \"$cal_name\" (#$cal_idx)..."
  
  local result
  result=$(osascript -l JavaScript "$HOME/icsBridge/outlook_remove_source.js" "$cal_name" "$cal_idx" "$chosen_id" 2>&1 || true)
  echo "$result"
  
  if echo "$result" | grep -q '"ok":true'; then
    json_delete_key "$chosen_id"
    echo "✅ Removed '$chosen_id' from tracking."
  else
    echo "⚠️  Removal failed or no events were found with that tag."
  fi
}

# Main menu
main_menu() {
  while true; do
    echo
    echo "╔═════════ ICS Bridge for Outlook ═════════╗"
    echo "║ 1) ➕ Add Calendar via Outlook Import    ║"
    echo "║ 2) 🗑️  Remove Imported Calendar         ║"
    echo "║ 3) 📋 List Imported Calendars          ║"
    echo "║ 4) ❌ Quit                               ║"
    echo "╚══════════════════════════════════════════╝"
    printf "Choose [1-4]: "; read opt
    case "$opt" in
      1) add_calendar ;;
      2) remove_calendar ;;
      3) echo; list_sources ;;
      4) echo "👋 Goodbye!"; exit 0 ;;
      *) echo "Invalid option." ;;
    esac
  done
}

main_menu
