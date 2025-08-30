#!/usr/bin/env bash
# Compatible with macOS Bash 3.2 (no mapfile, no jq required)
set -euo pipefail

# ===========================
# CONFIGURATION (edit freely)
# ===========================
if [[ -x "$HOME/icsBridge/.venv/bin/python3" ]]; then
  PYTHON="$HOME/icsBridge/.venv/bin/python3"
else
  PYTHON="python3"
fi

CALENDAR_NAME="Calendar"                     # Outlook calendar name to target
CALENDAR_INDEX=2                             # Occurrence of CALENDAR_NAME (2 = the second one)
DEFAULT_CATEGORY="Imported: Sports"          # (ignored by AppleScript writer)
SOURCES_FILE="$HOME/icsBridge/sources.json"  # Where we track installed sources
LOG_DIR="$HOME/icsBridge/logs"

mkdir -p "$(dirname "$SOURCES_FILE")" "$LOG_DIR"

# Initialize sources.json if missing
if [[ ! -f "$SOURCES_FILE" ]]; then
  echo '{}' > "$SOURCES_FILE"
fi

timestamp() { date +"%Y-%m-%dT%H:%M:%S%z"; }

# --- Lightweight validator for URL/file before parsing ---
validate_ical() {
  local src="$1"
  local content=""

  if [[ "$src" =~ ^https?:// ]]; then
    if command -v curl >/dev/null 2>&1; then
      content="$(curl -fsL --max-time 15 "$src" | head -n 40 || true)"
    else
      content="$("$PYTHON" - "$src" 2>/dev/null || true <<'PYCODE'
import sys, urllib.request, ssl
url = sys.argv[1]
ctx = ssl.create_default_context()
with urllib.request.urlopen(url, context=ctx, timeout=15) as r:
    data = r.read(4000).decode('utf-8', errors='ignore')
    print(data)
PYCODE
)"
    fi
    if [[ -z "${content:-}" ]]; then
      echo "ERROR: Could not fetch URL or empty response." >&2
      return 1
    fi
  else
    if [[ ! -f "$src" ]]; then
      echo "ERROR: File does not exist: $src" >&2
      return 1
    fi
    content="$(head -n 40 "$src")"
  fi

  echo "$content" | grep -q "BEGIN:VCALENDAR" || {
    echo "ERROR: Not a valid iCalendar file (missing BEGIN:VCALENDAR)" >&2
    return 1
  }
  if ! echo "$content" | grep -q "BEGIN:VEVENT"; then
    echo "WARNING: No events found (missing BEGIN:VEVENT). File may be empty." >&2
  fi
  return 0
}

# ===== JSON helpers (no jq) =====
json_keys() {
  "$PYTHON" - "$SOURCES_FILE" <<'PY'
import json,sys
p=sys.argv[1]
try:
    d=json.load(open(p))
except:
    d={}
for k in d.keys():
    print(k)
PY
}

json_get_field() {
  # usage: json_get_field KEY FIELD
  local key="$1" field="$2"
  "$PYTHON" - "$SOURCES_FILE" "$key" "$field" <<'PY'
import json,sys
p,k,f=sys.argv[1:4]
try: d=json.load(open(p))
except: d={}
v=d.get(k,{}).get(f,"")
print(v if v is not None else "")
PY
}

json_add_or_update() {
  # usage: json_add_or_update KEY ICS CAL_NAME CAL_IDX CATEGORY LAST_IMPORT
  local key="$1" ics="$2" cal="$3" idx="$4" cat="$5" ts="$6"
  "$PYTHON" - "$SOURCES_FILE" "$key" "$ics" "$cal" "$idx" "$cat" "$ts" <<'PY'
import json,sys
p,key,ics,cal,idx,cat,ts=sys.argv[1:7]
try:
    d=json.load(open(p))
except:
    d={}
try:
    idx=int(idx)
except:
    idx=1
d[key]={"ics":ics,"calendar":cal,"calendar_index":idx,"category":cat,"last_import":ts}
with open(p,"w") as f:
    json.dump(d,f,indent=2,ensure_ascii=False)
PY
}

json_delete_key() {
  # usage: json_delete_key KEY
  local key="$1"
  "$PYTHON" - "$SOURCES_FILE" "$key" <<'PY'
import json,sys
p,key=sys.argv[1:3]
try:
    d=json.load(open(p))
except:
    d={}
if key in d:
    del d[key]
with open(p,"w") as f:
    json.dump(d,f,indent=2,ensure_ascii=False)
PY
}

list_sources() {
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then
    echo "No tracked calendars yet."
    return
  fi
  echo "Tracked calendars:"
  local i=1
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    local ics last
    ics="$(json_get_field "$k"  "ics")"
    last="$(json_get_field "$k" "last_import")"
    [[ -z "$last" ]] && last="never"
    echo "$i) $k  —  $ics  (last_import: $last)"
    i=$((i+1))
  done <<< "$keys"
}

# Returns chosen source id via echo, or empty on cancel/invalid
select_source_interactive() {
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then
    echo ""
    return 0
  fi

  # Build array compatible with Bash 3.2
  local arr=()
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    arr[${#arr[@]}]="$k"
  done <<< "$keys"

  local count="${#arr[@]}"
  echo
  list_sources
  echo
  printf "Enter number to remove (or 'q' to cancel): "
  read choice
  if [[ "$choice" =~ ^[Qq]$ ]]; then
    echo ""
    return 0
  fi
  case "$choice" in
    ''|*[!0-9]*) echo ""; return 0 ;;
  esac
  if (( choice < 1 || choice > count )); then
    echo ""
    return 0
  fi
  echo "${arr[$((choice-1))]}"
}

add_calendar() {
  echo
  printf "Paste path or URL to .ics/.ical file: "
  read ICS
  if [[ -z "${ICS:-}" ]]; then
    echo "Aborted: no ICS provided."
    return
  fi

  # ✅ Validate before proceeding
  if ! validate_ical "$ICS"; then
    echo "Validation failed. Aborting import."
    return
  fi

  printf "Enter short source ID (e.g., detroit-lions-2025): "
  read SRC
  if [[ -z "${SRC:-}" ]]; then
    echo "Aborted: no source ID provided."
    return
  fi

  # Optional: override defaults per import (press Enter to accept defaults)
  printf "Outlook calendar name [%s]: " "$CALENDAR_NAME"
  read CAL_IN || true
  printf "Occurrence index for that name [%s]: " "$CALENDAR_INDEX"
  read IDX_IN || true
  printf "Category label [%s]: " "$DEFAULT_CATEGORY"
  read CAT_IN || true

  local cal_name cal_idx category
  cal_name="${CAL_IN:-$CALENDAR_NAME}"
  cal_idx="${IDX_IN:-$CALENDAR_INDEX}"
  category="${CAT_IN:-$DEFAULT_CATEGORY}"  # (ignored by AppleScript writer)

  local json_out log_file
  json_out="/tmp/${SRC}_events.json"
  log_file="$LOG_DIR/import_${SRC}_$(date +%Y%m%d_%H%M%S).log"

  echo
  echo "[*] Parsing ICS → JSON …"
  set +e
  "$PYTHON" "$HOME/icsBridge/fetch_public_ics.py" "$ICS" "$json_out" | tee -a "$log_file"
  local py_rc=$?
  set -e
  if [[ $py_rc -ne 0 ]]; then
    echo "Failed to parse ICS. See $log_file"
    return
  fi

  echo "[*] Writing to Outlook (AppleScript): \"${cal_name}\" (#${cal_idx})"
  set +e
  local result
  result=$(osascript "$HOME/icsBridge/outlook_create_events.applescript" \
    "$json_out" "$cal_name" "$cal_idx" "$SRC")
  local osa_rc=$?
  set -e
  echo "$result" | tee -a "$log_file"

  if [[ $osa_rc -ne 0 ]] || echo "$result" | grep -q "^ERROR"; then
    echo "osascript returned an error. See $log_file"
    return
  fi

  local ts
  ts=$(timestamp)
  json_add_or_update "$SRC" "$ICS" "$cal_name" "$cal_idx" "$category" "$ts"
  echo
  echo "[+] Imported and tracked as \"$SRC\"."
}

remove_calendar() {
  echo
  local chosen
  chosen="$(select_source_interactive)"
  if [[ -z "$chosen" ]]; then
    echo "Cancelled."
    return
  fi

  local cal_name cal_idx
  cal_name="$(json_get_field "$chosen" "calendar")"
  cal_idx="$(json_get_field "$chosen" "calendar_index")"
  [[ -z "$cal_name" ]] && cal_name="$CALENDAR_NAME"
  [[ -z "$cal_idx"  ]] && cal_idx="$CALENDAR_INDEX"

  echo "[*] Removing all events for source \"$chosen\" from \"${cal_name}\" (#${cal_idx}) …"
  local result
  result=$(osascript -l JavaScript "$HOME/icsBridge/outlook_remove_source.js" \
    "$cal_name" "$cal_idx" "$chosen" || true)
  echo "$result"

  echo "$result" | grep -q '"ok":true' && {
    json_delete_key "$chosen"
    echo "[-] Removed \"$chosen\" from tracking."
  }
}

main_menu() {
  while true; do
    echo
    echo "=== ICS → Outlook Manager ==="
    echo "1) Add calendar (import/refresh)"
    echo "2) Remove calendar (delete all events)"
    echo "3) List tracked calendars"
    echo "4) Quit"
    printf "Choose [1-4]: "
    read opt
    case "$opt" in
      1) add_calendar ;;
      2) remove_calendar ;;
      3) echo; list_sources ;;
      4) echo "Bye!"; exit 0 ;;
      *) echo "Invalid option." ;;
    esac
  done
}

case "${1:-}" in
  import)  add_calendar ;;
  remove)  remove_calendar ;;
  list)    list_sources ;;
  *)       main_menu ;;
esac
