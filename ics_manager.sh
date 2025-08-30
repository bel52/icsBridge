#!/usr/bin/env bash
# ICS Bridge Manager - Import ICS/JSON calendars into Outlook
# Supports: http://, https://, webcal://, local files, ICS and JSON formats
set -euo pipefail

# ===========================
# CONFIGURATION
# ===========================
if [[ -x "$HOME/icsBridge/.venv/bin/python3" ]]; then
  PYTHON="$HOME/icsBridge/.venv/bin/python3"
else
  PYTHON="python3"
fi

CALENDAR_NAME="Calendar"                     # Default Outlook calendar name
CALENDAR_INDEX=2                             # Default occurrence index
DEFAULT_CATEGORY="Imported"                  # Category label
SOURCES_FILE="$HOME/icsBridge/sources.json"  # Tracking file
LOG_DIR="$HOME/icsBridge/logs"

mkdir -p "$(dirname "$SOURCES_FILE")" "$LOG_DIR"

# Initialize sources.json if missing
if [[ ! -f "$SOURCES_FILE" ]]; then
  echo '{}' > "$SOURCES_FILE"
fi

timestamp() { date +"%Y-%m-%dT%H:%M:%S%z"; }

# Sanitize source ID (replace spaces and special chars)
sanitize_id() {
  echo "$1" | sed 's/[^a-zA-Z0-9_-]/_/g' | sed 's/__*/_/g'
}

# Validate calendar source
validate_source() {
  local src="$1"
  local content=""
  
  # Handle different source types
  if [[ "$src" =~ ^https?:// ]] || [[ "$src" =~ ^webcal:// ]]; then
    # For webcal, convert to https for validation
    local test_url="$src"
    if [[ "$src" =~ ^webcal:// ]]; then
      test_url="https://${src:9}"
    fi
    
    # Try to fetch first few lines
    if command -v curl >/dev/null 2>&1; then
      content="$(curl -fsL --max-time 15 "$test_url" 2>/dev/null | head -n 100 || true)"
    else
      content="$("$PYTHON" -c "
import urllib.request, ssl, sys
url = sys.argv[1]
if url.startswith('webcal://'):
    url = 'https://' + url[9:]
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
try:
    with urllib.request.urlopen(url, context=ctx, timeout=15) as r:
        print(r.read(4000).decode('utf-8', errors='ignore'))
except Exception as e:
    print(f'Error: {e}', file=sys.stderr)
    sys.exit(1)
" "$src" 2>/dev/null || true)"
    fi
    
    if [[ -z "${content:-}" ]]; then
      echo "WARNING: Could not fetch URL for validation (may still work)" >&2
      return 0  # Don't fail, let the Python script handle it
    fi
  elif [[ -f "$src" ]]; then
    content="$(head -n 100 "$src" 2>/dev/null || true)"
  else
    echo "ERROR: Source is not a valid URL or file: $src" >&2
    return 1
  fi
  
  # Check if it's a valid format (ICS or JSON)
  if echo "$content" | grep -q "BEGIN:VCALENDAR"; then
    echo "✓ Detected ICS/iCal format" >&2
    return 0
  elif echo "$content" | grep -q '^\s*[\[{]'; then
    echo "✓ Detected JSON format" >&2
    return 0
  elif [[ -z "$content" ]]; then
    echo "✓ Will attempt to fetch and parse" >&2
    return 0
  else
    echo "WARNING: Could not detect format (will try to parse anyway)" >&2
    return 0
  fi
}

# ===== JSON helpers =====
json_keys() {
  "$PYTHON" - "$SOURCES_FILE" <<'PY'
import json,sys
try: d=json.load(open(sys.argv[1]))
except: d={}
for k in sorted(d.keys()): print(k)
PY
}

json_get_field() {
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
  local key="$1" src="$2" cal="$3" idx="$4" cat="$5" ts="$6"
  "$PYTHON" - "$SOURCES_FILE" "$key" "$src" "$cal" "$idx" "$cat" "$ts" <<'PY'
import json,sys
p,key,src,cal,idx,cat,ts=sys.argv[1:7]
try: d=json.load(open(p))
except: d={}
try: idx=int(idx)
except: idx=1
d[key]={"source":src,"calendar":cal,"calendar_index":idx,"category":cat,"last_import":ts}
with open(p,"w") as f: json.dump(d,f,indent=2,ensure_ascii=False)
PY
}

json_delete_key() {
  local key="$1"
  "$PYTHON" - "$SOURCES_FILE" "$key" <<'PY'
import json,sys
p,key=sys.argv[1:3]
try: d=json.load(open(p))
except: d={}
if key in d: del d[key]
with open(p,"w") as f: json.dump(d,f,indent=2,ensure_ascii=False)
PY
}

# List tracked sources
list_sources() {
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then
    echo "No tracked calendars yet."
    return
  fi
  echo "═══════════════════════════════════════════════════════════"
  echo "Tracked Calendar Sources:"
  echo "═══════════════════════════════════════════════════════════"
  local i=1
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    local src last cal idx
    src="$(json_get_field "$k" "source")"
    # Fallback for old format
    [[ -z "$src" ]] && src="$(json_get_field "$k" "ics")"
    last="$(json_get_field "$k" "last_import")"
    cal="$(json_get_field "$k" "calendar")"
    idx="$(json_get_field "$k" "calendar_index")"
    [[ -z "$last" ]] && last="never"
    
    echo "$i) ID: $k"
    echo "   Source: $src"
    echo "   Target: \"$cal\" (#$idx)"
    echo "   Last Import: $last"
    echo "───────────────────────────────────────────────────────────"
    i=$((i+1))
  done <<< "$keys"
}

# Interactive source selection
select_source_interactive() {
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then
    echo ""
    return 0
  fi
  
  local arr=()
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    arr[${#arr[@]}]="$k"
  done <<< "$keys"
  
  local count="${#arr[@]}"
  echo
  list_sources
  echo
  printf "Enter number to select (or 'q' to cancel): "
  read choice
  
  if [[ "$choice" =~ ^[Qq]$ ]] || [[ -z "$choice" ]]; then
    echo ""
    return 0
  fi
  
  if ! [[ "$choice" =~ ^[0-9]+$ ]] || (( choice < 1 || choice > count )); then
    echo ""
    return 0
  fi
  
  echo "${arr[$((choice-1))]}"
}

# Add/Update calendar
add_calendar() {
  echo
  echo "═══════════════════════════════════════════════════════════"
  echo "                    Add Calendar Source                     "
  echo "═══════════════════════════════════════════════════════════"
  echo "Supported formats:"
  echo "  • ICS/iCal URLs (http://, https://, webcal://)"
  echo "  • JSON calendar feeds"
  echo "  • Local .ics, .ical, or .json files"
  echo "───────────────────────────────────────────────────────────"
  echo
  printf "Enter calendar source (URL or file path): "
  read SOURCE
  if [[ -z "${SOURCE:-}" ]]; then
    echo "❌ No source provided."
    return
  fi
  
  # Basic validation
  if ! validate_source "$SOURCE"; then
    printf "⚠️  Validation warning. Continue anyway? (y/N): "
    read confirm
    if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
      echo "Aborted."
      return
    fi
  fi
  
  printf "Enter a short ID for this calendar (e.g., lions-2025): "
  read SRC_RAW
  if [[ -z "${SRC_RAW:-}" ]]; then
    echo "❌ No ID provided."
    return
  fi
  
  # Sanitize the source ID
  SRC=$(sanitize_id "$SRC_RAW")
  if [[ "$SRC" != "$SRC_RAW" ]]; then
    echo "📝 ID sanitized to: $SRC"
  fi
  
  # Check if already exists
  local existing_src
  existing_src="$(json_get_field "$SRC" "source")"
  [[ -z "$existing_src" ]] && existing_src="$(json_get_field "$SRC" "ics")"
  
  if [[ -n "$existing_src" ]]; then
    echo "⚠️  Source '$SRC' already exists (last source: $existing_src)"
    printf "Update/refresh it? (y/N): "
    read confirm
    if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
      echo "Aborted."
      return
    fi
  fi
  
  # Get calendar details
  printf "Outlook calendar name [%s]: " "$CALENDAR_NAME"
  read CAL_IN || true
  printf "Occurrence index [%s]: " "$CALENDAR_INDEX"
  read IDX_IN || true
  printf "Category label [%s]: " "$DEFAULT_CATEGORY"
  read CAT_IN || true
  
  local cal_name cal_idx category
  cal_name="${CAL_IN:-$CALENDAR_NAME}"
  cal_idx="${IDX_IN:-$CALENDAR_INDEX}"
  category="${CAT_IN:-$DEFAULT_CATEGORY}"
  
  # Process the source
  local json_out log_file
  json_out="/tmp/${SRC}_events.json"
  log_file="$LOG_DIR/import_${SRC}_$(date +%Y%m%d_%H%M%S).log"
  
  echo
  echo "🔄 Fetching and parsing calendar data..."
  set +e
  "$PYTHON" "$HOME/icsBridge/fetch_public_ics.py" "$SOURCE" "$json_out" 2>&1 | tee -a "$log_file"
  local py_rc=$?
  set -e
  
  if [[ $py_rc -ne 0 ]]; then
    echo "❌ Failed to parse calendar source. See $log_file"
    return
  fi
  
  # Check if we got any events
  local event_count
  event_count=$("$PYTHON" -c "import json; print(len(json.load(open('$json_out'))['events']))" 2>/dev/null || echo "0")
  
  if [[ "$event_count" == "0" ]]; then
    echo "⚠️  No events found in the source."
    printf "Continue anyway? (y/N): "
    read confirm
    if [[ ! "$confirm" =~ ^[Yy]$ ]]; then
      return
    fi
  fi
  
  echo "📝 Writing $event_count events to Outlook: \"${cal_name}\" (#${cal_idx})"
  
  # Use the working AppleScript
  local applescript_file="$HOME/icsBridge/outlook_create_events.applescript"
  
  set +e
  local result
  result=$(osascript "$applescript_file" "$json_out" "$cal_name" "$cal_idx" "$SRC" 2>&1)
  local osa_rc=$?
  set -e
  echo "$result" | tee -a "$log_file"
  
  if [[ $osa_rc -ne 0 ]] || echo "$result" | grep -qi "error"; then
    echo
    echo "❌ Import failed. Log saved to: $log_file"
    echo
    echo "🔧 Troubleshooting steps:"
    echo "1. Ensure Legacy Outlook: Help → Revert to Legacy Outlook"
    echo "2. Grant permissions: System Settings → Privacy & Security → Automation"
    echo "   → Enable Terminal/iTerm for Microsoft Outlook"
    echo "3. Restart Terminal and Outlook"
    echo "4. Test with: osascript -e 'tell application \"Microsoft Outlook\" to count calendars'"
    return
  fi
  
  # Save to tracking
  local ts
  ts=$(timestamp)
  json_add_or_update "$SRC" "$SOURCE" "$cal_name" "$cal_idx" "$category" "$ts"
  echo
  echo "✅ Successfully imported and tracked as \"$SRC\""
}

# Remove calendar events
remove_calendar() {
  echo
  echo "═══════════════════════════════════════════════════════════"
  echo "                   Remove Calendar Events                   "
  echo "═══════════════════════════════════════════════════════════"
  
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
  
  echo "🗑️  Removing events for \"$chosen\" from \"${cal_name}\" (#${cal_idx})..."
  
  local result
  result=$(osascript -l JavaScript "$HOME/icsBridge/outlook_remove_source.js" \
    "$cal_name" "$cal_idx" "$chosen" 2>&1 || true)
  echo "$result"
  
  if echo "$result" | grep -q '"ok":true'; then
    json_delete_key "$chosen"
    echo "✅ Removed \"$chosen\" from tracking"
  else
    echo "⚠️  Some events may not have been removed"
  fi
}

# Refresh all sources
refresh_all() {
  echo
  echo "═══════════════════════════════════════════════════════════"
  echo "                    Refresh All Sources                     "
  echo "═══════════════════════════════════════════════════════════"
  
  local keys; keys="$(json_keys)"
  if [[ -z "${keys:-}" ]]; then
    echo "No sources to refresh."
    return
  fi
  
  while IFS= read -r k; do
    [[ -z "$k" ]] && continue
    
    local src cal_name cal_idx
    src="$(json_get_field "$k" "source")"
    [[ -z "$src" ]] && src="$(json_get_field "$k" "ics")"
    cal_name="$(json_get_field "$k" "calendar")"
    cal_idx="$(json_get_field "$k" "calendar_index")"
    
    echo "───────────────────────────────────────────────────────────"
    echo "Refreshing: $k"
    echo "Source: $src"
    
    local json_out="/tmp/${k}_events.json"
    
    if "$PYTHON" "$HOME/icsBridge/fetch_public_ics.py" "$src" "$json_out" 2>/dev/null; then
      osascript "$HOME/icsBridge/outlook_create_events.applescript" \
        "$json_out" "$cal_name" "$cal_idx" "$k" 2>/dev/null && \
        echo "✅ Refreshed $k" || echo "❌ Failed to update Outlook"
    else
      echo "❌ Failed to fetch source"
    fi
  done <<< "$keys"
  
  echo "═══════════════════════════════════════════════════════════"
  echo "Refresh complete!"
}

# Main menu
main_menu() {
  while true; do
    echo
    echo "╔═══════════════════════════════════════════════════════╗"
    echo "║            📅 ICS Bridge for Outlook                  ║"
    echo "╠═══════════════════════════════════════════════════════╣"
    echo "║  1) ➕ Add/Update calendar source                     ║"
    echo "║  2) 🗑️  Remove calendar events                        ║"
    echo "║  3) 📋 List tracked sources                           ║"
    echo "║  4) 🔄 Refresh all sources                            ║"
    echo "║  5) ❌ Quit                                           ║"
    echo "╚═══════════════════════════════════════════════════════╝"
    printf "Choose [1-5]: "
    read opt
    
    case "$opt" in
      1) add_calendar ;;
      2) remove_calendar ;;
      3) echo; list_sources ;;
      4) refresh_all ;;
      5) echo "👋 Goodbye!"; exit 0 ;;
      *) echo "Invalid option. Please choose 1-5." ;;
    esac
  done
}

# Handle command line args
case "${1:-}" in
  add|import) add_calendar ;;
  remove|delete) remove_calendar ;;
  list|show) list_sources ;;
  refresh|update) refresh_all ;;
  *) main_menu ;;
esac
