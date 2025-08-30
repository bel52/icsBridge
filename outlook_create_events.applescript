-- outlook_create_events.applescript
-- Usage:
--   osascript ~/icsBridge/outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>

on run argv
  if (count of argv) < 4 then
    return "ERROR: usage: outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>"
  end if

  set jsonPath to item 1 of argv
  set calName to item 2 of argv
  set calIndex to (item 3 of argv) as integer
  set sourceId to item 4 of argv

  -- Python prints TSV rows: summary, location, description,
  -- y, m, d, hh, mi, ss, yE, mE, dE, hhE, miE, ssE, all_day(0/1), uid
  set py to "
import json, sys
from datetime import datetime, timezone

def parse_iso(s):
    if s.endswith('Z'):
        return datetime.strptime(s, '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc).astimezone()
    return datetime.strptime(s, '%Y-%m-%dT%H:%M:%S')

def esc(s):
    return (s or '').replace('\\t',' ').replace('\\r',' ').replace('\\n',' ')

p = sys.argv[1]
events = json.load(open(p)).get('events', [])
for ev in events:
    uid = (ev.get('uid') or '').strip()
    if not uid:
        continue
    s = ev.get('start'); e = ev.get('end') or s
    sd = parse_iso(s); ed = parse_iso(e)
    print('\\t'.join([
        esc(ev.get('summary','(No title)')),
        esc(ev.get('location','')),
        esc(ev.get('description','')),
        str(sd.year), str(sd.month), str(sd.day), str(sd.hour), str(sd.minute), str(sd.second),
        str(ed.year), str(ed.month), str(ed.day), str(ed.hour), str(ed.minute), str(ed.second),
        '1' if ev.get('all_day', False) else '0',
        (ev.get('uid') or '').strip()
    ]))
"

  try
    set rows to do shell script "/usr/bin/python3 - <<'PY'\n" & py & "\nPY\n" & space & quoted form of jsonPath
  on error errMsg
    return "ERROR running python: " & errMsg
  end try

  tell application "Microsoft Outlook"
    activate

    -- Match calendars whose name is calName, then pick the Nth
    set calMatches to {}
    try
      set calMatches to (calendars whose name is calName)
    end try
    if (count of calMatches) < calIndex then
      return "ERROR: Calendar \"" & calName & "\" (#" & calIndex & ") not found."
    end if
    set targetCal to item calIndex of calMatches

    -- Month enum table for component date building
    set monthTable to {January, February, March, April, May, June, July, August, September, October, November, December}

    set AppleScript's text item delimiters to tab
    set createdCount to 0
    set failedCount to 0

    repeat with L in paragraphs of rows
      set one to (L as string)
      if one is not "" then
        set cols to text items of one
        if (count of cols) ≥ 17 then
          set evSummary to item 1 of cols
          set evLoc to item 2 of cols
          set evDesc to item 3 of cols

          set y to (item 4 of cols) as integer
          set mo to (item 5 of cols) as integer
          set dy to (item 6 of cols) as integer
          set hh to (item 7 of cols) as integer
          set mi to (item 8 of cols) as integer
          set ss to (item 9 of cols) as integer

          set yE to (item 10 of cols) as integer
          set moE to (item 11 of cols) as integer
          set dyE to (item 12 of cols) as integer
          set hhE to (item 13 of cols) as integer
          set miE to (item 14 of cols) as integer
          set ssE to (item 15 of cols) as integer

          set isAllDayFlag to item 16 of cols
          set evUID to item 17 of cols

          set startDate to (current date)
          set year of startDate to y
          set month of startDate to (item mo of monthTable)
          set day of startDate to dy
          set time of startDate to (hh * 3600 + mi * 60 + ss)

          set endDate to (current date)
          set year of endDate to yE
          set month of endDate to (item moE of monthTable)
          set day of endDate to dyE
          set time of endDate to (hhE * 3600 + miE * 60 + ssE)

          set notes to evDesc & return & return & "[SRC: " & sourceId & "]" & return & "[ICSUID: " & evUID & "]"
          set isAllDayBool to (isAllDayFlag = "1")

          try
            -- 1) Create with only 'subject' in properties (most compatible)
            set newEv to make new calendar event at targetCal with properties {subject:evSummary}

            -- 2) Set other fields after creation (one per line, no continuations)
            set location of newEv to evLoc
            set content of newEv to notes
            set start time of newEv to startDate
            set end time of newEv to endDate
            set all day event of newEv to isAllDayBool

            set createdCount to createdCount + 1
          on error errM
            set failedCount to failedCount + 1
          end try
        end if
      end if
    end repeat

    set AppleScript's text item delimiters to ""
    return "OK created=" & createdCount & " failed=" & failedCount
  end tell
end run
