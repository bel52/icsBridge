-- outlook_create_events_v2.applescript
-- Alternative version using temp file for Python script
-- Usage:
--   osascript ~/icsBridge/outlook_create_events_v2.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>

on run argv
  if (count of argv) < 4 then
    return "ERROR: usage: outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>"
  end if

  set jsonPath to item 1 of argv
  set calName to item 2 of argv
  set calIndex to (item 3 of argv) as integer
  set sourceId to item 4 of argv

  -- Write Python script to temp file
  set pyScript to "/tmp/parse_events_" & sourceId & ".py"
  set pyContent to "#!/usr/bin/env python3
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
  
  -- Write Python script to file
  do shell script "echo " & quoted form of pyContent & " > " & pyScript
  do shell script "chmod +x " & pyScript
  
  -- Execute Python script
  try
    set rows to do shell script "/usr/bin/python3 " & pyScript & " " & quoted form of jsonPath
  on error errMsg
    do shell script "rm -f " & pyScript
    return "ERROR running python: " & errMsg
  end try
  
  -- Clean up Python script
  do shell script "rm -f " & pyScript

  tell application "Microsoft Outlook"
    -- Find target calendar
    set calMatches to (calendars whose name is calName)
    if (count of calMatches) < calIndex then
      return "ERROR: Calendar \"" & calName & "\" (#" & calIndex & ") not found."
    end if
    set targetCal to item calIndex of calMatches

    set createdCount to 0
    set failedCount to 0
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to tab

    repeat with rowLine in paragraphs of rows
      if rowLine is not "" then
        set cols to text items of rowLine
        if (count of cols) ≥ 17 then
          try
            -- Parse all fields
            set evSummary to item 1 of cols
            set evLoc to item 2 of cols
            set evDesc to item 3 of cols
            
            set startYear to (item 4 of cols) as integer
            set startMonth to (item 5 of cols) as integer
            set startDay to (item 6 of cols) as integer
            set startHour to (item 7 of cols) as integer
            set startMin to (item 8 of cols) as integer
            set startSec to (item 9 of cols) as integer
            
            set endYear to (item 10 of cols) as integer
            set endMonth to (item 11 of cols) as integer
            set endDay to (item 12 of cols) as integer
            set endHour to (item 13 of cols) as integer
            set endMin to (item 14 of cols) as integer
            set endSec to (item 15 of cols) as integer
            
            set isAllDayText to item 16 of cols
            set evUID to item 17 of cols
            
            -- Create start date
            set startDate to date ("1/1/2000 12:00:00 AM")
            set year of startDate to startYear
            set month of startDate to startMonth
            set day of startDate to startDay
            set time of startDate to (startHour * 3600 + startMin * 60 + startSec)
            
            -- Create end date
            set endDate to date ("1/1/2000 12:00:00 AM")
            set year of endDate to endYear
            set month of endDate to endMonth
            set day of endDate to endDay
            set time of endDate to (endHour * 3600 + endMin * 60 + endSec)
            
            -- Build notes
            set notesText to evDesc & return & return & "[SRC: " & sourceId & "]" & return & "[ICSUID: " & evUID & "]"
            
            -- Create event
            if isAllDayText = "1" then
              set newEv to make new calendar event at targetCal with properties {subject:evSummary, all day event:true}
            else
              set newEv to make new calendar event at targetCal with properties {subject:evSummary, all day event:false}
            end if
            
            -- Set remaining properties
            set location of newEv to evLoc
            set content of newEv to notesText
            set start time of newEv to startDate
            set end time of newEv to endDate
            
            set createdCount to createdCount + 1
          on error
            set failedCount to failedCount + 1
          end try
        end if
      end if
    end repeat

    set AppleScript's text item delimiters to oldDelims
    return "OK created=" & createdCount & " failed=" & failedCount
  end tell
end run
