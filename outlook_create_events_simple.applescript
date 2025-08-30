-- outlook_create_events_simple.applescript
-- Simplified version that matches the working test pattern
-- Usage:
--   osascript ~/icsBridge/outlook_create_events_simple.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>

on run argv
  if (count of argv) < 4 then
    return "ERROR: usage: outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>"
  end if

  set jsonPath to item 1 of argv
  set calName to item 2 of argv
  set calIndex to (item 3 of argv) as integer
  set sourceId to item 4 of argv

  -- Read and parse JSON using Python
  set pyCmd to "/usr/bin/python3 -c 'import json, sys; import datetime
p = sys.argv[1]
data = json.load(open(p))
for ev in data.get(\"events\", []):
    uid = ev.get(\"uid\", \"\").strip()
    if not uid: continue
    summary = ev.get(\"summary\", \"(No title)\").replace(\"\t\", \" \").replace(\"\n\", \" \")
    location = ev.get(\"location\", \"\").replace(\"\t\", \" \").replace(\"\n\", \" \")
    desc = ev.get(\"description\", \"\").replace(\"\t\", \" \").replace(\"\n\", \" \")
    start = ev.get(\"start\", \"\")
    end = ev.get(\"end\", start)
    all_day = \"1\" if ev.get(\"all_day\", False) else \"0\"
    print(f\"{summary}\t{location}\t{desc}\t{start}\t{end}\t{all_day}\t{uid}\")
' " & quoted form of jsonPath

  try
    set eventData to do shell script pyCmd
  on error errMsg
    return "ERROR parsing JSON: " & errMsg
  end try

  tell application "Microsoft Outlook"
    -- Find target calendar
    set allCals to calendars whose name is calName
    if (count of allCals) < calIndex then
      return "ERROR: Calendar \"" & calName & "\" (#" & calIndex & ") not found."
    end if
    set targetCal to item calIndex of allCals

    set createdCount to 0
    set failedCount to 0
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to tab

    repeat with eventLine in paragraphs of eventData
      if eventLine is not "" then
        try
          set eventParts to text items of eventLine
          if (count of eventParts) ≥ 7 then
            set evSummary to item 1 of eventParts
            set evLocation to item 2 of eventParts
            set evDescription to item 3 of eventParts
            set startISO to item 4 of eventParts
            set endISO to item 5 of eventParts
            set allDayFlag to item 6 of eventParts
            set evUID to item 7 of eventParts
            
            -- Parse ISO dates using shell (more reliable)
            set startTimestamp to do shell script "date -j -f '%Y-%m-%dT%H:%M:%S' '" & (text 1 thru 19 of startISO) & "' '+%s' 2>/dev/null || date -j -f '%Y-%m-%dT%H:%M:%SZ' '" & startISO & "' '+%s'"
            set endTimestamp to do shell script "date -j -f '%Y-%m-%dT%H:%M:%S' '" & (text 1 thru 19 of endISO) & "' '+%s' 2>/dev/null || date -j -f '%Y-%m-%dT%H:%M:%SZ' '" & endISO & "' '+%s'"
            
            -- Convert timestamps to AppleScript dates
            set startDate to (date "Monday, January 1, 2001 12:00:00 AM") + (startTimestamp as integer)
            set endDate to (date "Monday, January 1, 2001 12:00:00 AM") + (endTimestamp as integer)
            
            -- Build notes with tags
            set noteContent to evDescription & return & return & "[SRC: " & sourceId & "]" & return & "[ICSUID: " & evUID & "]"
            
            -- Create event (following the pattern that works)
            set newEv to make new calendar event at targetCal
            tell newEv
              set subject to evSummary
              set location to evLocation
              set content to noteContent
              set start time to startDate
              set end time to endDate
              if allDayFlag = "1" then
                set all day event to true
              else
                set all day event to false
              end if
            end tell
            
            set createdCount to createdCount + 1
          end if
        on error
          set failedCount to failedCount + 1
        end try
      end if
    end repeat

    set AppleScript's text item delimiters to oldDelims
    return "Created " & createdCount & " events, " & failedCount & " failed"
  end tell
end run
