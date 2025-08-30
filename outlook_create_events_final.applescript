-- outlook_create_events_final.applescript
-- Most compatible version - uses simple string manipulation for dates
-- Usage:
--   osascript ~/icsBridge/outlook_create_events_final.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>

on run argv
  if (count of argv) < 4 then
    return "ERROR: usage: outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>"
  end if

  set jsonPath to item 1 of argv
  set calName to item 2 of argv
  set calIndex to (item 3 of argv) as integer
  set sourceId to item 4 of argv

  -- Simple Python to extract event data with parsed dates
  set pyScript to "import json, sys
from datetime import datetime
p = sys.argv[1]
data = json.load(open(p))
for ev in data.get('events', []):
    uid = ev.get('uid', '').strip()
    if not uid: continue
    summary = ev.get('summary', '(No title)').replace('\\t', ' ').replace('\\n', ' ')[:100]
    location = ev.get('location', '').replace('\\t', ' ').replace('\\n', ' ')[:100]
    desc = ev.get('description', '').replace('\\t', ' ').replace('\\n', ' ')[:500]
    
    # Parse dates
    start_str = ev.get('start', '')
    end_str = ev.get('end', start_str)
    
    # Handle both Z and non-Z formats
    if start_str.endswith('Z'):
        start_dt = datetime.strptime(start_str, '%Y-%m-%dT%H:%M:%SZ')
    else:
        start_dt = datetime.strptime(start_str[:19], '%Y-%m-%dT%H:%M:%S')
    
    if end_str.endswith('Z'):
        end_dt = datetime.strptime(end_str, '%Y-%m-%dT%H:%M:%SZ')
    else:
        end_dt = datetime.strptime(end_str[:19], '%Y-%m-%dT%H:%M:%S')
    
    all_day = '1' if ev.get('all_day', False) else '0'
    
    # Output as tab-separated: summary, location, desc, year, month, day, hour, min, year, month, day, hour, min, all_day, uid
    print(f'{summary}\\t{location}\\t{desc}\\t{start_dt.year}\\t{start_dt.month}\\t{start_dt.day}\\t{start_dt.hour}\\t{start_dt.minute}\\t{end_dt.year}\\t{end_dt.month}\\t{end_dt.day}\\t{end_dt.hour}\\t{end_dt.minute}\\t{all_day}\\t{uid}')"

  try
    set eventData to do shell script "/usr/bin/python3 -c " & quoted form of pyScript & " " & quoted form of jsonPath
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
      if length of eventLine > 0 then
        try
          set eventParts to text items of eventLine
          if (count of eventParts) ≥ 15 then
            set evSummary to item 1 of eventParts
            set evLocation to item 2 of eventParts
            set evDescription to item 3 of eventParts
            
            -- Start date components
            set startYear to (item 4 of eventParts) as integer
            set startMonth to (item 5 of eventParts) as integer
            set startDay to (item 6 of eventParts) as integer
            set startHour to (item 7 of eventParts) as integer
            set startMin to (item 8 of eventParts) as integer
            
            -- End date components
            set endYear to (item 9 of eventParts) as integer
            set endMonth to (item 10 of eventParts) as integer
            set endDay to (item 11 of eventParts) as integer
            set endHour to (item 12 of eventParts) as integer
            set endMin to (item 13 of eventParts) as integer
            
            set allDayFlag to item 14 of eventParts
            set evUID to item 15 of eventParts
            
            -- Create dates using string format (most compatible)
            set startDateStr to (startMonth as string) & "/" & (startDay as string) & "/" & (startYear as string) & " " & (startHour as string) & ":" & (startMin as string) & ":00"
            set endDateStr to (endMonth as string) & "/" & (endDay as string) & "/" & (endYear as string) & " " & (endHour as string) & ":" & (endMin as string) & ":00"
            
            set startDate to date startDateStr
            set endDate to date endDateStr
            
            -- Build notes with tags
            set noteContent to evDescription & return & return & "[SRC: " & sourceId & "]" & return & "[ICSUID: " & evUID & "]"
            
            -- Create event using the pattern that works
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
        on error errMsg
          set failedCount to failedCount + 1
        end try
      end if
    end repeat

    set AppleScript's text item delimiters to oldDelims
    return "✅ Created " & createdCount & " events, " & failedCount & " failed"
  end tell
end run
