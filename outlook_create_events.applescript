-- outlook_create_events.applescript
-- Working version that avoids parsing issues
-- Usage: osascript outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>

on run argv
	if (count of argv) < 4 then
		return "ERROR: usage: outlook_create_events.applescript <jsonPath> <calendarName> <occurrenceIndex> <sourceId>"
	end if
	
	set jsonPath to item 1 of argv
	set calName to item 2 of argv
	set calIndex to (item 3 of argv) as integer
	set sourceId to item 4 of argv
	
	-- Extract data using simple Python script in temp file
	set pyFile to "/tmp/extract_" & sourceId & ".py"
	set pyScript to "#!/usr/bin/env python3
import json
import sys
from datetime import datetime

data = json.load(open(sys.argv[1]))
events = data.get('events', [])

for ev in events:
    uid = ev.get('uid', '').strip()
    if not uid:
        continue
    
    # Clean text fields
    summary = ev.get('summary', '(No title)').replace('\\t', ' ').replace('\\n', ' ')[:100]
    location = ev.get('location', '').replace('\\t', ' ').replace('\\n', ' ')[:100]
    desc = ev.get('description', '').replace('\\t', ' ').replace('\\n', ' ')[:500]
    
    # Parse dates
    start_str = ev.get('start', '')
    end_str = ev.get('end', start_str)
    all_day = '1' if ev.get('all_day', False) else '0'
    
    # Handle Z suffix
    if start_str.endswith('Z'):
        start_dt = datetime.strptime(start_str, '%Y-%m-%dT%H:%M:%SZ')
    else:
        start_dt = datetime.strptime(start_str[:19], '%Y-%m-%dT%H:%M:%S')
    
    if end_str.endswith('Z'):
        end_dt = datetime.strptime(end_str, '%Y-%m-%dT%H:%M:%SZ')
    else:
        end_dt = datetime.strptime(end_str[:19], '%Y-%m-%dT%H:%M:%S')
    
    # Output tab-separated
    parts = [
        summary, location, desc,
        str(start_dt.year), str(start_dt.month), str(start_dt.day),
        str(start_dt.hour), str(start_dt.minute),
        str(end_dt.year), str(end_dt.month), str(end_dt.day),
        str(end_dt.hour), str(end_dt.minute),
        all_day, uid
    ]
    print('\\t'.join(parts))
"
	
	-- Write Python script to file and execute
	do shell script "cat > " & pyFile & " << 'PYEOF'
" & pyScript & "
PYEOF"
	
	try
		set eventData to do shell script "/usr/bin/python3 " & pyFile & " " & quoted form of jsonPath
		do shell script "rm -f " & pyFile
	on error errMsg
		do shell script "rm -f " & pyFile
		return "ERROR parsing JSON: " & errMsg
	end try
	
	tell application "Microsoft Outlook"
		-- Find calendar
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
					if (count of eventParts) = 15 then
						-- Extract fields
						set evSummary to item 1 of eventParts
						set evLocation to item 2 of eventParts
						set evDescription to item 3 of eventParts
						set sYear to item 4 of eventParts
						set sMonth to item 5 of eventParts
						set sDay to item 6 of eventParts
						set sHour to item 7 of eventParts
						set sMin to item 8 of eventParts
						set eYear to item 9 of eventParts
						set eMonth to item 10 of eventParts
						set eDay to item 11 of eventParts
						set eHour to item 12 of eventParts
						set eMin to item 13 of eventParts
						set allDayFlag to item 14 of eventParts
						set evUID to item 15 of eventParts
						
						-- Build dates using string concatenation
						set startStr to sMonth & "/" & sDay & "/" & sYear & " " & sHour & ":" & sMin & ":00"
						set endStr to eMonth & "/" & eDay & "/" & eYear & " " & eHour & ":" & eMin & ":00"
						
						set startDate to date startStr
						set endDate to date endStr
						
						-- Build note content
						set noteText to evDescription & return & return
						set noteText to noteText & "[SRC: " & sourceId & "]" & return
						set noteText to noteText & "[ICSUID: " & evUID & "]"
						
						-- Create event
						set newEv to make new calendar event at targetCal
						tell newEv
							set subject to evSummary
							set location to evLocation
							set content to noteText
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
		return "✅ Created " & createdCount & " events (" & failedCount & " failed)"
	end tell
end run
