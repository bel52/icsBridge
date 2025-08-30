-- Usage:
--   osascript outlook_tag_category_future_flex.applescript "CALENDAR_NAME" "CATEGORY_NAME" "NEEDLE" [DAYS_AHEAD]
-- Example:
--   osascript outlook_tag_category_future_flex.applescript "Calendar" "Personal" "lions" 60

on run argv
  if (count of argv) < 3 then
    display dialog "Usage: osascript outlook_tag_category_future_flex.applescript \"CALENDAR_NAME\" \"CATEGORY_NAME\" \"NEEDLE\" [DAYS_AHEAD]" buttons {"OK"} default button "OK"
    return
  end if

  set calName to item 1 of argv
  set catName to item 2 of argv
  set needle to item 3 of argv
  set daysAhead to 365
  if (count of argv) ≥ 4 then set daysAhead to (item 4 of argv) as integer

  set nowDate to (current date)
  set untilDate to nowDate + (daysAhead * days)

  set matchedCount to 0
  set scannedCount to 0

  tell application "Microsoft Outlook"
    set targetCals to calendars whose name is calName
    if (count of targetCals) is 0 then error "Calendar '" & calName & "' not found."
    set targetCal to item 1 of targetCals

    set calEvents to (calendar events of targetCal whose start time ≥ nowDate and start time ≤ untilDate)

    repeat with ev in calEvents
      set scannedCount to scannedCount + 1
      try
        set subj to ""
        set orgz to ""
        set noteText to ""
        try set subj to subject of ev end try
        try set orgz to organizer of ev end try
        try set noteText to the content of plain text content of ev end try

        set hit to false
        considering case
          -- do nothing, default is case-sensitive
        end considering
        ignoring case
          if (subj contains needle) or (orgz contains needle) or (noteText contains needle) then set hit to true
        end ignoring

        if hit then
          set curCat to ""
          try
            set curCat to category of ev
            if curCat is missing value then set curCat to ""
          end try

          if catName is "" then
            -- clear
            try
              set category of ev to ""
              set matchedCount to matchedCount + 1
            end try
          else
            if curCat is "" then
              try
                set category of ev to catName
                set matchedCount to matchedCount + 1
              end try
            else
              -- append if not already present (simple comma-space list)
              ignoring case
                if curCat does not contain catName then
                  try
                    set category of ev to (curCat & ", " & catName)
                    set matchedCount to matchedCount + 1
                  end try
                end if
              end ignoring
            end if
          end if
        end if
      end try
    end repeat
  end tell

  do shell script "echo Scanned " & scannedCount & " events; tagged " & matchedCount & " future events with category '" & catName & "' (needle: " & needle & ", calendar: " & calName & ")."
end run
