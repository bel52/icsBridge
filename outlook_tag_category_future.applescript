-- Usage:
--   osascript outlook_tag_category_future.applescript "SOURCE_ID" "CATEGORY_NAME" [DAYS_AHEAD]
-- Example:
--   osascript outlook_tag_category_future.applescript "lions" "Personal" 365

on run argv
  if (count of argv) < 2 then
    display dialog "Usage: osascript outlook_tag_category_future.applescript \"SOURCE_ID\" \"CATEGORY_NAME\" [DAYS_AHEAD]" buttons {"OK"} default button "OK"
    return
  end if

  set sourceId to item 1 of argv
  set catName to item 2 of argv
  set daysAhead to 365
  if (count of argv) ≥ 3 then set daysAhead to (item 3 of argv) as integer

  set nowDate to (current date)
  set untilDate to nowDate + (daysAhead * days)

  tell application "Microsoft Outlook"
    set matchedCount to 0
    set calEvents to calendar events whose start time ≥ nowDate and start time ≤ untilDate
    repeat with ev in calEvents
      try
        set n to ""
        set s to ""
        try
          set n to the content of plain text content of ev
        end try
        try
          set s to subject of ev
        end try

        if (n contains sourceId) or (s contains sourceId) then
          set prevCat to ""
          try
            set prevCat to category of ev
          end try
          if prevCat is missing value then set prevCat to ""

          if catName is "" then
            -- Clear all categories when empty string provided
            set category of ev to ""
            set matchedCount to matchedCount + 1
          else
            if prevCat is "" then
              set category of ev to catName
              set matchedCount to matchedCount + 1
            else if prevCat does not contain catName then
              set category of ev to (prevCat & ", " & catName)
              set matchedCount to matchedCount + 1
            end if
          end if
        end if
      end try
    end repeat
  end tell

  do shell script "echo Tagged " & matchedCount & " future events with category '" & catName & "' for source '" & sourceId & "'."
end run
