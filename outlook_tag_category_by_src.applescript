-- Usage:
--   osascript outlook_tag_category_by_src.applescript "SOURCE_ID" "CATEGORY_NAME" [DAYS_BACK]
-- Example:
--   osascript outlook_tag_category_by_src.applescript "lions-test" "Sports" 30

on run argv
  if (count of argv) < 2 then
    display dialog "Usage: osascript outlook_tag_category_by_src.applescript \"SOURCE_ID\" \"CATEGORY_NAME\" [DAYS_BACK]" buttons {"OK"} default button "OK"
    return
  end if

  set sourceId to item 1 of argv
  set catName to item 2 of argv
  set daysBack to 14
  if (count of argv) ≥ 3 then set daysBack to (item 3 of argv) as integer

  set nowDate to (current date)
  set sinceDate to nowDate - (daysBack * days)

  tell application "Microsoft Outlook"
    set matchedCount to 0
    set calEvents to calendar events whose start time ≥ sinceDate
    repeat with ev in calEvents
      try
        set n to ""
        try
          set n to the content of plain text content of ev
        end try
        -- Match by source token placed during import (e.g., "[ICSBridge:SOURCE_ID]" or your existing marker)
        if n contains sourceId then
          set category of ev to catName
          set matchedCount to matchedCount + 1
        end if
      end try
    end repeat
  end tell

  do shell script "echo Tagged " & matchedCount & " events with category '" & catName & "' for source '" & sourceId & "'."
end run
