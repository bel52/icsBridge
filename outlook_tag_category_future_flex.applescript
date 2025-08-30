-- Usage:
--   osascript outlook_tag_category_future_flex.applescript "CALENDAR_NAME" "CATEGORY_NAME" "NEEDLE" [DAYS_AHEAD]
-- Example:
--   osascript outlook_tag_category_future_flex.applescript "Calendar" "Personal" "lions" 60

on toLower(s)
  try
    return (do shell script "python3 - <<PY\nprint(input().lower())\nPY" input s)
  on error
    return s
  end try
end toLower

on run argv
  if (count of argv) < 3 then
    display dialog "Usage: osascript outlook_tag_category_future_flex.applescript \"CALENDAR_NAME\" \"CATEGORY_NAME\" \"NEEDLE\" [DAYS_AHEAD]" buttons {"OK"} default button "OK"
    return
  end if

  set calName to item 1 of argv
  set catName to item 2 of argv
  set needleRaw to item 3 of argv
  set daysAhead to 365
  if (count of argv) ≥ 4 then set daysAhead to (item 4 of argv) as integer

  set needle to toLower(needleRaw)
  set nowDate to (current date)
  set untilDate to nowDate + (daysAhead * days)

  tell application "Microsoft Outlook"
    set matchedCount to 0

    set targetCals to calendars whose name is calName
    if (count of targetCals) is 0 then error "Calendar '" & calName & "' not found."
    set targetCal to item 1 of targetCals

    set calEvents to (calendar events of targetCal whose start time ≥ nowDate and start time ≤ untilDate)
    repeat with ev in calEvents
      try
        set subj to ""
        set orgz to ""
        set noteText to ""

        try set subj to subject of ev end try
        try set orgz to organizer of ev end try
        try set noteText to the content of plain text content of ev end try

        set sSubj to toLower(subj as string)
        set sOrgz to toLower(orgz as string)
        set sNote to toLower(noteText as string)

        if (sSubj contains needle) or (sOrgz contains needle) or (sNote contains needle) then
          -- Get current categories in both models
          set prevCat to ""
          try
            set prevCat to category of ev
            if prevCat is missing value then set prevCat to ""
          end try

          set prevCatsList to {}
          try
            set prevCatsList to categories of ev
            if prevCatsList is missing value then set prevCatsList to {}
          end try

          if catName is "" then
            -- Clear categories in both models
            try
              set category of ev to ""
            end try
            try
              set categories of ev to {}
            end try
            set matchedCount to matchedCount + 1
          else
            set alreadyHas to false
            if prevCat is not "" then
              if prevCat contains catName then set alreadyHas to true
            end if
            if (alreadyHas is false) and ((count of prevCatsList) > 0) then
              repeat with c in prevCatsList
                if (c as string) is equal to catName then
                  set alreadyHas to true
                end if
              end repeat
            end if

            if alreadyHas is false then
              -- Prefer multi-category API; fall back to single
              if (count of prevCatsList) > 0 then
                set end of prevCatsList to catName
                try
                  set categories of ev to prevCatsList
                on error
                  set category of ev to ((prevCat is not "" as boolean) as string)
                end try
                set matchedCount to matchedCount + 1
              else if prevCat is not "" then
                try
                  set category of ev to (prevCat & ", " & catName)
                on error
                  try
                    set categories of ev to {prevCat, catName}
                  end try
                end try
                set matchedCount to matchedCount + 1
              else
                try
                  set category of ev to catName
                on error
                  try
                    set categories of ev to {catName}
                  end try
                end try
                set matchedCount to matchedCount + 1
              end if
            end if
          end if
        end if
      end try
    end repeat
  end tell

  do shell script "echo Tagged " & matchedCount & " future events with category '" & catName & "' (needle: " & needleRaw & ", calendar: " & calName & ")."
end run
