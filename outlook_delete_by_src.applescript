on run argv
  if (count of argv) is less than 3 then
    error "Usage: outlook_delete_by_src.applescript <SRC_ID> <CAL_NAME> <INDEX>"
  end if
  set srcID to item 1 of argv
  set calName to item 2 of argv
  set idxText to item 3 of argv
  try
    set calIndex to (idxText as integer)
  on error
    error "INDEX must be a number"
  end try

  set tagText to "[SRC: " & srcID & "]"
  set deletedCount to 0
  set checkedCount to 0

  tell application "Microsoft Outlook"
    set targetCal to missing value
    set seen to 0
    repeat with c in calendars
      if (name of c as text) is calName then
        set seen to seen + 1
        if seen = calIndex then
          set targetCal to c
          exit repeat
        end if
      end if
    end repeat
    if targetCal is missing value then
      error "Calendar '" & calName & "' (#" & calIndex & ") not found."
    end if

    set evs to calendar events of targetCal
    repeat with e in evs
      set checkedCount to checkedCount + 1
      try
        set d to content of e
      on error
        set d to ""
      end try
      if (d contains tagText) then
        delete e
        set deletedCount to deletedCount + 1
      end if
    end repeat
  end tell

  return "{\"ok\":true,\"deleted\":" & deletedCount & ",\"checked\":" & checkedCount & "}"
end run
