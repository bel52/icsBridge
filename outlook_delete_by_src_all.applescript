on run argv
  if (count of argv) is less than 1 then error "Usage: osascript outlook_delete_by_src_all.applescript <SRC_ID>"
  set srcID to item 1 of argv
  set tagText to "[SRC: " & srcID & "]"

  set totalDeleted to 0
  set totalChecked to 0
  set report to ""

  tell application "Microsoft Outlook"
    -- Track per-name occurrence index so you know where items lived
    set seenNames to {}
    set occurCounts to {}

    repeat with c in calendars
      set calName to (name of c as text)

      -- calc occurrence for this calendar name
      set occ to 1
      set foundName to false
      repeat with i from 1 to (count of seenNames)
        if (item i of seenNames) is calName then
          set item i of occurCounts to ((item i of occurCounts) + 1)
          set occ to (item i of occurCounts)
          set foundName to true
          exit repeat
        end if
      end repeat
      if not foundName then
        set end of seenNames to calName
        set end of occurCounts to 1
        set occ to 1
      end if

      set deletedHere to 0
      set checkedHere to 0

      repeat with e in (calendar events of c)
        set checkedHere to checkedHere + 1
        set totalChecked to totalChecked + 1
        set d to ""
        try
          set d to content of e
        end try
        if (d contains tagText) then
          delete e
          set deletedHere to deletedHere + 1
          set totalDeleted to totalDeleted + 1
        end if
      end repeat

      if deletedHere > 0 then
        set report to report & calName & " (#" & occ & "): deleted " & deletedHere & " (checked " & checkedHere & ")" & linefeed
      end if
    end repeat
  end tell

  if report is "" then
    set report to "No events with " & tagText & " found."
  end if
  return "{\"ok\":true,\"deleted_total\":" & totalDeleted & ",\"checked_total\":" & totalChecked & ",\"note\":\"See lines below\"}" & linefeed & report
end run
