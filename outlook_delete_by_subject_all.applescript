on run argv
  if (count of argv) is less than 1 then error "Usage: osascript outlook_delete_by_subject_all.applescript <SUBSTRING>"
  set needle to item 1 of argv
  set needleLower to my toLower(needle)
  set totalDeleted to 0
  tell application "Microsoft Outlook"
    repeat with c in calendars
      try
        set evs to (calendar events of c) as list
        repeat with e in evs
          set subj to ""
          set dsc to ""
          try set subj to subject of e as text end try
          try set dsc to content of e as text end try
          if (my containsCI(subj, needleLower)) or (my containsCI(dsc, needleLower)) then delete e
        end repeat
      end try
    end repeat
  end tell
  return "{\"ok\":true}"
end run
on toLower(t)
  try
    do shell script "/usr/bin/python3 - <<'PY'\nimport sys;print(sys.stdin.read().lower())\nPY" with input t
  on error
    return t
  end try
end toLower
on containsCI(hay, needleLower)
  try
    set hayLower to my toLower(hay)
    set AppleScript's text item delimiters to needleLower
    set parts to text items of hayLower
    set AppleScript's text item delimiters to ""
    if (count of parts) > 1 then return true
  end try
  return false
end containsCI
