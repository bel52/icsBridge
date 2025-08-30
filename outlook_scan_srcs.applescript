on run argv
  if (count of argv) is less than 2 then
    error "Usage: outlook_scan_srcs.applescript <CAL_NAME> <INDEX>"
  end if
  set calName to item 1 of argv
  set calIndex to (item 2 of argv) as integer

  set ids to {}
  set counts to {}

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

    repeat with e in (calendar events of targetCal)
      set d to ""
      try
        set d to content of e
      end try
      set srcID to my extractSRC(d)
      if srcID is not "" then my addCount(srcID, ids, counts)
    end repeat
  end tell

  -- Build JSON
  set json to "{\"sources\":["
  set n to (count of ids)
  repeat with i from 1 to n
    set json to json & "{\"id\":\"" & (item i of ids) & "\",\"count\":" & (item i of counts) & "}"
    if i < n then set json to json & ","
  end repeat
  set json to json & "]}"
  return json
end run

on addCount(srcID, ids, counts)
  set found to false
  repeat with i from 1 to (count of ids)
    if (item i of ids) is srcID then
      set item i of counts to ((item i of counts) + 1)
      set found to true
      exit repeat
    end if
  end repeat
  if not found then
    set end of ids to srcID
    set end of counts to 1
  end if
end addCount

on extractSRC(t)
  if t is missing value then return ""
  set s to (t as text)
  set tag to "[SRC: "
  set p to offset of tag in s
  if p = 0 then return ""
  set afterTag to text (p + (length of tag)) thru -1 of s
  set q to offset of "]" in afterTag
  if q = 0 then return ""
  set rawID to text 1 thru (q - 1) of afterTag
  return my trimBoth(rawID)
end extractSRC

on trimBoth(s)
  if s is "" then return s
  set ws to {" ", tab, return, linefeed}
  set out to s as text
  repeat while my beginsWithAny(out, ws)
    set out to text 2 thru -1 of out
    if out is "" then return out
  end repeat
  repeat while my endsWithAny(out, ws)
    set L to (length of out)
    if L is 0 then exit repeat
    set out to text 1 thru (L - 1) of out
  end repeat
  return out
end trimBoth

on beginsWithAny(t, lst)
  repeat with ch in lst
    if (t begins with (ch as text)) then return true
  end repeat
  return false
end beginsWithAny

on endsWithAny(t, lst)
  repeat with ch in lst
    if (t ends with (ch as text)) then return true
  end repeat
  return false
end endsWithAny
