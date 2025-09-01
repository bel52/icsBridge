on run
  set recs to {} -- each item: {id:"...", cal:"...", idx:n, count:n}

  tell application "Microsoft Outlook"
    set nameSeen to {} -- names encountered to track per-name index
    set idxSeen to {}  -- corresponding counts

    set i to 0
    repeat with c in calendars
      set calName to (name of c as text)

      -- compute occurrence index for this name
      set hit to false
      set occ to 0
      repeat with k from 1 to (count of nameSeen)
        if (item k of nameSeen) is calName then
          set item k of idxSeen to ((item k of idxSeen) + 1)
          set occ to (item k of idxSeen)
          set hit to true
          exit repeat
        end if
      end repeat
      if not hit then
        set end of nameSeen to calName
        set end of idxSeen to 1
        set occ to 1
      end if

      repeat with e in (calendar events of c)
        set d to ""
        try set d to content of e end try
        set srcID to my extractSRC(d)
        if srcID is not "" then my addOrBump(srcID, calName, occ, recs)
      end repeat
    end repeat
  end tell

  -- Build JSON
  set json to "{\"sources\":["
  set n to (count of recs)
  repeat with r from 1 to n
    set it to item r of recs
    set json to json & "{\"id\":\"" & (id of it) & "\",\"calendar\":\"" & (cal of it) & "\",\"index\":" & (idx of it) & ",\"count\":" & (count of it) & "}"
    if r < n then set json to json & ","
  end repeat
  set json to json & "]}"
  return json
end run

on addOrBump(srcID, calName, occ, recs)
  set found to false
  repeat with i from 1 to (count of recs)
    set it to item i of recs
    if ((id of it) is srcID) and ((cal of it) is calName) and ((idx of it) is occ) then
      set count of it to ((count of it) + 1)
      set item i of recs to it
      set found to true
      exit repeat
    end if
  end repeat
  if not found then
    set end of recs to {id:srcID, cal:calName, idx:occ, count:1}
  end if
end addOrBump

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
