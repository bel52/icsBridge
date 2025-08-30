#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# fetch_public_ics.py
# Usage:
#   source ~/calendarBridge/.venv/bin/activate
#   python3 fetch_public_ics.py "<ICS_URL>" "/tmp/public_events.json"
#
# Notes:
# - Minimal ICS parser for VEVENT blocks (SUMMARY, DTSTART, DTEND, UID, LOCATION, DESCRIPTION).
# - Handles: UTC (Z) times, local times (floating), and all-day VALUE=DATE.
# - No RRULE expansion (sports schedules typically list single games).
# - Output JSON structure consumed by outlook_write_events.js (JXA).
#
import sys, json, urllib.request, ssl, datetime, re

def fetch_text(url: str) -> str:
    # Allow older LibreSSL on macOS Python 3.9
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(url, context=ctx, timeout=30) as r:
            return r.read().decode('utf-8', errors='replace')
    except Exception as e:
        print(f"ERROR: failed to download {url}: {e}", file=sys.stderr)
        sys.exit(2)

def unfold_ics_lines(text: str):
    # RFC5545: lines can be folded (next line begins with space/tab)
    lines = text.splitlines()
    out = []
    for line in lines:
        if line.startswith((' ', '\t')) and out:
            out[-1] += line[1:]
        else:
            out.append(line)
    return out

def parse_dt(value: str, params: dict):
    # Supports:
    # - YYYYMMDD (all-day) when VALUE=DATE
    # - YYYYMMDDTHHMMSSZ (UTC)
    # - YYYYMMDDTHHMMSS (floating, treat as local)
    if params.get('VALUE','').upper() == 'DATE':
        # All day: interpret as midnight local
        dt = datetime.datetime.strptime(value, '%Y%m%d')
        return dt, True
    # date-time
    is_utc = value.endswith('Z')
    fmt = '%Y%m%dT%H%M%S' + ('Z' if is_utc else '')
    dt = datetime.datetime.strptime(value, fmt)
    if is_utc:
        # Keep as aware UTC then convert to ISO Z string later
        return dt.replace(tzinfo=datetime.timezone.utc), False
    # Floating (treat as local naive)
    return dt, False

def parse_property(line: str):
    # e.g. "DTSTART;TZID=America/New_York;VALUE=DATE:20250901"
    if ':' not in line:
        return None, None, {}, ''
    raw_name, raw_value = line.split(':', 1)
    parts = raw_name.split(';')
    name = parts[0].upper()
    params = {}
    for p in parts[1:]:
        if '=' in p:
            k,v = p.split('=',1)
            params[k.upper()] = v
    return name, raw_name, params, raw_value

def clean_text(v: str) -> str:
    # Unescape common ICS sequences
    v = v.replace('\\n', '\n').replace('\\,', ',').replace('\\;', ';')
    return v.strip()

def main():
    if len(sys.argv) != 3:
        print("Usage: python3 fetch_public_ics.py <ICS_URL> </path/to/output.json>", file=sys.stderr)
        sys.exit(1)

    url = sys.argv[1]
    out_json = sys.argv[2]

    raw = fetch_text(url)
    lines = unfold_ics_lines(raw)

    events = []
    cur = None

    for line in lines:
        if line.startswith('BEGIN:VEVENT'):
            cur = {'raw': {}}
        elif line.startswith('END:VEVENT'):
            if cur:
                # Normalize into output structure
                uid = cur.get('UID','').strip()
                summary = clean_text(cur.get('SUMMARY',''))
                description = clean_text(cur.get('DESCRIPTION',''))
                location = clean_text(cur.get('LOCATION',''))

                # Date handling
                all_day = False
                start_iso = None
                end_iso = None

                if 'DTSTART_parsed' in cur:
                    dt, is_all_day_s = cur['DTSTART_parsed']
                    all_day = all_day or is_all_day_s
                    if isinstance(dt, datetime.datetime) and dt.tzinfo is not None:
                        start_iso = dt.isoformat().replace('+00:00','Z')
                    else:
                        # local naive -> treat as local and emit ISO without TZ;
                        # JXA will interpret as local time.
                        start_iso = dt.strftime('%Y-%m-%dT%H:%M:%S')

                if 'DTEND_parsed' in cur:
                    dt, is_all_day_e = cur['DTEND_parsed']
                    all_day = all_day or is_all_day_e
                    if isinstance(dt, datetime.datetime) and dt.tzinfo is not None:
                        end_iso = dt.isoformat().replace('+00:00','Z')
                    else:
                        end_iso = dt.strftime('%Y-%m-%dT%H:%M:%S')
                else:
                    # Some ICS omit DTEND for all-day; treat as same day end
                    if 'DTSTART_parsed' in cur:
                        sdt, isad = cur['DTSTART_parsed']
                        if isad:
                            edt = sdt + datetime.timedelta(days=1)
                            end_iso = edt.strftime('%Y-%m-%dT%H:%M:%S')
                            all_day = True

                if uid and start_iso:
                    events.append({
                        "uid": uid,
                        "summary": summary or "(No title)",
                        "description": description,
                        "location": location,
                        "start": start_iso,  # ISO 8601 (Z for UTC or local without tz)
                        "end": end_iso,
                        "all_day": bool(all_day),
                    })
            cur = None
        elif cur is not None:
            name, raw_name, params, value = parse_property(line)
            if not name:
                continue
            if name in ('SUMMARY','UID','DESCRIPTION','LOCATION'):
                cur[name] = value
            elif name in ('DTSTART','DTEND'):
                try:
                    dt_parsed = parse_dt(value, params)
                    cur[name + '_parsed'] = dt_parsed
                except Exception:
                    pass
            # stash raw for debugging if needed
            cur['raw'][name] = value

    # Write JSON
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump({"events": events}, f, ensure_ascii=False, indent=2)
    print(f"Wrote {len(events)} events to {out_json}")

if __name__ == "__main__":
    main()
