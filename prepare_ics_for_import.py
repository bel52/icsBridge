#!/usr/bin/env python3
import sys
import urllib.request
import ssl
from datetime import datetime, date, timezone
from zoneinfo import ZoneInfo
from typing import Optional
from icalendar import Calendar

def fetch_text(url: str) -> str:
    if url.startswith('webcal://'):
        url = 'https://' + url[9:]
    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
        return r.read().decode('utf-8', errors='ignore')

def get_calendar_default_tz(cal: Calendar) -> Optional[ZoneInfo]:
    tzid = cal.get('X-WR-TIMEZONE')
    if tzid:
        try:
            return ZoneInfo(str(tzid))
        except Exception:
            return None
    return None

def to_utc(dt, default_tz: Optional[ZoneInfo], tz_param: Optional[str]):
    if not isinstance(dt, datetime):
        return dt
    if dt.tzinfo is not None:
        return dt.astimezone(timezone.utc)
    if tz_param:
        try:
            return dt.replace(tzinfo=ZoneInfo(tz_param)).astimezone(timezone.utc)
        except Exception:
            pass
    if default_tz:
        try:
            return dt.replace(tzinfo=default_tz).astimezone(timezone.utc)
        except Exception:
            pass
    try:
        return dt.replace(tzinfo=ZoneInfo("America/New_York")).astimezone(timezone.utc)
    except Exception:
        return dt.replace(tzinfo=timezone.utc)

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 prepare_ics_for_import.py <URL> <SOURCE_ID> <OUTPUT_PATH>", file=sys.stderr)
        sys.exit(1)

    source_url, source_id, output_file = sys.argv[1], sys.argv[2], sys.argv[3]

    print(f"Fetching and processing: {source_url}")
    try:
        ics_data = fetch_text(source_url)
        cal = Calendar.from_ical(ics_data)
        tag = f"\n\n[SRC: {source_id}]"

        default_tz = get_calendar_default_tz(cal)
        if 'X-WR-TIMEZONE' in cal:
            del cal['X-WR-TIMEZONE']

        for component in cal.walk('VEVENT'):
            description = component.get('description', '')
            if tag not in description:
                component['description'] = description + tag

            for prop_name in ['dtstart', 'dtend']:
                if prop_name in component:
                    v = component.get(prop_name)
                    dt = v.dt
                    if isinstance(dt, date) and not isinstance(dt, datetime):
                        continue
                    tzid_param = None
                    try:
                        tzid_param = v.params.get('TZID')
                    except Exception:
                        tzid_param = None
                    dt_utc = to_utc(dt, default_tz, tzid_param)
                    component.pop(prop_name)
                    component.add(prop_name, dt_utc)

        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())

        print(f"Successfully created UTC-normalized ICS at: {output_file}")
    except Exception as e:
        print(f"ERROR: Failed to process calendar. {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
