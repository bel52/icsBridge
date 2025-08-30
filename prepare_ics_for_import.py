#!/usr/bin/env python3
import sys
import urllib.request
import ssl
from datetime import datetime, date, timezone
from zoneinfo import ZoneInfo
from typing import Optional
from icalendar import Calendar

"""
prepare_ics_for_import.py

Goal:
- Fetch an ICS
- Tag each VEVENT's description with [SRC: <id>]
- Convert ANY timed events to UTC ("Z" form) so Outlook displays them in the
  user's local time (Eastern, in your case) without misapplying X-WR-TIMEZONE/TZID.
- Remove X-WR-TIMEZONE entirely to avoid client confusion.
- Leave all-day events (DATE-only) unchanged.
"""

def fetch_text(url: str) -> str:
    if url.startswith('webcal://'):
        url = 'https://' + url[9:]
    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
        return r.read().decode('utf-8', errors='ignore')

def get_calendar_default_tz(cal: Calendar) -> Optional[ZoneInfo]:
    """Return ZoneInfo from X-WR-TIMEZONE if present/valid; else None."""
    tzid = cal.get('X-WR-TIMEZONE')
    if tzid:
        try:
            return ZoneInfo(str(tzid))
        except Exception:
            return None
    return None

def to_utc(dt, default_tz: Optional[ZoneInfo], tz_param: Optional[str]):
    """
    Normalize a datetime to UTC:
    - If dt has tzinfo: convert directly to UTC
    - Else if TZID param present: localize then to UTC
    - Else if calendar default_tz present: localize then to UTC
    - Else: assume it's already "floating local" in ET; treat as America/New_York then to UTC
    """
    if not isinstance(dt, datetime):
        # DATE (all-day) – return unchanged
        return dt

    if dt.tzinfo is not None:
        return dt.astimezone(timezone.utc)

    # No tzinfo: try TZID param first
    if tz_param:
        try:
            return dt.replace(tzinfo=ZoneInfo(tz_param)).astimezone(timezone.utc)
        except Exception:
            pass

    # Then try calendar default
    if default_tz:
        try:
            return dt.replace(tzinfo=default_tz).astimezone(timezone.utc)
        except Exception:
            pass

    # Final fallback: assume Eastern wall-clock (your requirement)
    try:
        return dt.replace(tzinfo=ZoneInfo("America/New_York")).astimezone(timezone.utc)
    except Exception:
        # Last resort: force UTC without shift (not preferred, but safe)
        return dt.replace(tzinfo=timezone.utc)

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 prepare_ics_for_import.py <URL> <SOURCE_ID> <OUTPUT_PATH>", file=sys.stderr)
        sys.exit(1)

    source_url = sys.argv[1]
    source_id = sys.argv[2]
    output_file = sys.argv[3]

    print(f"Fetching and processing: {source_url}")
    try:
        ics_data = fetch_text(source_url)
        cal = Calendar.from_ical(ics_data)
        tag = f"\n\n[SRC: {source_id}]"

        default_tz = get_calendar_default_tz(cal)

        # Remove calendar-level timezone to prevent misinterpretation
        if 'X-WR-TIMEZONE' in cal:
            del cal['X-WR-TIMEZONE']

        # Walk events and normalize times
        for component in cal.walk('VEVENT'):
            # Tag description
            description = component.get('description', '')
            if tag not in description:
                component['description'] = description + tag

            for prop_name in ['dtstart', 'dtend']:
                if prop_name in component:
                    v = component.get(prop_name)
                    dt = v.dt

                    # Leave true all-day DATEs alone
                    if isinstance(dt, date) and not isinstance(dt, datetime):
                        continue

                    # Extract TZID param if present
                    tzid_param = None
                    try:
                        tzid_param = v.params.get('TZID')
                    except Exception:
                        tzid_param = None

                    # Convert to UTC
                    dt_utc = to_utc(dt, default_tz, tzid_param)

                    # Replace property: ensure UTC "Z" format
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
