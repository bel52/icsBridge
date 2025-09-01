#!/usr/bin/env python3
import sys
import os
import ssl
import urllib.request
from datetime import datetime, date, timezone, timedelta
from zoneinfo import ZoneInfo
from typing import Optional
from icalendar import Calendar

def fetch_text(url: str) -> str:
    """
    Download an iCalendar file from a URL (supports webcal://).
    Returns the raw text.
    """
    # Convert webcal to https
    if url.startswith('webcal://'):
        url = 'https://' + url[len('webcal://'):]
    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as response:
        return response.read().decode('utf-8', errors='ignore')

def get_calendar_default_tz(cal: Calendar) -> Optional[ZoneInfo]:
    """
    Return a ZoneInfo object for the calendar's X-WR-TIMEZONE,
    or None if unavailable or invalid.
    """
    tzid = cal.get('X-WR-TIMEZONE')
    if tzid:
        try:
            return ZoneInfo(str(tzid))
        except Exception:
            return None
    return None

def to_utc(dt, default_tz: Optional[ZoneInfo], tz_param: Optional[str]):
    """
    Convert a datetime (with or without timezone info) into UTC.
    - If the datetime already has tzinfo, convert to UTC.
    - If tz_param (TZID) is provided, use that as the timezone.
    - Otherwise fall back to the calendar default tz, then to America/New_York.
    """
    if not isinstance(dt, datetime):
        return dt
    if dt.tzinfo:
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
    # Final fallback: assume America/New_York
    try:
        return dt.replace(tzinfo=ZoneInfo('America/New_York')).astimezone(timezone.utc)
    except Exception:
        return dt.replace(tzinfo=timezone.utc)

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 prepare_ics_for_import.py <URL_or_file> <SOURCE_ID> <OUTPUT_PATH>", file=sys.stderr)
        sys.exit(1)

    source, source_id, output_file = sys.argv[1], sys.argv[2], sys.argv[3]

    print(f"Fetching and processing: {source}")
    try:
        ics_data = fetch_text(source)
        cal = Calendar.from_ical(ics_data)
        tag = f"\n\n[SRC: {source_id}]"

        default_tz = get_calendar_default_tz(cal)
        # Remove X-WR-TIMEZONE to prevent Outlook second-guessing
        if 'X-WR-TIMEZONE' in cal:
            del cal['X-WR-TIMEZONE']

        for comp in cal.walk('VEVENT'):
            # Append [SRC: id] tag if not present
            desc = comp.get('description', '')
            if tag not in desc:
                comp['description'] = desc + tag

            dtstart_utc = None

            # Normalize DTSTART
            if 'dtstart' in comp:
                vstart = comp.get('dtstart')
                dtstart_val = vstart.dt
                # Leave all-day events unchanged (date-only)
                if not (isinstance(dtstart_val, date) and not isinstance(dtstart_val, datetime)):
                    tzid_param = None
                    try:
                        tzid_param = vstart.params.get('TZID')
                    except Exception:
                        pass
                    dtstart_utc = to_utc(dtstart_val, default_tz, tzid_param)
                    comp.pop('dtstart')
                    comp.add('dtstart', dtstart_utc)
                else:
                    dtstart_utc = dtstart_val  # keep date-only

            # Normalize DTEND or compute from duration
            if 'dtend' in comp:
                vend = comp.get('dtend')
                dtend_val = vend.dt
                # If dtend is a date-only (all-day), leave unchanged
                if not (isinstance(dtend_val, date) and not isinstance(dtend_val, datetime)):
                    tzid_param = None
                    try:
                        tzid_param = vend.params.get('TZID')
                    except Exception:
                        pass
                    dtend_utc = to_utc(dtend_val, default_tz, tzid_param)
                    comp.pop('dtend')
                    comp.add('dtend', dtend_utc)
            elif 'duration' in comp and isinstance(dtstart_utc, datetime):
                # If there is no DTEND but a DURATION is specified,
                # compute DTEND as DTSTART + DURATION
                try:
                    duration = comp.decoded('duration')
                    if isinstance(duration, timedelta):
                        dtend_calc = dtstart_utc + duration
                        comp.add('dtend', dtend_calc)
                        # Remove DURATION since we now have an explicit DTEND
                        comp.pop('duration')
                except Exception:
                    pass

        # Write normalized calendar to output file
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())

        print(f"Successfully created UTC-normalized ICS at: {output_file}")
    except Exception as e:
        print(f"ERROR: Failed to process calendar. {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()

