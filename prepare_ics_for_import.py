#!/usr/bin/env python3
import sys, os
import urllib.request, ssl
from urllib.parse import urlparse, unquote
from datetime import datetime, date, timezone
from zoneinfo import ZoneInfo
from typing import Optional
from icalendar import Calendar

def fetch_text(url_or_path: str) -> str:
  # webcal -> https
  if url_or_path.startswith('webcal://'):
    url_or_path = 'https://' + url_or_path[9:]
  parsed = urlparse(url_or_path)

  # Plain path
  if parsed.scheme == '' and os.path.exists(url_or_path):
    with open(url_or_path, 'rb') as f:
      return f.read().decode('utf-8', errors='ignore')

  # file://
  if parsed.scheme == 'file':
    path = unquote(parsed.path)
    with open(path, 'rb') as f:
      return f.read().decode('utf-8', errors='ignore')

  # http/https
  ctx = ssl._create_unverified_context()
  req = urllib.request.Request(url_or_path, headers={'User-Agent': 'Mozilla/5.0'})
  with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
    return r.read().decode('utf-8', errors='ignore')

def get_calendar_default_tz(cal: Calendar) -> Optional[ZoneInfo]:
  tzid = cal.get('X-WR-TIMEZONE')
  if tzid:
    try: return ZoneInfo(str(tzid))
    except Exception: return None
  return None

def to_utc(dt, default_tz: Optional[ZoneInfo], tz_param: Optional[str]):
  if not isinstance(dt, datetime): return dt
  if dt.tzinfo is not None: return dt.astimezone(timezone.utc)
  if tz_param:
    try: return dt.replace(tzinfo=ZoneInfo(tz_param)).astimezone(timezone.utc)
    except Exception: pass
  if default_tz:
    try: return dt.replace(tzinfo=default_tz).astimezone(timezone.utc)
    except Exception: pass
  try: return dt.replace(tzinfo=ZoneInfo("America/New_York")).astimezone(timezone.utc)
  except Exception: return dt.replace(tzinfo=timezone.utc)

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
    if 'X-WR-TIMEZONE' in cal:
      del cal['X-WR-TIMEZONE']

    for comp in cal.walk('VEVENT'):
      description = comp.get('description', '')
      if tag not in description:
        comp['description'] = description + tag
      for prop in ['dtstart', 'dtend']:
        if prop in comp:
          v = comp.get(prop)
          dt = v.dt
          if isinstance(dt, date) and not isinstance(dt, datetime):
            continue  # all-day, leave as is
          tzid = None
          try: tzid = v.params.get('TZID')
          except Exception: tzid = None
          dt_utc = to_utc(dt, default_tz, tzid)
          comp.pop(prop)
          comp.add(prop, dt_utc)

    with open(output_file, 'wb') as f:
      f.write(cal.to_ical())
    print(f"Successfully created UTC-normalized ICS at: {output_file}")
  except Exception as e:
    print(f"ERROR: Failed to process calendar. {e}", file=sys.stderr)
    sys.exit(1)

if __name__ == "__main__":
  main()
