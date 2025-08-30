#!/usr/bin/env python3
import sys
import urllib.request
import ssl
import pytz
from datetime import datetime
from icalendar import Calendar, Event

def fetch_text(url: str) -> str:
    """Fetches calendar data from a URL."""
    if url.startswith('webcal://'):
        url = 'https://' + url[9:]
    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
        return r.read().decode('utf-8', errors='ignore')

# A standard, known-good VTIMEZONE definition for America/New_York.
VTIMEZONE_DEFINITION = """BEGIN:VTIMEZONE
TZID:America/New_York
LAST-MODIFIED:20221013T163533Z
BEGIN:DAYLIGHT
TZNAME:EDT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZNAME:EST
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE"""

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 prepare_ics_for_import.py <URL> <SOURCE_ID> <OUTPUT_PATH>", file=sys.stderr)
        sys.exit(1)
    
    source_url = sys.argv[1]
    source_id = sys.argv[2]
    output_file = sys.argv[3]
    local_tz = pytz.timezone('America/New_York')
    
    print(f"Fetching and processing: {source_url}")
    try:
        ics_data = fetch_text(source_url)
        cal = Calendar.from_ical(ics_data)
        tag = f"\n\n[SRC: {source_id}]"

        tz_component = Calendar.from_ical(VTIMEZONE_DEFINITION).walk('VTIMEZONE')[0]
        cal.add_component(tz_component)

        for component in cal.walk():
            if component.name == "VEVENT":
                description = component.get('description', '')
                if tag not in description:
                    component['description'] = description + tag

                for prop in ['dtstart', 'dtend']:
                    if prop in component:
                        dt = component.get(prop).dt
                        if isinstance(dt, datetime):
                            # This is the new, robust logic
                            # 1. Convert the original time to the target timezone (America/New_York)
                            if dt.tzinfo is not None:
                                local_dt = dt.astimezone(local_tz)
                            else:
                                local_dt = local_tz.localize(dt)
                            
                            # 2. Make the datetime naive again, but now it represents the correct "wall clock" time
                            naive_dt = local_dt.replace(tzinfo=None)
                            
                            # 3. Replace the old property with the new naive time, but add the TZID parameter
                            # This is the most unambiguous way to tell Outlook the correct time.
                            component.pop(prop)
                            component.add(prop, naive_dt, parameters={'TZID': local_tz.zone})
        
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())
        
        print(f"Successfully created timezone-aware ICS file at: {output_file}")

    except Exception as e:
        print(f"ERROR: Failed to process calendar. {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
