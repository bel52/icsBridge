#!/usr/bin/env python3
import sys
import urllib.request
import ssl
from icalendar import Calendar

def fetch_text(url: str) -> str:
    if url.startswith('webcal://'):
        url = 'https://' + url[9:]
    ctx = ssl._create_unverified_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
        return r.read().decode('utf-8', errors='ignore')

def main():
    if len(sys.argv) != 3:
        print("Usage: python3 fetch_and_format_ics.py <URL> <OUTPUT_PATH>", file=sys.stderr)
        sys.exit(1)
    
    source_url = sys.argv[1]
    output_file = sys.argv[2]
    
    print(f"Fetching from: {source_url}")
    try:
        ics_data = fetch_text(source_url)
        # The icalendar library will parse and clean up the data
        cal = Calendar.from_ical(ics_data)
        
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())
        
        print(f"Successfully created clean ICS file at: {output_file}")
    except Exception as e:
        print(f"ERROR: Failed to process calendar. {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
