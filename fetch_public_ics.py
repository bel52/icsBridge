#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# fetch_public_ics.py - Enhanced version with webcal://, JSON, and CSV support
# Usage:
#   python3 fetch_public_ics.py "<URL_OR_FILE>" "/tmp/output_events.json"
#
import sys, json, urllib.request, ssl, datetime, re, os

def fetch_text(url: str) -> str:
    """Fetch content from URL (supports http://, https://, webcal://)"""
    # Convert webcal:// to https://
    if url.startswith('webcal://'):
        url = 'https://' + url[9:]
        print(f"Converting webcal:// to https:// → {url}", file=sys.stderr)
    
    # Allow older LibreSSL on macOS
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
        })
        with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
            content = r.read()
            # Try to decode as UTF-8, fallback to Latin-1
            try:
                return content.decode('utf-8')
            except:
                return content.decode('latin-1', errors='replace')
    except Exception as e:
        print(f"ERROR: Failed to download {url}: {e}", file=sys.stderr)
        sys.exit(2)

def read_file(filepath: str) -> str:
    """Read local file"""
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        return f.read()

def unfold_ics_lines(text: str):
    """RFC5545: unfold lines that begin with space/tab"""
    lines = text.splitlines()
    out = []
    for line in lines:
        if line.startswith((' ', '\t')) and out:
            out[-1] += line[1:]
        else:
            out.append(line)
    return out

def parse_dt(value: str, params: dict):
    """Parse ICS datetime values"""
    # Handle VALUE=DATE (all-day events)
    if params.get('VALUE','').upper() == 'DATE':
        if len(value) == 8:  # YYYYMMDD
            dt = datetime.datetime.strptime(value, '%Y%m%d')
            return dt, True
    
    # Handle datetime with timezone
    is_utc = value.endswith('Z')
    
    # Clean up value - remove timezone info for parsing
    clean_val = value.replace('Z', '').replace('T', '')
    
    try:
        if len(clean_val) == 8:  # YYYYMMDD
            dt = datetime.datetime.strptime(clean_val, '%Y%m%d')
            return dt, True
        elif len(clean_val) == 14:  # YYYYMMDDHHMMSS
            dt = datetime.datetime.strptime(clean_val, '%Y%m%d%H%M%S')
        elif len(clean_val) == 15 and 'T' in value:  # YYYYMMDDTHHMMSS
            dt = datetime.datetime.strptime(value.replace('Z',''), '%Y%m%dT%H%M%S')
        else:
            # Fallback
            dt = datetime.datetime.strptime(value[:8], '%Y%m%d')
            return dt, True
    except:
        # If all else fails, use today
        dt = datetime.datetime.now()
        return dt, False
    
    if is_utc:
        dt = dt.replace(tzinfo=datetime.timezone.utc)
    
    return dt, False

def parse_property(line: str):
    """Parse ICS property line"""
    if ':' not in line:
        return None, None, {}, ''
    raw_name, raw_value = line.split(':', 1)
    parts = raw_name.split(';')
    name = parts[0].upper()
    params = {}
    for p in parts[1:]:
        if '=' in p:
            k,v = p.split('=',1)
            params[k.upper()] = v.strip('"')
    return name, raw_name, params, raw_value

def clean_text(v: str) -> str:
    """Unescape ICS text"""
    v = v.replace('\\n', '\n').replace('\\N', '\n')
    v = v.replace('\\,', ',').replace('\\;', ';')
    v = v.replace('\\\\', '\\')
    return v.strip()

def parse_ics(text: str) -> list:
    """Parse ICS/iCal content into events"""
    lines = unfold_ics_lines(text)
    events = []
    cur = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.upper() == 'BEGIN:VEVENT':
            cur = {'raw': {}}
        elif line.upper() == 'END:VEVENT':
            if cur:
                # Process event
                uid = cur.get('UID','').strip()
                if not uid:
                    uid = f"generated_{len(events)}_{datetime.datetime.now().timestamp()}"
                
                summary = clean_text(cur.get('SUMMARY',''))
                description = clean_text(cur.get('DESCRIPTION',''))
                location = clean_text(cur.get('LOCATION',''))
                
                # Date handling
                all_day = False
                start_iso = None
                end_iso = None
                
                if 'DTSTART_parsed' in cur:
                    dt, is_all_day = cur['DTSTART_parsed']
                    all_day = is_all_day
                    if hasattr(dt, 'tzinfo') and dt.tzinfo:
                        start_iso = dt.isoformat().replace('+00:00','Z')
                    else:
                        start_iso = dt.strftime('%Y-%m-%dT%H:%M:%S')
                
                if 'DTEND_parsed' in cur:
                    dt, is_all_day = cur['DTEND_parsed']
                    all_day = all_day or is_all_day
                    if hasattr(dt, 'tzinfo') and dt.tzinfo:
                        end_iso = dt.isoformat().replace('+00:00','Z')
                    else:
                        end_iso = dt.strftime('%Y-%m-%dT%H:%M:%S')
                elif start_iso:
                    # No end time - use start + 1 hour or 1 day
                    if all_day:
                        end_dt = cur['DTSTART_parsed'][0] + datetime.timedelta(days=1)
                    else:
                        end_dt = cur['DTSTART_parsed'][0] + datetime.timedelta(hours=1)
                    end_iso = end_dt.strftime('%Y-%m-%dT%H:%M:%S')
                
                if start_iso:
                    events.append({
                        "uid": uid,
                        "summary": summary or "(No title)",
                        "description": description,
                        "location": location,
                        "start": start_iso,
                        "end": end_iso or start_iso,
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
                except Exception as e:
                    print(f"Warning: Could not parse {name}: {value} - {e}", file=sys.stderr)
            
            cur['raw'][name] = value
    
    return events

def parse_json_feed(text: str) -> list:
    """Parse JSON feed format (common calendar export format)"""
    try:
        data = json.loads(text)
        events = []
        
        # Handle different JSON structures
        if isinstance(data, list):
            items = data
        elif 'events' in data:
            items = data['events']
        elif 'items' in data:
            items = data['items']
        else:
            items = [data]
        
        for item in items:
            # Try to extract event info from various formats
            uid = (item.get('id') or item.get('uid') or 
                   item.get('eventId') or f"json_{len(events)}")
            
            summary = (item.get('summary') or item.get('title') or 
                      item.get('subject') or item.get('name') or "(No title)")
            
            description = (item.get('description') or item.get('body') or 
                          item.get('details') or "")
            
            location = (item.get('location') or item.get('venue') or 
                       item.get('place') or "")
            
            # Parse dates
            start = (item.get('start') or item.get('startTime') or 
                    item.get('start_time') or item.get('begins'))
            end = (item.get('end') or item.get('endTime') or 
                  item.get('end_time') or item.get('ends'))
            
            if isinstance(start, dict):
                start = start.get('dateTime') or start.get('date')
            if isinstance(end, dict):
                end = end.get('dateTime') or end.get('date')
            
            # Convert to ISO format if needed
            if start:
                if not isinstance(start, str):
                    start = str(start)
                # Ensure ISO format
                if 'T' not in start and len(start) == 10:
                    start = start + 'T00:00:00'
                
                all_day = item.get('allDay', False) or (len(start) == 10)
                
                if not end:
                    end = start
                elif not isinstance(end, str):
                    end = str(end)
                
                if 'T' not in end and len(end) == 10:
                    end = end + 'T23:59:59'
                
                events.append({
                    "uid": str(uid),
                    "summary": summary,
                    "description": description,
                    "location": location,
                    "start": start,
                    "end": end,
                    "all_day": all_day
                })
        
        return events
    except:
        return []

def detect_format(content: str) -> str:
    """Detect content format"""
    content = content.strip()
    if content.startswith('BEGIN:VCALENDAR'):
        return 'ics'
    elif content.startswith('{') or content.startswith('['):
        return 'json'
    elif 'BEGIN:VCALENDAR' in content[:1000]:
        return 'ics'
    else:
        return 'unknown'

def main():
    if len(sys.argv) != 3:
        print("Usage: python3 fetch_public_ics.py <URL_OR_FILE> </path/to/output.json>", file=sys.stderr)
        sys.exit(1)
    
    source = sys.argv[1]
    out_json = sys.argv[2]
    
    # Fetch or read content
    if source.startswith(('http://', 'https://', 'webcal://')):
        print(f"Fetching from URL: {source}", file=sys.stderr)
        content = fetch_text(source)
    elif os.path.isfile(source):
        print(f"Reading local file: {source}", file=sys.stderr)
        content = read_file(source)
    else:
        print(f"ERROR: Invalid source (not a URL or file): {source}", file=sys.stderr)
        sys.exit(1)
    
    # Detect and parse format
    format_type = detect_format(content)
    print(f"Detected format: {format_type}", file=sys.stderr)
    
    if format_type == 'ics':
        events = parse_ics(content)
    elif format_type == 'json':
        events = parse_json_feed(content)
    else:
        # Try ICS as fallback
        print("Unknown format, trying ICS parser...", file=sys.stderr)
        events = parse_ics(content)
    
    # Write output
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump({"events": events}, f, ensure_ascii=False, indent=2)
    
    print(f"Wrote {len(events)} events to {out_json}")

if __name__ == "__main__":
    main()
