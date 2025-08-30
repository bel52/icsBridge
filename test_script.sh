#!/bin/bash
# Test script for ICS Bridge

echo "🧪 Testing ICS Bridge Installation"
echo "═══════════════════════════════════════════════════════════"

# Check files exist
echo "📁 Checking files..."
for file in ~/icsBridge/ics_manager.sh ~/icsBridge/fetch_public_ics.py ~/icsBridge/outlook_create_events.js ~/icsBridge/outlook_remove_source.js; do
  if [[ -f "$file" ]]; then
    echo "  ✅ $(basename $file)"
  else
    echo "  ❌ Missing: $(basename $file)"
  fi
done

echo ""
echo "🔧 Testing Python..."
python3 -c "import json, datetime, urllib.request, ssl; print('  ✅ Python modules OK')"

echo ""
echo "📅 Testing Outlook connection..."
osascript -e 'tell application "Microsoft Outlook" to return "  ✅ Outlook version: " & version' 2>/dev/null || echo "  ❌ Cannot connect to Outlook"

echo ""
echo "📊 Testing calendar access..."
osascript -e 'tell application "Microsoft Outlook" to return "  ✅ Found " & (count of calendars) & " calendars"' 2>/dev/null || echo "  ❌ Cannot access calendars"

echo ""
echo "🌐 Testing webcal conversion..."
python3 -c "
url = 'webcal://example.com/calendar.ics'
if url.startswith('webcal://'):
    url = 'https://' + url[9:]
    print(f'  ✅ Webcal conversion: {url}')
"

echo ""
echo "═══════════════════════════════════════════════════════════"
echo "📝 Test Summary:"
echo ""
echo "To test the full workflow:"
echo "1. Run: cd ~/icsBridge && ./ics_manager.sh"
echo "2. Choose option 1 (Add calendar)"
echo "3. Try these test sources:"
echo ""
echo "Sports calendar:"
echo "  https://ics.calendarlabs.com/1982/c0cbc494/Detroit_Lions_Schedule.ics"
echo ""
echo "Webcal URL (MSU events):"
echo "  webcal://events.msu.edu/export.php?calendar=default&type=ical"
echo ""
echo "NASA launches:"
echo "  https://www.nasa.gov/templateimages/redesign/calendar/iCal/nasa_calendar.ics"
echo ""
echo "═══════════════════════════════════════════════════════════"
