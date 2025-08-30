# icsBridge — Import public iCalendar (ICS) into Outlook (macOS)

Local-only tool to import public ICS/ICAL schedules (e.g., sports) into the **Outlook desktop app** on macOS.
No Exchange/Graph calls — Outlook syncs them upstream like any event you created.

## What it does

* **Add**: Paste an `.ics`/`.ical` URL or file; events are created in your chosen Outlook calendar.
* **List**: See tracked sources you’ve imported.
* **Remove**: Delete all events imported for a given source (based on `[SRC: <id>]` tag in the notes).

Events are tagged in the notes with:

```
[SRC: <your-source-id>]
[ICSUID: <upstream-uid>]
```

## Repo layout

```
ics_manager.sh                     # interactive CLI (Add/Remove/List) + URL/file validation
fetch_public_ics.py                # tiny ICS → JSON parser (no external deps)
outlook_create_events.applescript  # creates events (AppleScript, most-compatible approach)
outlook_remove_source.js           # removes events tagged with [SRC: ...]
logs/                              # import logs (gitignored)
sources.json                       # tracks sources (gitignored)
```

## Requirements

* macOS with **Legacy Outlook** (Help → “Revert to Legacy Outlook”). New Outlook’s scripting is limited.
* Allow your terminal app to control Outlook:

  * **System Settings → Privacy & Security → Automation** → (Terminal/iTerm/VS Code) → **Microsoft Outlook**
  * Sometimes also **Privacy & Security → Accessibility** → enable your terminal app.

## Quick start

```bash
cd ~/icsBridge
./ics_manager.sh
# 1) Add calendar
#   - Paste an ICS URL (e.g. https://ics.calendarlabs.com/.../Detroit_Lions_Schedule.ics) or file path
#   - Pick a short source ID (e.g. Lions_2025)
#   - Calendar name: e.g. "Calendar" (or try a writable one like "Sports Imports")
#   - Occurrence index: e.g. 2 for the second "Calendar"
```

## Troubleshooting

### “A privilege violation occurred. (-10004)”

macOS privacy is blocking writes to Outlook. Fix:

1. Use **Legacy Outlook**.
2. System Settings → Privacy & Security → **Automation** → enable your terminal app for **Microsoft Outlook**.
3. (Sometimes) also enable your terminal app in **Accessibility**.
4. Quit/relaunch Outlook and your terminal app.

Sanity test (creates a 1-hour “Test” event in your default calendar):

```applescript
tell application "Microsoft Outlook"
  make new calendar event with properties {subject:"Test via Terminal", start time:(current date), end time:(current date) + 3600}
end tell
```

### “A property can’t go after this identifier. (-2740)” / “Expected ':' but found property. (-2741)”

These are strict AppleScript parser quirks. The included `outlook_create_events.applescript` avoids line continuations and sets properties in simple, compatible statements.

### Target calendar might be read-only

Some calendars (shared/subscribed) don’t allow scripted writes. If writes fail only for a specific calendar:

* Try creating a local calendar (once):

  ```applescript
  tell application "Microsoft Outlook"
    if (count of (calendars whose name is "Sports Imports")) = 0 then
      make new calendar with properties {name:"Sports Imports"}
    end if
  end tell
  ```
* Then import into **"Sports Imports"** (occurrence index `1`).

## Remove imported events

Use the menu → **Remove** → choose your `sourceId`.
Only events tagged with `[SRC: <sourceId>]` are removed.

## Notes

* `.ics` and `.ical` are both accepted (same iCalendar format).
* The importer writes a JSON temp file to `/tmp/<source>_events.json`.
* If you modify the AppleScript or JXA files, re-run the manager — no rebuild needed.

