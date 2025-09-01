# icsBridge v1.1.0 — 2025-09-01

## What's new
- **Persistent tracking** outside the repo (~/.icsbridge) so tracked calendars survive code changes.
- **Rebuild tracking** reads `[SRC: …]` tags from Outlook and repopulates the tracker.
- **Local .ics import** normalizes times to UTC and stamps `[SRC: <id>]`.
- **Remove by ID** reliably deletes only events stamped with that source.
- **Defaults** (calendar name + index) stored in `.icsbridge_config`.

## Notes
- Events are **stored in Outlook in UTC** (import-time conversion), but show correctly in **Eastern Time** in Outlook.
- The persistent tracker is at `~/.icsbridge/tracked.jsonl`; keep it backed up if you care.

