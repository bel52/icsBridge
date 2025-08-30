# icsBridge — v1.0

A tiny toolchain to import public ICS calendars into **Outlook for Mac** with the **correct time** and a reliable way to **remove** those events later.

---

## What it does

* **Normalizes** any ICS feed:

  * Converts all *timed* events to **UTC** (`Z`) so Outlook shows the right local time (avoids double‑shift bugs).
  * Removes calendar-level `X-WR-TIMEZONE` so Outlook doesn’t second-guess.
  * Tags each event description with **`[SRC: <your-id>]`** for surgical removal later.
  * Leaves **all-day** events untouched.
* **Imports** via Outlook’s native ICS import dialog (stable/known‑good on macOS).
* **Removes** previously imported events by scanning your chosen Outlook calendar and deleting any event whose description contains `[SRC: <your-id>]`.
* **Persists defaults**: target Outlook **calendar name** and **occurrence index** are stored and reused so you’re not prompted every time.

---

## Requirements

* macOS with **Microsoft Outlook** (new Outlook works; tested on recent builds).
* Python **3.9+** (uses stdlib `zoneinfo`).
* Python packages (installed automatically into `.venv` on first run):

  * `icalendar`
  * `python-dateutil`

---

## Install / First Run

```bash
git clone https://github.com/bel52/icsBridge.git
cd icsBridge
./ics_manager.sh
# Choose: 4) Set Default Target Calendar
#   Name:  Calendar
#   Index: 2
```

---

## Add a calendar

```bash
./ics_manager.sh
# 1) Add Calendar via Outlook Import
#   URL: <your ICS feed>
#   ID:  <short id, e.g., lions>
```

* The tool normalizes the ICS to UTC and writes `/tmp/<id>.ics`.
* Outlook’s import window opens. Select your default calendar (e.g., “Calendar” #2) and confirm.
* The source is tracked in `.tracked_sources.json` **and** `.sources/<id>.json`.

---

## Remove a calendar’s events

```bash
./ics_manager.sh
# 2) Remove Imported Calendar
#   a) pick a tracked entry, or
#   b) enter the SRC ID (even if not tracked)
```

Removal uses `outlook_delete_by_src.applescript` to delete events in your target calendar whose **description** contains `[SRC: <id>]`.

---

## Why UTC?

Outlook for Mac can misinterpret combos of `X-WR-TIMEZONE`, property `TZID`, and local settings, causing off-by-hours. Writing *timed* events in **UTC** is the most reliable path—Outlook renders them correctly in your local zone.

---

## Troubleshooting

* **See what’s happening**: `ics_manager.sh` prints verbose steps; no screen clearing.
* **Sanity-check a normalized ICS**:

  * It should have **no `X-WR-TIMEZONE`**.
  * `DTSTART`/`DTEND` should end with **`Z`** (UTC).
* **Still wrong times?**

  * Re-remove and re-import.
  * Make sure you’re importing the freshly normalized `/tmp/<id>.ics`.

---

## Security note

`prepare_ics_for_import.py` fetches feeds over HTTPS and currently uses an **unverified SSL context** for compatibility with some ICS hosts. For stricter security, switch to the default verified context or pin hosts—easy tweak.

---

## What’s in v1.0

* **Stable import path** (Outlook’s native importer).
* **UTC-normalized ICS** to avoid tz confusion.
* **One-shot removal** by `[SRC: id]`.
* **Persistent defaults** for calendar name/index.
* Leaned repo (removed legacy/unneeded scripts).

---

## Repo Layout (v1.0)

* `ics_manager.sh` — menu tool; normalize → open Outlook importer; removal by tag; persistent defaults.
* `prepare_ics_for_import.py` — fetch + normalize ICS to UTC; tag descriptions.
* `outlook_delete_by_src.applescript` — delete by `[SRC: <id>]` in chosen calendar.
* `.icsbridge_config` — stored defaults (calendar name + index).
* `.tracked_sources.json` — newline-JSON tracker (append-only log).
* `.sources/` — per-source markers (redundant/handy for recovery).
* `requirements.txt` — Python deps (venv installs automatically).
* `VERSION` — current version string (e.g., `1.0.0`).

---

## Changelog

### 1.0.0

* Normalize all timed events to UTC; strip `X-WR-TIMEZONE`.
* Persist target calendar name + index.
* Stable import via Outlook dialog (no JXA writes).
* Robust removal by `[SRC: <id>]`.
* Cleaned up legacy scripts.

---

## License

(Choose one appropriate for your project; e.g., MIT.)

EOF
