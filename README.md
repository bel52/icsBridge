# icsBridge — v1.1

A lightweight toolchain to import public ICS calendars into Outlook for Mac with the correct time handling, persistent defaults, and a reliable way to surgically remove imported events later.

---

## 🚀 What it does

**Normalizes and imports ICS feeds:**

* Converts all *timed events* to UTC (Z) → Outlook shows the correct local time (avoids double-shift bugs).
* Strips `X-WR-TIMEZONE` from calendar headers so Outlook doesn’t reinterpret.
* Leaves **all-day events untouched**.
* Tags each event’s description with `[SRC: <your-id>]` so they can be removed precisely.
* Opens the normalized `.ics` in Outlook’s **native importer** (stable/known-good on macOS).

**Removes imported events:**

* Scans your chosen Outlook calendar.
* Deletes only events containing `[SRC: <id>]` in the description.
* Works even if the source is no longer tracked.

**Persists defaults:**

* Stores your chosen target Outlook calendar name and occurrence index in `.icsbridge_config`.
* Reuses these automatically on subsequent runs.

---

## 🖥 Requirements

* macOS with Microsoft Outlook (tested on current builds of *New Outlook*).
* Python 3.9+ (uses stdlib `zoneinfo`).
* Python packages (auto-installed into `.venv` on first run):

  * `icalendar`
  * `python-dateutil`

---

## ⚙️ Install / First Run

```bash
git clone https://github.com/bel52/icsBridge.git
cd icsBridge
./ics_manager.sh
# 4) Set Default Target Calendar
#   Name:  Calendar
#   Index: 2
```

---

## ➕ Add a calendar

```bash
./ics_manager.sh
# 1) Add Calendar via Outlook Import
#   URL: <ICS feed URL>
#   ID:  <short id, e.g., lions>
```

* A normalized ICS is written to `/tmp/<id>.ics`.
* Outlook’s import dialog opens — select your default calendar (e.g., *Calendar* #2).
* The source is recorded in `.tracked_sources.json` and `.sources/<id>.json`.

---

## ➖ Remove a calendar’s events

```bash
./ics_manager.sh
# 2) Remove Imported Calendar
#   a) Pick from tracked list
#   b) Or enter SRC ID manually
```

* Uses `outlook_delete_by_src.applescript` to delete events containing `[SRC: <id>]` in your target calendar.

---

## ❓ Why UTC?

Outlook for Mac often misinterprets combinations of:

* `X-WR-TIMEZONE`
* property `TZID`
* local system settings

This causes **off-by-hours bugs**. Writing events in UTC is the most reliable path — Outlook consistently renders them correctly in local time.

---

## 🔧 Troubleshooting

* `ics_manager.sh` is verbose — no screen clearing; every step is shown.
* To sanity-check a normalized ICS:

  * No `X-WR-TIMEZONE` header
  * All timed events’ `DTSTART`/`DTEND` end with `Z` (UTC)
* If events still look wrong:

  * Remove them and re-import
  * Confirm you’re importing the **fresh** `/tmp/<id>.ics`

---

## 🔒 Security note

* `prepare_ics_for_import.py` fetches feeds over HTTPS.
* Currently uses an **unverified SSL context** for compatibility with some ICS hosts.
* For stricter security: switch to default verified context or pin trusted hosts.

---

## 📦 Repo Layout (v1.1)

* `ics_manager.sh` — main menu tool; normalize → open Outlook importer; removal by tag; persistent defaults.
* `prepare_ics_for_import.py` — fetch + normalize ICS to UTC; tag descriptions.
* `outlook_delete_by_src.applescript` — delete events by `[SRC: <id>]`.
* `.icsbridge_config` — stored defaults (calendar name + index).
* `.tracked_sources.json` — append-only log of sources.
* `.sources/` — per-source markers for recovery.
* `requirements.txt` — Python dependencies.
* `VERSION` — current version string (e.g., `1.1.0`).

---

## 📝 Changelog

**1.1.0**

* Improved handling of persistent defaults.
* Cleaner verbose logging (no screen clears).
* Hardened removal path (works even without tracked source file).
* Streamlined repo (removed redundant test scripts).

**1.0.0**

* Normalize timed events to UTC; strip `X-WR-TIMEZONE`.
* Persist target calendar defaults.
* Stable import via Outlook’s native dialog.
* One-shot removal by `[SRC: id]`.
* Repo cleanup.

---

## 📜 License

(Choose an appropriate license; e.g., MIT.)

---

👉 Repo: [https://github.com/bel52/icsBridge](https://github.com/bel52/icsBridge)
