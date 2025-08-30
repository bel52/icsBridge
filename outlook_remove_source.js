// outlook_remove_source.js
// Remove all Outlook events tagged with a given [SRC: <sourceId>] from a specific calendar.
//
// Usage:
//   osascript -l JavaScript outlook_remove_source.js "Calendar" 2 "detroit-lions-2025"
//
function run(argv) {
  if (argv.length < 3) {
    console.log(JSON.stringify({ error: "Usage: outlook_remove_source.js <calendarName> <index> <sourceId>" }));
    return;
  }
  const [calName, occStr, sourceId] = argv;
  const occurrenceIndex = Math.max(1, parseInt(occStr, 10) || 1);

  const app = Application('Microsoft Outlook');
  app.includeStandardAdditions = true;

  const cal = findCal(app, calName, occurrenceIndex);
  if (!cal) {
    console.log(JSON.stringify({ error: `Calendar "${calName}" (#${occurrenceIndex}) not found` }));
    return;
  }

  // Scan a wide date window to catch seasons/schedules
  const now = new Date();
  const startScan = new Date(now.getTime() - 500 * 24 * 3600 * 1000);
  const endScan   = new Date(now.getTime() + 500 * 24 * 3600 * 1000);

  const existing = eventsInRange(app, cal, startScan, endScan);
  let deleted = 0, checked = 0;

  existing.forEach(ev => {
    const body = safe(() => ev.content()) || "";
    checked++;
    if (body.includes(`[SRC: ${sourceId}]`)) {
      try { ev.delete(); deleted++; } catch (e) { /* ignore */ }
    }
  });

  console.log(JSON.stringify({ ok: true, deleted, checked }));
}

// ---------- helpers ----------
function safe(fn) { try { return fn(); } catch (e) { return null; } }

function findCal(app, name, nth) {
  const matches = [];
  try {
    app.calendars().forEach(c => { try { if (c.name() === name) matches.push(c); } catch (e) {} });
  } catch (e) {}
  try {
    app.accounts().forEach(acc => {
      try { acc.calendars().forEach(c => { if (c.name() === name) matches.push(c); }); } catch (e) {}
    });
  } catch (e) {}
  return matches[nth - 1] || null;
}

function eventsInRange(app, cal, start, end) {
  // Prefer calendar-scoped query; fall back to app-wide filter
  try {
    return cal.calendarEvents().filter(ev => {
      const s = safe(() => ev.startTime());
      return s && s >= start && s <= end;
    });
  } catch (e) {
    try {
      return app.calendarEvents().filter(ev => {
        const s = safe(() => ev.startTime());
        return ev.calendar() && ev.calendar().id() === cal.id() && s && s >= start && s <= end;
      });
    } catch (e2) {
      return [];
    }
  }
}
