// outlook_remove_source.js
// Scans a calendar and removes events with a specific [SRC: <id>] tag.
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

  // Scan a wide date window
  const now = new Date();
  const startScan = new Date(now.getTime() - 365 * 24 * 3600 * 1000); // 1 year past
  const endScan   = new Date(now.getTime() + 365 * 24 * 3600 * 1000); // 1 year future

  const existing = eventsInRange(app, cal, startScan, endScan);
  let deleted = 0;

  existing.forEach(ev => {
    const body = safe(() => ev.content()) || "";
    if (body.includes(`[SRC: ${sourceId}]`)) {
      try { 
        ev.delete();
        deleted++;
      } catch (e) { /* ignore errors on deletion */ }
    }
  });

  console.log(JSON.stringify({ ok: true, deleted: deleted, checked: existing.length }));
}

// Helper functions
function safe(fn) { try { return fn(); } catch (e) { return null; } }

function findCal(app, name, nth) {
  const matches = [];
  try { app.calendars().forEach(c => { try { if (c.name() === name) matches.push(c); } catch (e) {} }); } catch (e) {}
  return matches[nth - 1] || null;
}

function eventsInRange(app, cal, start, end) {
  try {
    return cal.calendarEvents().filter(ev => {
      const s = safe(() => ev.startTime());
      return s && s >= start && s <= end;
    });
  } catch (e) {
    return []; // Fail safely
  }
}
