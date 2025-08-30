// Minimal creator: create events into target calendar with [SRC]/[ICSUID] tags.
// No scanning, no updates, no categories — reduces macOS privacy friction.
function run(argv){
  if(argv.length<5){
    console.log(JSON.stringify({error:"Usage: outlook_upsert_events.js <jsonPath> <calendarName> <occurrenceIndex> <categoryIgnored> <sourceId>"}));
    return;
  }
  const [jsonPath, calName, occStr, _categoryIgnored, sourceId] = argv;
  const idx = Math.max(1, parseInt(occStr,10) || 1);

  const app = Application('Microsoft Outlook');
  app.includeStandardAdditions = true;

  // Read parsed JSON
  const jsonText = app.doShellScript(`/bin/cat '${jsonPath.replace(/'/g,"'\\''")}'`);
  const data = JSON.parse(jsonText);
  const events = (data && data.events) || [];

  // Find Nth calendar named calName
  const cal = findCal(app, calName, idx);
  if(!cal){
    console.log(JSON.stringify({error:`Calendar "${calName}" (#${idx}) not found`}));
    return;
  }

  // Create each event (minimal properties only)
  let created = 0, failed = 0;
  events.forEach(it => {
    const uid = (it.uid || "").trim();
    if(!uid) { failed++; return; }

    const start = parseIso(it.start);
    const end   = it.end ? parseIso(it.end) : new Date(start.getTime() + 3600000);
    const allDay = !!it.all_day;

    const details = (it.description || "").trim();
    const tagBlock = `[SRC: ${sourceId}]\n[ICSUID: ${uid}]`;
    const body = details ? `${details}\n\n${tagBlock}` : tagBlock;

    try {
      const newEv = app.CalendarEvent({
        subject: it.summary || "(No title)",
        location: it.location || "",
        content: body,
        calendar: cal
      });
      app.make({ new: newEv, at: cal });

      if (allDay) {
        const s = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0,0,0);
        const e = new Date(s.getTime() + 86400000);
        newEv.allDayEvent = true;
        newEv.startTime = s;
        newEv.endTime   = e;
      } else {
        newEv.allDayEvent = false;
        newEv.startTime = start;
        newEv.endTime   = end;
      }

      created++;
    } catch (e) {
      // Include a tiny hint if it's a privilege error
      const msg = String(e);
      if (msg.indexOf("-10004") !== -1) {
        console.log(JSON.stringify({error:"privilege_violation", hint:"Grant your terminal app Automation permission to control Microsoft Outlook (System Settings → Privacy & Security → Automation). Also ensure Legacy Outlook.", item: it.summary || "(No title)"}));
      }
      failed++;
    }
  });

  console.log(JSON.stringify({ok:true, created, failed, processed: events.length}));
}

// Helpers
function parseIso(s){ return s.endsWith('Z') ? new Date(s) : new Date(s); }

function findCal(app, name, nth) {
  const matches = [];
  try { app.calendars().forEach(c => { try { if (c.name() === name) matches.push(c); } catch(e){} }); } catch(e){}
  try {
    app.accounts().forEach(a => {
      try { a.calendars().forEach(c => { if (c.name() === name) matches.push(c); }); } catch(e){}
    });
  } catch(e){}
  return matches[nth-1] || null;
}
