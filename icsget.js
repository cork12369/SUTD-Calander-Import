(() => {
  // Installs/overwrites window.exportSUTDToICS
  const DEFAULTS = {
    weeks: 14,
    direction: "forward",      // "forward" (Next Week) or "backward" (Previous Week)
    clickDelayMs: 2200,        // bump to 3000 if slow
    tzid: "Asia/Singapore",
    filenamePrefix: "sutd_schedule",
  };

  const TIME_RE = /\b(\d{1,2}):(\d{2})\s*(AM|PM)\s*-\s*(\d{1,2}):(\d{2})\s*(AM|PM)\b/i;
  const COURSE_LINE_RE = /^\s*\d{1,2}\s*\.\s*\d{3}\s*-\s*[A-Z0-9]{2,}\b/i;

  function getScheduleDocument() {
    const iframe =
      document.querySelector("#ptifrmtgtframe") ||
      document.querySelector('iframe[name="TargetContent"]');
    if (!iframe) return document;
    return iframe.contentDocument || iframe.contentWindow.document;
  }

  function textOf(el) {
    return (el?.innerText || el?.textContent || "").trim();
  }

  function wait(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function formatICSLocal(dt) {
    return (
      dt.getFullYear() +
      pad2(dt.getMonth() + 1) +
      pad2(dt.getDate()) +
      "T" +
      pad2(dt.getHours()) +
      pad2(dt.getMinutes()) +
      pad2(dt.getSeconds())
    );
  }

  function escapeICS(s) {
    return String(s || "")
      .replace(/\\/g, "\\\\")
      .replace(/\n/g, "\\n")
      .replace(/;/g, "\\;")
      .replace(/,/g, "\\,")
      .trim();
  }

  function foldLine(line) {
    const limit = 75;
    if (line.length <= limit) return line;
    let out = "";
    for (let i = 0; i < line.length; i += limit) {
      out += (i === 0 ? "" : "\r\n ") + line.slice(i, i + limit);
    }
    return out;
  }

  function downloadText(filename, content) {
    const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function singaporeVTimezoneLines(tzid) {
    return [
      "BEGIN:VTIMEZONE",
      `TZID:${tzid}`,
      "X-LIC-LOCATION:Asia/Singapore",
      "BEGIN:STANDARD",
      "TZOFFSETFROM:+0800",
      "TZOFFSETTO:+0800",
      "TZNAME:SGT",
      "DTSTART:19700101T000000",
      "END:STANDARD",
      "END:VTIMEZONE",
    ];
  }

  function djb2(str) {
    let h = 5381;
    for (let i = 0; i < str.length; i++) h = ((h << 5) + h) + str.charCodeAt(i);
    return (h >>> 0).toString(16);
  }

  function getScheduleTable(doc) {
    return doc.querySelector("#WEEKLY_SCHED_HTMLAREA");
  }

  function getWeekLabel(doc) {
    const t = textOf(doc.body);
    const m = t.match(/Week of\s*[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}[^]*?(?=\n|$)/i);
    return (m ? m[0] : "").slice(0, 80);
  }

  function parseWeekStart(doc) {
    const t = textOf(doc.body);
    const m = t.match(/Week of\s*([0-9]{1,2})\/([0-9]{1,2})\/([0-9]{4})/i);
    if (!m) return null;
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    const yy = parseInt(m[3], 10);
    return new Date(yy, mm - 1, dd, 0, 0, 0, 0);
  }

  function addDays(d, n) {
    const x = new Date(d.getTime());
    x.setDate(x.getDate() + n);
    return x;
  }

  function parseTimeParts(hStr, mStr, ampm) {
    let h = parseInt(hStr, 10);
    const m = parseInt(mStr, 10);
    const up = ampm.toUpperCase();
    if (up === "AM") {
      if (h === 12) h = 0;
    } else {
      if (h !== 12) h += 12;
    }
    return { h, m };
  }

  function getDayHeaderCenters(table) {
    const dayRe = /\bMonday\b|\bTuesday\b|\bWednesday\b|\bThursday\b|\bFriday\b|\bSaturday\b|\bSunday\b/i;
    const headers = Array.from(table.querySelectorAll("th"))
      .filter((th) => dayRe.test(textOf(th)))
      .map((th) => {
        const r = th.getBoundingClientRect();
        return { cx: r.left + r.width / 2 };
      })
      .sort((a, b) => a.cx - b.cx)
      .slice(0, 7);

    return headers.map(h => h.cx);
  }

  function dayIndexFromCell(td, dayCenters) {
    const r = td.getBoundingClientRect();
    const cx = r.left + r.width / 2;
    let best = 0, bestDist = Infinity;
    for (let i = 0; i < dayCenters.length; i++) {
      const d = Math.abs(cx - dayCenters[i]);
      if (d < bestDist) { bestDist = d; best = i; }
    }
    return best; // 0..6 (Mon..Sun)
  }

  function cleanCourseLine(s) {
    return s.replace(/\s*\.\s*/g, ".").replace(/\s+/g, " ").trim();
  }

  function extractEventFromSpan(spanEl, weekStart, dayCenters) {
    const lines = textOf(spanEl).split("\n").map(s => s.trim()).filter(Boolean);
    if (!lines.length) return null;
    if (!COURSE_LINE_RE.test(lines[0])) return null;

    const tIdx = lines.findIndex(l => TIME_RE.test(l));
    if (tIdx === -1) return null;

    const m = lines[tIdx].match(TIME_RE);
    if (!m) return null;

    const { h: sh, m: sm } = parseTimeParts(m[1], m[2], m[3]);
    const { h: eh, m: em } = parseTimeParts(m[4], m[5], m[6]);

    const td = spanEl.closest("td");
    if (!td) return null;

    const dayIdx = dayIndexFromCell(td, dayCenters);
    const date = addDays(weekStart, dayIdx);

    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), sh, sm, 0);
    const end   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), eh, em, 0);

    const course = cleanCourseLine(lines[0]);
    const title = lines[1] || "";
    const kind = lines[2] || "";
    const summary = [course, title, kind].filter(Boolean).join(" - ");

    const after = lines.slice(tIdx + 1);
    const location = after.find(x => x && !/^Instructors?:?/i.test(x)) || "";

    return { summary, location, description: lines.join("\n"), start, end };
  }

  function scrapeWeek(doc) {
    const weekStart = parseWeekStart(doc);
    if (!weekStart) throw new Error("Can't find 'Week of dd/mm/yyyy' on the page.");

    const table = getScheduleTable(doc);
    if (!table) throw new Error("Can't find #WEEKLY_SCHED_HTMLAREA.");

    const dayCenters = getDayHeaderCenters(table);
    if (dayCenters.length < 7) throw new Error("Can't locate Mon–Sun headers reliably.");

    const spans = Array.from(table.querySelectorAll("td > span"));
    const events = [];

    for (const sp of spans) {
      const ev = extractEventFromSpan(sp, weekStart, dayCenters);
      if (ev) events.push(ev);
    }

    // de-dupe within week
    const seen = new Set();
    const deduped = [];
    for (const e of events) {
      const key = `${e.summary}__${e.start.toISOString()}__${e.end.toISOString()}__${e.location}`;
      if (seen.has(key)) continue;
      seen.add(key);
      deduped.push(e);
    }

    return { weekStart, events: deduped };
  }

  function buildICS(allEvents, tzid) {
    const now = new Date();
    const dtstamp = formatICSLocal(now);

    const lines = [];
    lines.push("BEGIN:VCALENDAR");
    lines.push("VERSION:2.0");
    lines.push("PRODID:-//SUTD//MyPortal Weekly Schedule Export//EN");
    lines.push("CALSCALE:GREGORIAN");
    lines.push("METHOD:PUBLISH");
    lines.push(...singaporeVTimezoneLines(tzid));

    for (const e of allEvents) {
      const uidBase = `${e.summary}|${e.start.toISOString()}|${e.end.toISOString()}|${e.location}`;
      const uid = `${djb2(uidBase)}@sutd-myportal`;

      lines.push("BEGIN:VEVENT");
      lines.push(`UID:${uid}`);
      lines.push(`DTSTAMP:${dtstamp}`);
      lines.push(`DTSTART;TZID=${tzid}:${formatICSLocal(e.start)}`);
      lines.push(`DTEND;TZID=${tzid}:${formatICSLocal(e.end)}`);
      lines.push(`SUMMARY:${escapeICS(e.summary)}`);
      if (e.location) lines.push(`LOCATION:${escapeICS(e.location)}`);
      if (e.description) lines.push(`DESCRIPTION:${escapeICS(e.description)}`);
      lines.push("END:VEVENT");
    }

    lines.push("END:VCALENDAR");
    return lines.map(foldLine).join("\r\n") + "\r\n";
  }

  async function gotoNextWeek(doc, direction, clickDelayMs, beforeWeekLabel) {
    const nextId = direction === "backward" ? "DERIVED_CLASS_S_SSR_PREV_WEEK" : "DERIVED_CLASS_S_SSR_NEXT_WEEK";
    const btn = doc.getElementById(nextId) || Array.from(doc.querySelectorAll("input,button,a"))
      .find(el => ((el.value || textOf(el)) || "").includes(direction === "backward" ? "Previous Week" : "Next Week"));

    if (!btn) throw new Error("Can't find Next/Previous Week button in the schedule frame.");

    btn.click();

    // Wait until week label changes AND table has some spans
    for (let i = 0; i < 40; i++) {
      await wait(Math.max(250, clickDelayMs / 6));
      const after = getWeekLabel(doc);
      const table = getScheduleTable(doc);
      const spanCount = table ? table.querySelectorAll("td > span").length : 0;

      if (after && after !== beforeWeekLabel && spanCount > 0) return;
    }

    // Fallback: just wait a bit more
    await wait(clickDelayMs);
  }

  window.exportSUTDToICS = async function (opts = {}) {
    const cfg = { ...DEFAULTS, ...opts };
    const doc = getScheduleDocument();

    const all = [];
    let firstWeekStart = null;

    for (let i = 0; i < cfg.weeks; i++) {
      const beforeLabel = getWeekLabel(doc);

      const { weekStart, events } = scrapeWeek(doc);
      if (!firstWeekStart) firstWeekStart = weekStart;

      console.log(`[${i + 1}/${cfg.weeks}] ${weekStart.toDateString()} → ${events.length} events`);
      all.push(...events);

      if (i < cfg.weeks - 1) {
        await gotoNextWeek(doc, cfg.direction, cfg.clickDelayMs, beforeLabel);
      }
    }

    // de-dupe across weeks
    const seen = new Set();
    const deduped = [];
    for (const e of all) {
      const key = `${e.summary}__${e.start.toISOString()}__${e.end.toISOString()}__${e.location}`;
      if (seen.has(key)) continue;
      seen.add(key);
      deduped.push(e);
    }
    deduped.sort((a, b) => a.start - b.start);

    const fnameDate = firstWeekStart
      ? `${firstWeekStart.getFullYear()}${pad2(firstWeekStart.getMonth() + 1)}${pad2(firstWeekStart.getDate())}`
      : "export";
    const filename = `${cfg.filenamePrefix}_${fnameDate}_${cfg.weeks}w.ics`;

    downloadText(filename, buildICS(deduped, cfg.tzid));
    console.log(`✅ Exported ${deduped.length} events → ${filename}`);
    return deduped;
  };

  console.log("✅ Installed exportSUTDToICS. Now run: exportSUTDToICS({ weeks: 14 })");
})();
