#!/usr/bin/env node

// sync.js
var import_playwright = require("playwright");
var import_fs = require("fs");
var import_util = require("util");
var { values: args } = (0, import_util.parseArgs)({
  options: {
    out: { type: "string", default: "events.ics" },
    url: { type: "string", default: "https://outlook.office.com/calendar" },
    "folder-id": { type: "string" },
    diagnose: { type: "boolean", default: false }
  }
});
var OUTPUT_FILE = args.out;
var OWA_URL = args.url;
function getEventCategories(ev) {
  const raw = ev.Categories ?? ev.categories ?? [];
  if (!Array.isArray(raw)) return [];
  return raw.map((c) => {
    if (typeof c === "string") return c.trim();
    if (c && typeof c === "object") {
      return String(c.DisplayName || c.Name || c.name || "").trim();
    }
    return "";
  }).filter(Boolean);
}
var WINDOWS_TO_IANA = {
  "AUS Eastern Standard Time": "Australia/Sydney",
  "AUS Central Standard Time": "Australia/Darwin",
  "AUS Central W. Standard Time": "Australia/Eucla",
  "E. Australia Standard Time": "Australia/Brisbane",
  "Cen. Australia Standard Time": "Australia/Adelaide",
  "Tasmania Standard Time": "Australia/Hobart",
  "W. Australia Standard Time": "Australia/Perth",
  "Lord Howe Standard Time": "Australia/Lord_Howe",
  UTC: "UTC",
  "Eastern Standard Time": "America/New_York",
  "Central Standard Time": "America/Chicago",
  "Mountain Standard Time": "America/Denver",
  "Pacific Standard Time": "America/Los_Angeles",
  "GMT Standard Time": "Europe/London",
  "Romance Standard Time": "Europe/Paris",
  "W. Europe Standard Time": "Europe/Berlin",
  "New Zealand Standard Time": "Pacific/Auckland",
  "India Standard Time": "Asia/Calcutta",
  "China Standard Time": "Asia/Shanghai",
  "Tokyo Standard Time": "Asia/Tokyo",
  "Singapore Standard Time": "Asia/Singapore"
};
function toIana(tzId) {
  if (!tzId) return "UTC";
  if (tzId.includes("/") && !WINDOWS_TO_IANA[tzId]) return tzId;
  return WINDOWS_TO_IANA[tzId] ?? tzId;
}
function makeVTimezone(ianaId) {
  if (ianaId === "UTC") {
    return "BEGIN:VTIMEZONE\r\nTZID:UTC\r\nBEGIN:STANDARD\r\nDTSTART:19700101T000000\r\nTZOFFSETFROM:+0000\r\nTZOFFSETTO:+0000\r\nTZNAME:UTC\r\nEND:STANDARD\r\nEND:VTIMEZONE";
  }
  function offsetStr(date) {
    const fmt = new Intl.DateTimeFormat("en", {
      timeZone: ianaId,
      timeZoneName: "shortOffset"
    });
    const raw = fmt.formatToParts(date).find((p) => p.type === "timeZoneName")?.value ?? "GMT+0";
    const m = raw.match(/GMT([+-])(\d+)(?::(\d+))?/);
    if (!m) return "+0000";
    return `${m[1]}${m[2].padStart(2, "0")}${(m[3] ?? "0").padStart(2, "0")}`;
  }
  function tzName(date) {
    const fmt = new Intl.DateTimeFormat("en", {
      timeZone: ianaId,
      timeZoneName: "short"
    });
    return fmt.formatToParts(date).find((p) => p.type === "timeZoneName")?.value ?? ianaId;
  }
  const jan = /* @__PURE__ */ new Date("2025-01-15T12:00:00Z");
  const jul = /* @__PURE__ */ new Date("2025-07-15T12:00:00Z");
  const oJan = offsetStr(jan), oJul = offsetStr(jul);
  const nJan = tzName(jan), nJul = tzName(jul);
  if (oJan === oJul) {
    return [
      "BEGIN:VTIMEZONE",
      `TZID:${ianaId}`,
      "BEGIN:STANDARD",
      "DTSTART:19700101T000000",
      `TZOFFSETFROM:${oJan}`,
      `TZOFFSETTO:${oJan}`,
      `TZNAME:${nJan}`,
      "END:STANDARD",
      "END:VTIMEZONE"
    ].join("\r\n");
  }
  const southernDST = oJan > oJul;
  const [stdOff, , stdName] = southernDST ? [oJul, oJan, nJul] : [oJan, oJul, nJan];
  const [dstOff, , dstName] = southernDST ? [oJan, oJul, nJan] : [oJul, oJan, nJul];
  return [
    "BEGIN:VTIMEZONE",
    `TZID:${ianaId}`,
    "BEGIN:STANDARD",
    "DTSTART:19700101T030000",
    `TZOFFSETFROM:${dstOff}`,
    `TZOFFSETTO:${stdOff}`,
    `TZNAME:${stdName}`,
    "RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=4",
    "END:STANDARD",
    "BEGIN:DAYLIGHT",
    "DTSTART:19701001T020000",
    `TZOFFSETFROM:${stdOff}`,
    `TZOFFSETTO:${dstOff}`,
    `TZNAME:${dstName}`,
    "RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=10",
    "END:DAYLIGHT",
    "END:VTIMEZONE"
  ].join("\r\n");
}
function toIcalUtc(dateStr) {
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  return d.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "");
}
function toIcalLocal(dateStr) {
  return dateStr.replace(/[-:]/g, "").slice(0, 15);
}
function esc(str = "") {
  const normalized = String(str).replace(/\r\n?/g, "\n");
  return normalized.replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
}
function bodyToPlainText(rawBody = "") {
  const text = String(rawBody);
  const withBreaks = text.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<\/div>/gi, "\n");
  const withoutTags = withBreaks.replace(/<[^>]+>/g, "");
  return withoutTags.replace(/\r\n?/g, "\n").trim();
}
function fold(line) {
  if (line.length <= 75) return line;
  const chunks = [line.slice(0, 75)];
  let i = 75;
  while (i < line.length) {
    chunks.push(" " + line.slice(i, i + 74));
    i += 74;
  }
  return chunks.join("\r\n");
}
function eventToVEvent(ev) {
  const uid = ev.UID || ev.iCalUId || ev.id || ev.ItemId?.Id || crypto.randomUUID();
  const summary = ev.Subject || ev.subject || "(No title)";
  const start = (typeof ev.Start === "string" ? ev.Start : ev.Start?.DateTime) || ev.start?.dateTime || ev.StartDate || null;
  const end = (typeof ev.End === "string" ? ev.End : ev.End?.DateTime) || ev.end?.dateTime || ev.EndDate || null;
  const winTzId = ev.StartTimeZoneId || ev.Start?.TimeZone || ev.start?.timeZone || null;
  const ianaId = toIana(winTzId);
  const location = ev.Location?.DisplayName || ev.location?.displayName || "";
  const categories = getEventCategories(ev);
  const rawBody = ev.TextBody || ev.Body?.Value || ev.body?.content || ev.Preview || "";
  const plainBody = bodyToPlainText(rawBody);
  const allDay = ev.IsAllDayEvent ?? ev.isAllDay ?? false;
  if (!start) {
    console.warn(`  \u26A0\uFE0F  Skipping "${summary}" \u2014 no start time found.`);
    return null;
  }
  let startStr, endStr;
  if (allDay) {
    startStr = `DTSTART;VALUE=DATE:${start.slice(0, 10).replace(/-/g, "")}`;
    endStr = `DTEND;VALUE=DATE:${(end || start).slice(0, 10).replace(/-/g, "")}`;
  } else {
    const hasOffset = /Z$|[+-]\d{2}:\d{2}$|[+-]\d{4}$/.test(start.trim());
    if (hasOffset) {
      const s = toIcalUtc(start);
      const e = toIcalUtc(end || start);
      if (!s) {
        console.warn(`  \u26A0\uFE0F  Skipping "${summary}" \u2014 bad date: ${start}`);
        return null;
      }
      startStr = `DTSTART:${s}`;
      endStr = `DTEND:${e || s}`;
    } else {
      startStr = `DTSTART;TZID=${ianaId}:${toIcalLocal(start)}`;
      endStr = `DTEND;TZID=${ianaId}:${toIcalLocal(end || start)}`;
    }
  }
  const lines = [
    "BEGIN:VEVENT",
    fold(`UID:${esc(uid)}`),
    fold(`SUMMARY:${esc(summary)}`),
    startStr,
    endStr
  ];
  if (location) lines.push(fold(`LOCATION:${esc(location)}`));
  if (plainBody) lines.push(fold(`DESCRIPTION:${esc(plainBody)}`));
  if (categories.length) {
    lines.push(fold(`CATEGORIES:${esc(categories.join(","))}`));
  }
  lines.push(`DTSTAMP:${toIcalUtc((/* @__PURE__ */ new Date()).toISOString())}`);
  lines.push("END:VEVENT");
  return { vevent: lines.join("\r\n"), ianaId };
}
function buildIcs(results) {
  const tzIds = [
    ...new Set(results.filter((r) => r.ianaId).map((r) => r.ianaId))
  ];
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//outlook-calendar-sync//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    ...tzIds.map(makeVTimezone),
    ...results.map((r) => r.vevent),
    "END:VCALENDAR"
  ].join("\r\n");
}
function extractEvents(json) {
  if (!json || typeof json !== "object") return [];
  if (json.Body && Array.isArray(json.Body.Items))
    return json.Body.Items.filter(
      (i) => i.__type?.startsWith("CalendarItem") || i.Subject !== void 0
    );
  if (json.Body?.CalendarItem) return [json.Body.CalendarItem];
  if (Array.isArray(json.value) && json.value[0]?.subject !== void 0)
    return json.value;
  try {
    const msgs = json?.Body?.ResponseMessages?.Items ?? [];
    const events = [];
    for (const msg of msgs)
      events.push(...msg?.RootFolder?.Items ?? msg?.Items ?? []);
    if (events.length) return events;
  } catch (_) {
  }
  if (Array.isArray(json)) return json.flatMap(extractEvents);
  return [];
}
async function fetchFullBodies(page, owaBaseUrl, events) {
  const serviceUrl = `${new URL(owaBaseUrl).origin}/owa/service.svc`;
  const cookies = await page.context().cookies();
  const canary = cookies.find((c) => c.name === "X-OWA-CANARY")?.value ?? "";
  let n = 500;
  const total = events.length;
  let fetched = 0;
  let failed = 0;
  const BATCH = 5;
  for (let i = 0; i < events.length; i += BATCH) {
    const batch = events.slice(i, i + BATCH);
    await Promise.all(
      batch.map(async (ev) => {
        const itemId = ev.ItemId?.Id;
        const changeKey = ev.ItemId?.ChangeKey;
        if (!itemId) return;
        try {
          const body = {
            __type: "GetCalendarEventJsonRequest:#Exchange",
            Header: {
              __type: "JsonRequestHeaders:#Exchange",
              RequestServerVersion: "V2018_01_08"
            },
            Body: {
              __type: "GetCalendarEventRequest:#Exchange",
              ItemId: {
                __type: "ItemId:#Exchange",
                Id: itemId,
                ChangeKey: changeKey
              },
              // Request the body in text format so we don't have to strip HTML
              AdditionalProperties: [
                { __type: "PropertyUri:#Exchange", FieldURI: "TextBody" }
              ]
            }
          };
          const resp = await page.evaluate(
            async ({ url, canary: canary2, n: n2, body: body2 }) => {
              const r = await fetch(
                `${url}?action=GetCalendarEvent&app=Calendar&n=${n2}`,
                {
                  method: "POST",
                  headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    Action: "GetCalendarEvent",
                    "X-OWA-CANARY": canary2,
                    "X-OWA-ActionId": String(n2)
                  },
                  body: JSON.stringify(body2),
                  credentials: "include"
                }
              );
              return r.ok ? r.json() : null;
            },
            { url: serviceUrl, canary, n: n++, body }
          );
          if (resp?.Body?.CalendarItem) {
            const full = resp.Body.CalendarItem;
            if (full.TextBody) ev.TextBody = full.TextBody;
            if (full.Body) ev.Body = full.Body;
            fetched++;
          }
        } catch (_) {
          failed++;
        }
      })
    );
    const done = Math.min(i + BATCH, total);
    process.stdout.write(`\r  \u{1F4EC}  Fetching descriptions\u2026 ${done}/${total}`);
    if (i + BATCH < events.length) {
      await new Promise((r) => setTimeout(r, 300));
    }
  }
  process.stdout.write("\n");
  if (fetched > 0)
    console.log(`  \u2705  Got full descriptions for ${fetched} event(s)`);
  if (failed > 0)
    console.log(
      `  \u26A0\uFE0F  Failed to fetch ${failed} event(s) \u2014 they'll use Preview text`
    );
}
async function main() {
  console.log("\u{1F680}  Launching browser \u2014 please sign in when prompted.");
  console.log(`    Output \u2192 ${OUTPUT_FILE}
`);
  const browser = await import_playwright.chromium.launch({ headless: false, channel: "chrome" });
  const context = await browser.newContext();
  const page = await context.newPage();
  const collectedEvents = /* @__PURE__ */ new Map();
  let skippedByFolder = 0;
  let owaCapturedOrigin = OWA_URL;
  page.on("response", async (response) => {
    const ct = response.headers()["content-type"] ?? "";
    const status = response.status();
    if (!ct.includes("json") || status !== 200) return;
    try {
      const json = await response.json();
      const events = extractEvents(json);
      if (events.length) {
        owaCapturedOrigin = response.url();
        const action = new URL(response.url()).searchParams.get("action") ?? response.url().split("/").pop();
        console.log(
          `  \u2705  Captured ${events.length} event(s) from ?action=${action}`
        );
        for (const ev of events) {
          const key = ev.UID || ev.iCalUId || ev.id || ev.ItemId?.Id || JSON.stringify(ev).slice(0, 80);
          collectedEvents.set(key, ev);
        }
      }
    } catch (_) {
    }
  });
  await page.goto(OWA_URL, { waitUntil: "domcontentloaded" });
  console.log(
    "\u{1F464}  Waiting for you to log in and navigate to the shared calendar\u2026"
  );
  console.log("    Browse to different weeks/months to capture more events.");
  console.log("    Press ENTER in this terminal when you're done browsing.\n");
  await new Promise((resolve) => {
    process.stdin.once("data", resolve);
    process.stdin.resume();
  });
  process.stdin.pause();
  if (args.diagnose) {
    const all = [...collectedEvents.values()];
    if (!all.length) {
      console.error("No events captured \u2014 browse the calendar first.");
      await browser.close();
      waitAndExit(1);
      return;
    }
    (0, import_fs.writeFileSync)("diagnose.json", JSON.stringify(all[0], null, 2), "utf8");
    const folderMap = {};
    for (const ev of all) {
      const fid = ev.ParentFolderId?.Id ?? "unknown";
      if (!folderMap[fid]) folderMap[fid] = [];
      folderMap[fid].push(ev.Subject ?? ev.subject ?? "(no title)");
    }
    (0, import_fs.writeFileSync)(
      "diagnose_folders.json",
      JSON.stringify(folderMap, null, 2),
      "utf8"
    );
    console.log("\n\u{1F50D}  Raw event \u2192 diagnose.json");
    console.log("\u{1F4C1}  Folder map \u2192 diagnose_folders.json");
    console.log(
      `
    ${all.length} events across ${Object.keys(folderMap).length} folder(s)`
    );
    await browser.close();
    waitAndExit(0);
    return;
  }
  const allEvents = [...collectedEvents.values()];
  let targetFolderId = args["folder-id"];
  if (allEvents.length > 0) {
    const folderMap = {};
    for (const ev of allEvents) {
      const fid = ev.ParentFolderId?.Id ?? "unknown";
      if (!folderMap[fid]) folderMap[fid] = [];
      folderMap[fid].push(ev.Subject ?? ev.subject ?? "(no title)");
    }
    const folderIds = Object.keys(folderMap);
    if (targetFolderId && !folderIds.includes(targetFolderId)) {
      console.log(
        `
\u26A0\uFE0F   Specified --folder-id not found in captured events. Select one instead:`
      );
      targetFolderId = null;
    }
    if (!targetFolderId) {
      if (folderIds.length === 1) {
        targetFolderId = folderIds[0];
      } else {
        console.log(`
\u{1F4C2}  Detected events from multiple calendars:`);
        folderIds.forEach((fid, idx) => {
          const subjects = [...new Set(folderMap[fid])].slice(0, 3).join(", ");
          console.log(
            `    [${idx + 1}] ${folderMap[fid].length} events \u2014 e.g. ${subjects}`
          );
        });
        while (!targetFolderId) {
          process.stdout.write(
            `
Type a number [1-${folderIds.length}] and press Enter: `
          );
          process.stdin.resume();
          const answer = await new Promise((resolve) => {
            process.stdin.once("data", (d) => resolve(d.toString().trim()));
          });
          process.stdin.pause();
          const n = parseInt(answer, 10);
          if (!isNaN(n) && n >= 1 && n <= folderIds.length) {
            targetFolderId = folderIds[n - 1];
          }
        }
      }
    }
    for (const [key, ev] of collectedEvents.entries()) {
      const fid = ev.ParentFolderId?.Id ?? "unknown";
      if (fid !== targetFolderId) {
        collectedEvents.delete(key);
        skippedByFolder++;
      }
    }
  }
  const eventsArray = [...collectedEvents.values()];
  if (eventsArray.length > 0) {
    console.log(
      `
\u{1F4EC}  Fetching full descriptions for ${eventsArray.length} events\u2026`
    );
    console.log(
      "    (browser will stay open briefly \u2014 please don't close it)\n"
    );
    await fetchFullBodies(page, owaCapturedOrigin, eventsArray);
  }
  await browser.close();
  console.log(`
\u{1F4E6}  Processing ${collectedEvents.size} unique event(s)\u2026`);
  if (targetFolderId) console.log(`    Folder filter: ${targetFolderId}`);
  if (skippedByFolder > 0)
    console.log(`    Skipped (other calendars): ${skippedByFolder}`);
  const results = [];
  for (const ev of collectedEvents.values()) {
    const r = eventToVEvent(ev);
    if (r) results.push(r);
  }
  if (results.length === 0) {
    console.error(
      "\n\u26A0\uFE0F   No events were converted. Try browsing in Month view.\n"
    );
    waitAndExit(1);
    return;
  }
  const tzUsed = [...new Set(results.map((r) => r.ianaId))];
  console.log(`    Timezones detected: ${tzUsed.join(", ")}`);
  const descCount = results.filter(
    (r) => r.vevent.includes("DESCRIPTION:")
  ).length;
  console.log(`    Events with descriptions: ${descCount}/${results.length}`);
  (0, import_fs.writeFileSync)(OUTPUT_FILE, buildIcs(results), "utf8");
  console.log(`
\u2705  Saved ${results.length} event(s) to ${OUTPUT_FILE}`);
  console.log("\n\u{1F4C5}  To import into Google Calendar:");
  console.log("    1. Open calendar.google.com");
  console.log("    2. Settings (\u2699) \u2192 Import & Export \u2192 Import");
  console.log(`    3. Choose ${OUTPUT_FILE} and select target calendar`);
  console.log("    4. Click Import\n");
  waitAndExit(0);
}
function waitAndExit(code = 0) {
  console.log("\nPress Enter to exit...");
  process.stdin.resume();
  process.stdin.once("data", () => process.exit(code));
}
main().catch((err) => {
  console.error("Fatal error:", err);
  waitAndExit(1);
});
