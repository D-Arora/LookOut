#!/usr/bin/env node
/**
 * Outlook Shared Calendar → Google Calendar Sync
 *
 * Usage:
 *   node sync.js                  # interactive login, saves events.ics
 *   node sync.js --out my.ics     # custom output filename
 *   node sync.js --folder-id ...  # override folder filter
 */

import { chromium } from "playwright";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { parseArgs } from "util";

const { values: args } = parseArgs({
  options: {
    out: { type: "string", default: "events.ics" },
    url: { type: "string", default: "https://outlook.office.com/calendar" },
    "folder-id": { type: "string" },
    diagnose: { type: "boolean", default: false },
  },
});

let OUTPUT_FILE = args.out;
if (OUTPUT_FILE === "events.ics") {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, "0");
  // +1 because getMonth() is 0-indexed
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yy = String(now.getFullYear()).slice(-2);

  // We use " - " instead of " | " since the pipe character
  // is invalid/forbidden in Windows filenames
  const baseName = `Events - ${dd}-${mm}-${yy}`;

  OUTPUT_FILE = `${baseName}.ics`;
  let counter = 1;
  while (existsSync(OUTPUT_FILE)) {
    OUTPUT_FILE = `${baseName} (${counter}).ics`;
    counter++;
  }
}

const OWA_URL = args.url;

// ANU MD 29' folder id discovered from diagnose_folders.json in this account.
// Override at runtime with: --folder-id <ParentFolderId.Id>

function getEventCategories(ev) {
  const raw = ev.Categories ?? ev.categories ?? [];
  if (!Array.isArray(raw)) return [];
  return raw
    .map((c) => {
      if (typeof c === "string") return c.trim();
      if (c && typeof c === "object") {
        return String(c.DisplayName || c.Name || c.name || "").trim();
      }
      return "";
    })
    .filter(Boolean);
}

// ── Windows TZ ID → IANA TZ ID ───────────────────────────────────────────────
const WINDOWS_TO_IANA = {
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
  "Singapore Standard Time": "Asia/Singapore",
};

function toIana(tzId) {
  if (!tzId) return "UTC";
  if (tzId.includes("/") && !WINDOWS_TO_IANA[tzId]) return tzId;
  return WINDOWS_TO_IANA[tzId] ?? tzId;
}

// ── VTIMEZONE block ───────────────────────────────────────────────────────────
function makeVTimezone(ianaId) {
  if (ianaId === "UTC") {
    return "BEGIN:VTIMEZONE\r\nTZID:UTC\r\nBEGIN:STANDARD\r\nDTSTART:19700101T000000\r\nTZOFFSETFROM:+0000\r\nTZOFFSETTO:+0000\r\nTZNAME:UTC\r\nEND:STANDARD\r\nEND:VTIMEZONE";
  }

  function offsetStr(date) {
    const fmt = new Intl.DateTimeFormat("en", {
      timeZone: ianaId,
      timeZoneName: "shortOffset",
    });
    const raw =
      fmt.formatToParts(date).find((p) => p.type === "timeZoneName")?.value ??
      "GMT+0";
    const m = raw.match(/GMT([+-])(\d+)(?::(\d+))?/);
    if (!m) return "+0000";
    return `${m[1]}${m[2].padStart(2, "0")}${(m[3] ?? "0").padStart(2, "0")}`;
  }

  function tzName(date) {
    const fmt = new Intl.DateTimeFormat("en", {
      timeZone: ianaId,
      timeZoneName: "short",
    });
    return (
      fmt.formatToParts(date).find((p) => p.type === "timeZoneName")?.value ??
      ianaId
    );
  }

  const jan = new Date("2025-01-15T12:00:00Z");
  const jul = new Date("2025-07-15T12:00:00Z");
  const oJan = offsetStr(jan),
    oJul = offsetStr(jul);
  const nJan = tzName(jan),
    nJul = tzName(jul);

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
      "END:VTIMEZONE",
    ].join("\r\n");
  }

  const southernDST = oJan > oJul;
  const [stdOff, , stdName] = southernDST
    ? [oJul, oJan, nJul]
    : [oJan, oJul, nJan];
  const [dstOff, , dstName] = southernDST
    ? [oJan, oJul, nJan]
    : [oJul, oJan, nJul];

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
    "END:VTIMEZONE",
  ].join("\r\n");
}

// ── iCal helpers ──────────────────────────────────────────────────────────────

function toIcalUtc(dateStr) {
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  return d
    .toISOString()
    .replace(/[-:]/g, "")
    .replace(/\.\d{3}/, "");
}

function toIcalLocal(dateStr) {
  return dateStr.replace(/[-:]/g, "").slice(0, 15);
}

function esc(str = "") {
  const normalized = String(str).replace(/\r\n?/g, "\n");
  return normalized
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

function bodyToPlainText(rawBody = "") {
  const text = String(rawBody);
  const withBreaks = text
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n");
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

// ── Event converter ───────────────────────────────────────────────────────────

function eventToVEvent(ev) {
  const uid =
    ev.UID || ev.iCalUId || ev.id || ev.ItemId?.Id || crypto.randomUUID();
  const summary = ev.Subject || ev.subject || "(No title)";

  const start =
    (typeof ev.Start === "string" ? ev.Start : ev.Start?.DateTime) ||
    ev.start?.dateTime ||
    ev.StartDate ||
    null;
  const end =
    (typeof ev.End === "string" ? ev.End : ev.End?.DateTime) ||
    ev.end?.dateTime ||
    ev.EndDate ||
    null;

  const winTzId =
    ev.StartTimeZoneId || ev.Start?.TimeZone || ev.start?.timeZone || null;
  const ianaId = toIana(winTzId);

  const location = ev.Location?.DisplayName || ev.location?.displayName || "";
  const categories = getEventCategories(ev);

  // Body: GetCalendarEvent returns TextBody (plain) or Body.Value (may be HTML).
  // GetCalendarView only has Preview (truncated). We prefer the full body if available.
  const rawBody =
    ev.TextBody || ev.Body?.Value || ev.body?.content || ev.Preview || "";
  const plainBody = bodyToPlainText(rawBody);

  const allDay = ev.IsAllDayEvent ?? ev.isAllDay ?? false;

  if (!start) {
    console.warn(`  ⚠️  Skipping "${summary}" — no start time found.`);
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
        console.warn(`  ⚠️  Skipping "${summary}" — bad date: ${start}`);
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
    endStr,
  ];
  if (location) lines.push(fold(`LOCATION:${esc(location)}`));
  if (plainBody) lines.push(fold(`DESCRIPTION:${esc(plainBody)}`));
  if (categories.length) {
    lines.push(fold(`CATEGORIES:${esc(categories.join(","))}`));
  }

  lines.push(`DTSTAMP:${toIcalUtc(new Date().toISOString())}`);
  lines.push("END:VEVENT");

  return { vevent: lines.join("\r\n"), ianaId };
}

// ── .ics builder ──────────────────────────────────────────────────────────────

function buildIcs(results) {
  const tzIds = [
    ...new Set(results.filter((r) => r.ianaId).map((r) => r.ianaId)),
  ];
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//outlook-calendar-sync//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    ...tzIds.map(makeVTimezone),
    ...results.map((r) => r.vevent),
    "END:VCALENDAR",
  ].join("\r\n");
}

// ── Event extractor ───────────────────────────────────────────────────────────

function extractEvents(json) {
  if (!json || typeof json !== "object") return [];
  if (json.Body && Array.isArray(json.Body.Items))
    return json.Body.Items.filter(
      (i) => i.__type?.startsWith("CalendarItem") || i.Subject !== undefined,
    );
  if (json.Body?.CalendarItem) return [json.Body.CalendarItem];
  if (Array.isArray(json.value) && json.value[0]?.subject !== undefined)
    return json.value;
  try {
    const msgs = json?.Body?.ResponseMessages?.Items ?? [];
    const events = [];
    for (const msg of msgs)
      events.push(...(msg?.RootFolder?.Items ?? msg?.Items ?? []));
    if (events.length) return events;
  } catch (_) {}
  if (Array.isArray(json)) return json.flatMap(extractEvents);
  return [];
}

// ── Fetch full event bodies via GetCalendarEvent ──────────────────────────────
// GetCalendarView returns only a truncated Preview field.
// GetCalendarEvent returns TextBody (full plain-text body).
// We replay one POST per event using the same session cookies Playwright holds.

async function fetchFullBodies(page, owaBaseUrl, events) {
  // Extract the base OWA service URL from the page
  // e.g. https://outlook.office.com/owa/service.svc
  const serviceUrl = `${new URL(owaBaseUrl).origin}/owa/service.svc`;

  // We need the canary token that OWA uses for CSRF protection.
  // It's stored in a cookie called "X-OWA-CANARY".
  const cookies = await page.context().cookies();
  const canary = cookies.find((c) => c.name === "X-OWA-CANARY")?.value ?? "";

  // Also need the action counter — we'll just use a high number to avoid collisions
  let n = 500;

  const total = events.length;
  let fetched = 0;
  let failed = 0;

  // Process in batches to avoid hammering the server
  const BATCH = 5;

  for (let i = 0; i < events.length; i += BATCH) {
    const batch = events.slice(i, i + BATCH);

    await Promise.all(
      batch.map(async (ev) => {
        const itemId = ev.ItemId?.Id;
        const changeKey = ev.ItemId?.ChangeKey;
        if (!itemId) return; // no item ID, can't fetch

        try {
          const body = {
            __type: "GetCalendarEventJsonRequest:#Exchange",
            Header: {
              __type: "JsonRequestHeaders:#Exchange",
              RequestServerVersion: "V2018_01_08",
            },
            Body: {
              __type: "GetCalendarEventRequest:#Exchange",
              ItemId: {
                __type: "ItemId:#Exchange",
                Id: itemId,
                ChangeKey: changeKey,
              },
              // Request the body in text format so we don't have to strip HTML
              AdditionalProperties: [
                { __type: "PropertyUri:#Exchange", FieldURI: "TextBody" },
              ],
            },
          };

          const resp = await page.evaluate(
            async ({ url, canary, n, body }) => {
              const r = await fetch(
                `${url}?action=GetCalendarEvent&app=Calendar&n=${n}`,
                {
                  method: "POST",
                  headers: {
                    "Content-Type": "application/json; charset=utf-8",
                    Action: "GetCalendarEvent",
                    "X-OWA-CANARY": canary,
                    "X-OWA-ActionId": String(n),
                  },
                  body: JSON.stringify(body),
                  credentials: "include",
                },
              );
              return r.ok ? r.json() : null;
            },
            { url: serviceUrl, canary, n: n++, body },
          );

          if (resp?.Body?.CalendarItem) {
            const full = resp.Body.CalendarItem;
            // Merge TextBody (and anything else useful) back onto the original event
            if (full.TextBody) ev.TextBody = full.TextBody;
            if (full.Body) ev.Body = full.Body;
            fetched++;
          }
        } catch (_) {
          failed++;
        }
      }),
    );

    // Progress indicator
    const done = Math.min(i + BATCH, total);
    process.stdout.write(`\r  📬  Fetching descriptions… ${done}/${total}`);

    // Small delay between batches to be polite to the server
    if (i + BATCH < events.length) {
      await new Promise((r) => setTimeout(r, 300));
    }
  }

  process.stdout.write("\n");
  if (fetched > 0)
    console.log(`  ✅  Got full descriptions for ${fetched} event(s)`);
  if (failed > 0)
    console.log(
      `  ⚠️  Failed to fetch ${failed} event(s) — they'll use Preview text`,
    );
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  console.log("🚀  Launching browser — please sign in when prompted.");
  console.log(`    Output → ${OUTPUT_FILE}\n`);

  const browser = await chromium.launch({ headless: false, channel: "chrome" });
  const context = await browser.newContext();
  const page = await context.newPage();
  const collectedEvents = new Map();
  let skippedByFolder = 0;

  // Track the OWA origin so we can replay API calls later
  let owaCapturedOrigin = OWA_URL;

  page.on("response", async (response) => {
    const ct = response.headers()["content-type"] ?? "";
    const status = response.status();
    if (!ct.includes("json") || status !== 200) return;
    try {
      const json = await response.json();
      const events = extractEvents(json);
      if (events.length) {
        // Capture the actual origin in case of redirects
        owaCapturedOrigin = response.url();
        const action =
          new URL(response.url()).searchParams.get("action") ??
          response.url().split("/").pop();
        console.log(
          `  ✅  Captured ${events.length} event(s) from ?action=${action}`,
        );
        for (const ev of events) {
          const key =
            ev.UID ||
            ev.iCalUId ||
            ev.id ||
            ev.ItemId?.Id ||
            JSON.stringify(ev).slice(0, 80);
          collectedEvents.set(key, ev);
        }
      }
    } catch (_) {}
  });

  await page.goto(OWA_URL, { waitUntil: "domcontentloaded" });
  console.log(
    "👤  Waiting for you to log in and navigate to the shared calendar…",
  );
  console.log("    Browse to different weeks/months to capture more events.");
  console.log("    Press ENTER in this terminal when you're done browsing.\n");

  await new Promise((resolve) => {
    process.stdin.once("data", resolve);
    process.stdin.resume();
  });
  process.stdin.pause();

  // ── Diagnose mode ────────────────────────────────────────────────────────
  if (args.diagnose) {
    const all = [...collectedEvents.values()];
    if (!all.length) {
      console.error("No events captured — browse the calendar first.");
      await browser.close();
      waitAndExit(1);
      return;
    }

    // Dump full raw sample for field inspection
    writeFileSync("diagnose.json", JSON.stringify(all[0], null, 2), "utf8");

    // Dump a folder map: folderId → [event subjects] so we can identify ANU MD 29'
    const folderMap = {};
    for (const ev of all) {
      const fid = ev.ParentFolderId?.Id ?? "unknown";
      if (!folderMap[fid]) folderMap[fid] = [];
      folderMap[fid].push(ev.Subject ?? ev.subject ?? "(no title)");
    }
    writeFileSync(
      "diagnose_folders.json",
      JSON.stringify(folderMap, null, 2),
      "utf8",
    );

    console.log("\n🔍  Raw event → diagnose.json");
    console.log("📁  Folder map → diagnose_folders.json");
    console.log(
      `\n    ${all.length} events across ${Object.keys(folderMap).length} folder(s)`,
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
        `\n⚠️   Specified --folder-id not found in captured events. Select one instead:`,
      );
      targetFolderId = null;
    }

    if (!targetFolderId) {
      if (folderIds.length === 1) {
        targetFolderId = folderIds[0];
      } else {
        console.log(`\n📂  Detected events from multiple calendars:`);
        folderIds.forEach((fid, idx) => {
          // get up to 3 unique event titles
          const subjects = [...new Set(folderMap[fid])].slice(0, 3).join(", ");
          console.log(
            `    [${idx + 1}] ${folderMap[fid].length} events — e.g. ${subjects}\n`,
          );
        });

        while (!targetFolderId) {
          process.stdout.write(
            `\nType a number [1-${folderIds.length}] to select which calendar you would like to save and press [Enter]: `,
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

    // Filter down collected events
    for (const [key, ev] of collectedEvents.entries()) {
      const fid = ev.ParentFolderId?.Id ?? "unknown";
      if (fid !== targetFolderId) {
        collectedEvents.delete(key);
        skippedByFolder++;
      }
    }
  }

  // ── Fetch full descriptions before closing the browser ──────────────────
  const eventsArray = [...collectedEvents.values()];
  if (eventsArray.length > 0) {
    console.log(
      `\n📬  Fetching full descriptions for ${eventsArray.length} events…`,
    );
    console.log(
      "    (browser will stay open briefly — please don't close it)\n",
    );
    await fetchFullBodies(page, owaCapturedOrigin, eventsArray);
  }

  await browser.close();

  // ── Build .ics ────────────────────────────────────────────────────────────
  console.log(`\n📦  Processing ${collectedEvents.size} unique event(s)…`);
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
      "\n⚠️   No events were converted. Try browsing in Month view.\n",
    );
    waitAndExit(1);
    return;
  }

  const tzUsed = [...new Set(results.map((r) => r.ianaId))];
  console.log(`    Timezones detected: ${tzUsed.join(", ")}`);

  const descCount = results.filter((r) =>
    r.vevent.includes("DESCRIPTION:"),
  ).length;
  console.log(`    Events with descriptions: ${descCount}/${results.length}`);

  writeFileSync(OUTPUT_FILE, buildIcs(results), "utf8");

  console.log(`\n✅  Saved ${results.length} event(s) to ${OUTPUT_FILE}`);
  console.log("\n📅  To import into Google Calendar:");
  console.log("    1. Open calendar.google.com");
  console.log("    2. Settings (⚙) → Import & Export → Import");
  console.log(`    3. Choose ${OUTPUT_FILE} and select target calendar`);
  console.log("    4. Click Import\n");
  waitAndExit(0);
}

function waitAndExit(code = 0) {
  console.log("\nPress [Enter] to exit...");
  process.stdin.resume();
  process.stdin.once("data", () => process.exit(code));
}

main().catch((err) => {
  console.error("Fatal error:", err);
  waitAndExit(1);
});
