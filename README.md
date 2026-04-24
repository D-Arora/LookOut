# Outlook Shared Calendar → Google Calendar Sync

---

## For recipients (no Node required)

You need the `dist/` folder. It contains a native binary and a bundled
Chromium — nothing else to install.

**Windows** — double-click `run.bat`, or in a terminal:
```
run.bat
run.bat --days 60
```

**Mac / Linux**
```bash
chmod +x run.sh     # first time only
./run.sh
./run.sh --days 60
```

A Chrome window opens. Log in, browse the calendar months you want,
press Enter. `events.ics` is saved next to the binary.

**Import into Google Calendar**
1. Open calendar.google.com
2. Settings ⚙ → Import & Export → Import
3. Choose `events.ics`, pick your target calendar, click Import

---

## For developers — building the binary

`node build.mjs` builds a binary **for the current platform only**.
pkg cannot cross-compile on Windows (known upstream bug).

```bash
npm install
npx playwright install chromium   # ~170 MB, one-time
node build.mjs                    # produces dist/
```

**To build for all platforms at once**, use GitHub Actions — push the
repo and the included workflow (`.github/workflows/build.yml`) will
build on real Windows, Mac (ARM + Intel), and Linux runners and
attach the binaries to a release. Trigger it by pushing a version tag:

```bash
git tag v1.0.0 && git push origin v1.0.0
```

Or trigger it manually from the Actions tab without a tag.

---

## CLI options

| Flag | Default | Description |
|------|---------|-------------|
| `--days N` | 30 | How many days ahead to capture |
| `--out FILE` | events.ics | Output filename |
| `--folder-id ID` | (hardcoded) | Override the calendar folder filter |
| `--category-colors FILE` | — | JSON map of category name → hex colour |
| `--diagnose` | false | Dump raw event data and folder map, then exit |

**Custom category colours** (`colours.json`):
```json
{
  "Lecture": "#0078D4",
  "Practical": "#107C10",
  "Tutorial": "#D83B01"
}
```
```bash
./run.sh --category-colors colours.json
```

---

## Troubleshooting

**"No events were converted"**
→ Browse in Month view and page through the months you want before pressing Enter.

**Wrong folder / missing events**
→ Run `node sync.js --diagnose`, check `diagnose_folders.json` to find the
  right folder ID, then pass it via `--folder-id`.

**Timings are wrong after import**
→ Ensure you're using the latest `events.ics`. The file embeds an
  `Australia/Sydney` VTIMEZONE block that Google Calendar understands correctly.
