/**
 * build.mjs — produces platform binaries in ./dist/
 *
 * Run once on your own machine (needs Node 20+ and internet access):
 *   npm install
 *   node build.mjs
 *
 * Output:
 *   dist/
 *     LookOut-macos-arm64   ← Apple Silicon Mac
 *     LookOut-macos-x64     ← Intel Mac
 *     LookOut-win-x64.exe   ← Windows
 *     LookOut-linux-x64     ← Linux
 *     pw-browsers/                        ← Playwright browser (ship alongside binary)
 *
 * The binary + pw-browsers/ folder must stay together.
 * Share as a zip: the binary and the pw-browsers/ folder.
 */

import { execSync } from "child_process";
import {
  existsSync,
  mkdirSync,
  cpSync,
  readdirSync,
  readFileSync,
  writeFileSync,
} from "fs";
import { join, resolve } from "path";
import { fileURLToPath } from "url";
import * as resedit from "resedit";
import * as p2i from "png2icons";

const __dir = fileURLToPath(new URL(".", import.meta.url));
const dist = join(__dir, "dist");

// ── Step 1: bundle sync.js (ESM) → bundle.cjs (CommonJS) via esbuild ─────────
// pkg requires CommonJS. We mark playwright as external so it stays as a
// require() call at runtime — the binary will look for it next to itself.
console.log("📦  Bundling sync.js → bundle.cjs …");
execSync(
  "npx esbuild sync.js " +
    "--bundle " +
    "--platform=node " +
    "--format=cjs " +
    "--external:playwright " +
    "--external:fs " +
    "--external:util " +
    "--outfile=bundle.cjs",
  { stdio: "inherit" },
);

// ── Step 2: run pkg on the bundle ─────────────────────────────────────────────
// pkg cannot cross-compile on Windows (spawn UNKNOWN bug in fabricator.js —
// it tries to run macOS/Linux strip/codesign tools that don't exist on Windows).
// Solution: detect the host platform and only build that platform's binary.
// To build all platforms, run this script on each OS, or use GitHub Actions
// (see README for the workflow file).
console.log("\n🔨  Compiling binary for current platform via pkg …");
mkdirSync(dist, { recursive: true });

const platform = process.platform;
const arch = process.arch;

const hostTarget =
  platform === "win32"
    ? { target: "node20-win-x64", out: "LookOut.exe" }
    : platform === "darwin"
      ? arch === "arm64"
        ? { target: "node20-macos-arm64", out: "LookOut" }
        : { target: "node20-macos-x64", out: "LookOut" }
      : { target: "node20-linux-x64", out: "LookOut" };

console.log(`  → ${hostTarget.target}`);
execSync(
  `npx @yao-pkg/pkg bundle.cjs --target ${hostTarget.target} --output dist/${hostTarget.out}`,
  { stdio: "inherit" },
);

// Optional: Inject icon if assets/LookOut.png exists
const iconPngPath = join(__dir, "assets", "LookOut.png");
if (existsSync(iconPngPath)) {
  console.log("🎨  Injecting LookOut.png icon...");
  const pngData = readFileSync(iconPngPath);

  if (platform === "win32") {
    const icoData = p2i.createICO(pngData, p2i.BICUBIC2, 0, false);
    if (icoData) {
      const exeData = readFileSync(join(dist, hostTarget.out));
      const exe = resedit.NtExecutable.from(exeData);
      const res = resedit.NtExecutableResource.from(exe);
      const iconFile = resedit.Data.IconFile.from(icoData);
      resedit.Resource.IconGroupEntry.replaceIconsForResource(
        res.entries,
        1,
        1033,
        iconFile.icons.map((item) => item.data),
      );
      res.outputResource(exe);
      writeFileSync(join(dist, hostTarget.out), Buffer.from(exe.generate()));
      console.log(
        "  ✅  Windows .ico injected (Note: Windows caches icons! If you don't see it, rename or move the file to refresh).",
      );
    }
  }
}

const targets = [hostTarget]; // used below when copying browser and writing launchers

// ── Step 3: copy Playwright browser alongside each binary ─────────────────────
// Playwright looks for browsers relative to the executable via PLAYWRIGHT_BROWSERS_PATH
// or the default ~/.cache/ms-playwright. We ship them in pw-browsers/ next to the binary
// and set the env var in a tiny launcher script.
console.log("\n🌐  Locating Playwright browser cache …");

// Find where playwright installed its browsers
let pwCache = null;
try {
  const result = execSync("npx playwright --version", { encoding: "utf8" });
  // Default cache locations
  const candidates = [
    join(process.env.HOME ?? "", ".cache", "ms-playwright"),
    join(process.env.LOCALAPPDATA ?? "", "ms-playwright"), // Windows
    join(__dir, "node_modules", "playwright", ".local-browsers"),
  ];
  for (const c of candidates) {
    if (existsSync(c) && readdirSync(c).length > 0) {
      pwCache = c;
      break;
    }
  }
} catch (_) {}

if (!pwCache) {
  console.log("  ⚠️  Could not locate Playwright browser cache.");
  console.log("     Run: npx playwright install chromium");
  console.log("     Then re-run this build script.\n");
} else {
  console.log(`  Found browser cache at: ${pwCache}`);

  if (platform === "darwin") {
    // For macOS, we bundle into an app wrapper
    const appDir = join(dist, "LookOut.app");
    const macOsDir = join(appDir, "Contents", "MacOS");
    const resourcesDir = join(appDir, "Contents", "Resources");
    mkdirSync(macOsDir, { recursive: true });
    mkdirSync(resourcesDir, { recursive: true });

    // Copy binary into .app
    cpSync(join(dist, "LookOut"), join(macOsDir, "LookOut-bin"));

    // Inject macOS App Icon if we have LookOut.png
    let iconKey = "";
    if (existsSync(iconPngPath)) {
      const pngData = readFileSync(iconPngPath);
      const icnsData = p2i.createICNS(pngData, p2i.BICUBIC2, 0);
      if (icnsData) {
        writeFileSync(join(resourcesDir, "AppIcon.icns"), icnsData);
        iconKey = "<key>CFBundleIconFile</key><string>AppIcon.icns</string>";
        console.log("  ✅  macOS .icns injected.");
      }
    }

    // Default Info.plist
    writeFileSync(
      join(appDir, "Contents", "Info.plist"),
      `<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0"><dict><key>CFBundleExecutable</key><string>LookOut</string><key>CFBundleIdentifier</key><string>com.example.lookout</string><key>CFBundleName</key><string>LookOut</string>${iconKey}<key>CFBundleVersion</key><string>1.0</string><key>CFBundlePackageType</key><string>APPL</string><key>LSUIElement</key><true/></dict></plist>`,
      "utf8",
    );

    const pwDest = join(macOsDir, "pw-browsers");
    console.log(`  Copying browser to: ${pwDest} …`);
    cpSync(pwCache, pwDest, { recursive: true });
    console.log("  ✅  Browser copied into app.");

    // Create the macOS Launcher script that opens a terminal window so the CLI stays visible
    const macLauncher = [
      "#!/bin/bash",
      'DIR="$(cd "$(dirname "$0")" && pwd)"',
      'export PLAYWRIGHT_BROWSERS_PATH="$DIR/pw-browsers"',
      'if [ "$1" != "--launched-in-terminal" ]; then',
      '  osascript -e "tell application \\"Terminal\\"" \\',
      '            -e "  do script \\"cd \\\\\\"$HOME/Desktop\\\\\\"; \\\\\\"$DIR/LookOut-bin\\\\\\" \\\\\\"--launched-in-terminal\\\\\\"\\"" \\',
      '            -e "  activate" \\',
      '            -e "end tell"',
      "  exit 0",
      "fi",
      'cd "$HOME/Desktop"', // fallbacks to running in current context if they somehow hit the flag
      '"$DIR/LookOut-bin" "$@"',
    ].join("\n");
    writeFileSync(join(macOsDir, "LookOut"), macLauncher, { mode: 0o755 });
    console.log("  Wrote macOS App bundle launcher");
  } else {
    const pwDest = join(dist, "pw-browsers");
    console.log(`  Copying to: ${pwDest} …`);
    cpSync(pwCache, pwDest, { recursive: true });
    console.log("  ✅  Browser copied.");
  }
}

// ── Step 4: generate launcher scripts ─────────────────────────────────────────
// The binary needs PLAYWRIGHT_BROWSERS_PATH pointed at pw-browsers/.
// A tiny shell/bat wrapper handles this so the user can just double-click.

// Write launcher for the platform we just built
if (platform === "win32") {
  const winLauncher = `@echo off\nset DIR=%~dp0\nset PLAYWRIGHT_BROWSERS_PATH=%DIR%pw-browsers\n"%DIR%LookOut.exe" %*\n`;
  writeFileSync(join(dist, "run.bat"), winLauncher);
  console.log("  Wrote run.bat");
} else if (platform === "darwin") {
  console.log("  Skipping run.sh for macOS — use LookOut.app instead.");
} else {
  const macLinuxLauncher = [
    "#!/bin/bash",
    "# Launcher — sets Playwright browser path relative to this script",
    'DIR="$(cd "$(dirname "$0")" && pwd)"',
    'export PLAYWRIGHT_BROWSERS_PATH="$DIR/pw-browsers"',
    `"$DIR/${hostTarget.out}" "$@"`,
    "",
  ].join("\n");
  writeFileSync(join(dist, "run.sh"), macLinuxLauncher, { mode: 0o755 });
  console.log("  Wrote run.sh");
}

console.log("\n✅  Build complete! Contents of dist/:");
for (const f of readdirSync(dist)) {
  console.log("   ", f);
}

console.log(`
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  To distribute:
    Zip the entire dist/ folder and share it.
    Recipients run:
      Mac:        LookOut.app
      Linux:      ./run.sh
      Windows:    run.bat   (or double-click LookOut.exe)
    No Node, npm, or Playwright setup needed.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`);
