/**
 * build.mjs — produces platform binaries in ./dist/
 *
 * Run via GitHub Actions (recommended) or locally on each platform:
 *   npm install
 *   npx playwright install chromium
 *   node build.mjs
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
import { join } from "path";
import { fileURLToPath } from "url";

const __dir = fileURLToPath(new URL(".", import.meta.url));
const dist = join(__dir, "dist");

const platform = process.platform;
const arch = process.arch;

// ── Step 1: esbuild — ESM → CJS ───────────────────────────────────────────────
console.log("📦  Bundling sync.js → bundle.cjs …");
execSync(
  "npx esbuild sync.js --bundle --platform=node --format=cjs " +
    "--external:playwright --external:fs --external:util --outfile=bundle.cjs",
  { stdio: "inherit" },
);

// ── Step 2: compile binary ────────────────────────────────────────────────────
console.log("\n🔨  Compiling binary …");
mkdirSync(dist, { recursive: true });

const hostTarget =
  platform === "win32"
    ? { target: "node20-win-x64", out: "LookOut.exe" }
    : platform === "darwin"
      ? {
          target: arch === "arm64" ? "node20-macos-arm64" : "node20-macos-x64",
          out: "LookOut",
        }
      : { target: "node20-linux-x64", out: "LookOut" };

console.log(`  → ${hostTarget.target}`);

const iconPngPath = join(__dir, "assets", "LookOut.png");
const iconIcoPath = join(__dir, "assets", "LookOut.ico");

if (platform === "win32") {
  // Windows: use pkg-exe-build which handles icon embedding internally,
  // before the snapshot is sealed — all other PE patching tools (rcedit,
  // resedit) corrupt the exe by modifying it after pkg locks it.
  const { default: p2i } = await import("png2icons");
  const pngData = readFileSync(iconPngPath);
  const icoData = p2i.createICO(pngData, p2i.BICUBIC2, 0, true);
  writeFileSync(iconIcoPath, icoData);

  const { default: pkgBuild } = await import("pkg-exe-build");
  await pkgBuild({
    entry: "bundle.cjs",
    out: join(dist, hostTarget.out),
    target: hostTarget.target,
    icon: iconIcoPath,
    properties: {
      FileDescription: "LookOut Calendar Sync",
      ProductName: "LookOut",
      OriginalFilename: "LookOut.exe",
    },
  });
  console.log("  ✅  Windows binary compiled with icon");
} else {
  // macOS / Linux: use @yao-pkg/pkg directly
  execSync(
    `npx @yao-pkg/pkg bundle.cjs --target ${hostTarget.target} --output dist/${hostTarget.out}`,
    { stdio: "inherit" },
  );
}

// ── Step 3: locate Playwright browser cache ───────────────────────────────────
console.log("\n🌐  Locating Playwright browser cache …");
let pwCache = null;
const candidates = [
  join(process.env.HOME ?? "", "Library", "Caches", "ms-playwright"), // macOS
  join(process.env.HOME ?? "", ".cache", "ms-playwright"), // Linux
  join(process.env.LOCALAPPDATA ?? "", "ms-playwright"), // Windows
  join(__dir, "node_modules", "playwright", ".local-browsers"),
];
for (const c of candidates) {
  if (existsSync(c) && readdirSync(c).length > 0) {
    pwCache = c;
    break;
  }
}

if (!pwCache) {
  console.error("  ⚠️  Could not locate Playwright browser cache.");
  console.error("     Run: npx playwright install chromium");
  process.exit(1);
}
console.log(`  Found: ${pwCache}`);

// ── Step 4: platform-specific packaging ──────────────────────────────────────

if (platform === "darwin") {
  // ── macOS: build a proper .app bundle ──────────────────────────────────────
  // A .app is just a folder with a specific structure. The key thing is that
  // the executable inside must have +x permissions — which normal zip loses.
  // The GitHub Actions workflow uses `ditto` to zip, which preserves permissions.

  const appDir = join(dist, "LookOut.app");
  const contentsDir = join(appDir, "Contents");
  const macOsDir = join(contentsDir, "MacOS");
  const resourcesDir = join(contentsDir, "Resources");
  mkdirSync(macOsDir, { recursive: true });
  mkdirSync(resourcesDir, { recursive: true });

  // Move the raw binary inside the bundle (rename to avoid clash with launcher)
  cpSync(join(dist, "LookOut"), join(macOsDir, "LookOut-bin"));

  // App icon: PNG → ICNS
  let iconEntry = "";
  if (existsSync(iconPngPath)) {
    try {
      const { default: p2i } = await import("png2icons");
      const pngData = readFileSync(iconPngPath);
      const icnsData = p2i.createICNS(pngData, p2i.BICUBIC2, 0);
      if (icnsData) {
        writeFileSync(join(resourcesDir, "AppIcon.icns"), icnsData);
        iconEntry = `<key>CFBundleIconFile</key><string>AppIcon</string>`;
        console.log("  ✅  AppIcon.icns written");
      }
    } catch (e) {
      console.warn(`  ⚠️  Icon conversion failed: ${e.message}`);
    }
  }

  // Info.plist — LSUIElement=true hides the app from the Dock while running
  writeFileSync(
    join(contentsDir, "Info.plist"),
    [
      `<?xml version="1.0" encoding="UTF-8"?>`,
      `<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">`,
      `<plist version="1.0"><dict>`,
      `<key>CFBundleExecutable</key><string>LookOut</string>`,
      `<key>CFBundleIdentifier</key><string>com.lookout.app</string>`,
      `<key>CFBundleName</key><string>LookOut</string>`,
      iconEntry,
      `<key>CFBundleVersion</key><string>1.0</string>`,
      `<key>CFBundlePackageType</key><string>APPL</string>`,
      `<key>LSUIElement</key><true/>`,
      `</dict></plist>`,
    ].join(""),
    "utf8",
  );

  // Copy Playwright browser into the bundle
  cpSync(pwCache, join(macOsDir, "pw-browsers"), { recursive: true });
  console.log("  ✅  Playwright browser copied into .app");

  // Launcher script inside the .app — opens a Terminal window so output is visible
  writeFileSync(
    join(macOsDir, "LookOut"),
    [
      "#!/bin/bash",
      'DIR="$(cd "$(dirname "$0")" && pwd)"',
      'export PLAYWRIGHT_BROWSERS_PATH="$DIR/pw-browsers"',
      // If not already running inside a terminal, re-launch in one
      'if [ -z "$TERM" ] && [ -z "$LAUNCHED_IN_TERMINAL" ]; then',
      "  export LAUNCHED_IN_TERMINAL=1",
      "  osascript \\",
      '    -e "tell application \\"Terminal\\"" \\',
      '    -e "  do script \\"LAUNCHED_IN_TERMINAL=1 PLAYWRIGHT_BROWSERS_PATH=\\\\\\"$DIR/pw-browsers\\\\\\" \\\\\\"$DIR/LookOut-bin\\\\\\" $*\\"" \\',
      '    -e "  activate" \\',
      '    -e "end tell"',
      "  exit 0",
      "fi",
      '"$DIR/LookOut-bin" "$@"',
    ].join("\n"),
    { mode: 0o755 },
  );

  // Clean up the loose binary (it's now inside the .app)
  execSync(`rm -f "${join(dist, "LookOut")}"`, { stdio: "inherit" });

  console.log("  ✅  LookOut.app built");
  console.log("\n⚠️   IMPORTANT — zip with ditto, not regular zip:");
  console.log(
    "      ditto -c -k --sequesterRsrc --keepParent dist/LookOut.app LookOut-macos.zip",
  );
  console.log(
    "      Regular zip strips execute permissions and the app won't open.\n",
  );
} else if (platform === "win32") {
  // ── Windows: flat folder with exe + pw-browsers/ + run.bat ────────────────
  cpSync(pwCache, join(dist, "pw-browsers"), { recursive: true });
  writeFileSync(
    join(dist, "run.bat"),
    '@echo off\r\nset DIR=%~dp0\r\nset PLAYWRIGHT_BROWSERS_PATH=%DIR%pw-browsers\r\n"%DIR%LookOut.exe" %*\r\n',
  );
  console.log("  ✅  pw-browsers copied, run.bat written");
} else {
  // ── Linux: flat folder with binary + pw-browsers/ + run.sh ────────────────
  cpSync(pwCache, join(dist, "pw-browsers"), { recursive: true });
  writeFileSync(
    join(dist, "run.sh"),
    [
      "#!/bin/bash",
      'DIR="$(cd "$(dirname "$0")" && pwd)"',
      'export PLAYWRIGHT_BROWSERS_PATH="$DIR/pw-browsers"',
      '"$DIR/LookOut" "$@"',
      "",
    ].join("\n"),
    { mode: 0o755 },
  );
  console.log("  ✅  pw-browsers copied, run.sh written");
}

// ── Done ──────────────────────────────────────────────────────────────────────
console.log("\n✅  Build complete! dist/ contents:");
for (const f of readdirSync(dist)) console.log("   ", f);
