#!/usr/bin/env node
// Dev-mode helper: teach macOS Launch Services that ketcher:// should route
// to the local dev Electron.app.
//
// Why this exists:
//   Electron's `app.setAsDefaultProtocolClient('ketcher')` asks Launch
//   Services at runtime to make this process the handler, but Launch
//   Services only persists the association if the binary's Info.plist
//   declares the scheme via CFBundleURLTypes. In dev mode, the binary is
//   `node_modules/electron/dist/Electron.app/Contents/MacOS/Electron`, whose
//   stock Info.plist has no CFBundleURLTypes for `ketcher`, so the
//   registration silently fails and Word gets "no app can open this URL"
//   (which Mac Word surfaces as the generic "An unexpected error has
//   occurred").
//
// Usage:
//   node tools/register-scheme.js             # patch + lsregister
//   node tools/register-scheme.js --revert    # remove our Info.plist entry
//
// This is mac-only. It's a no-op on Linux / Windows, where registration
// works through electron-builder + .desktop / registry entries once packaged.

'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const os = require('os');

if (process.platform !== 'darwin') {
  console.log('[register-scheme] Not macOS — nothing to do.');
  process.exit(0);
}

const SCHEME = 'ketcher';
const REPO_ROOT = path.resolve(__dirname, '..');
const ELECTRON_APP = path.join(REPO_ROOT, 'node_modules', 'electron', 'dist', 'Electron.app');
const PLIST_PATH = path.join(ELECTRON_APP, 'Contents', 'Info.plist');
const LSREG = '/System/Library/Frameworks/CoreServices.framework/Versions/A/Frameworks/LaunchServices.framework/Versions/A/Support/lsregister';

function plutil(args) {
  return execFileSync('/usr/bin/plutil', args, { encoding: 'utf8' });
}

function keyExists(key) {
  try {
    plutil(['-extract', key, 'xml1', '-o', '-', PLIST_PATH]);
    return true;
  } catch { return false; }
}

function main() {
  if (!fs.existsSync(PLIST_PATH)) {
    console.error(`[register-scheme] Could not find ${PLIST_PATH}`);
    console.error('   Did you run `npm install` yet?');
    process.exit(1);
  }

  if (process.argv.includes('--revert')) {
    if (keyExists('CFBundleURLTypes')) {
      plutil(['-remove', 'CFBundleURLTypes', PLIST_PATH]);
      console.log('[register-scheme] Removed CFBundleURLTypes from Electron dev Info.plist.');
    } else {
      console.log('[register-scheme] No CFBundleURLTypes to remove.');
    }
    execFileSync(LSREG, ['-f', ELECTRON_APP], { stdio: 'inherit' });
    return;
  }

  // Add (or replace) the CFBundleURLTypes array with a single entry for ketcher://
  if (keyExists('CFBundleURLTypes')) {
    plutil(['-remove', 'CFBundleURLTypes', PLIST_PATH]);
  }
  plutil(['-insert', 'CFBundleURLTypes', '-json',
    JSON.stringify([{
      CFBundleURLName: 'Ketcher Structure',
      CFBundleURLSchemes: [SCHEME],
    }]),
    PLIST_PATH,
  ]);

  console.log(`[register-scheme] Added CFBundleURLTypes=${SCHEME}:// to ${PLIST_PATH}`);

  // Force Launch Services to re-read the bundle. Without -f it often caches.
  execFileSync(LSREG, ['-f', ELECTRON_APP], { stdio: 'inherit' });
  console.log('[register-scheme] Launch Services refreshed.');

  // Quick verification.
  try {
    const out = execFileSync(LSREG, ['-dump'], { encoding: 'utf8' });
    const hit = out.split('\n').find((l) => l.includes(`${SCHEME}:`));
    if (hit) {
      console.log(`[register-scheme] OK. Launch Services now lists: ${hit.trim()}`);
    } else {
      console.warn('[register-scheme] lsregister did not list the scheme — you may need to log out/in.');
    }
  } catch (e) {
    console.warn('[register-scheme] Could not verify via lsregister -dump:', e.message);
  }

  console.log('\nNow quit Ketcher Desktop (Cmd-Q) and rerun `npm run dev`.');
  console.log('Clicking a ketcher:// link in Word should route here.');
}

main();
