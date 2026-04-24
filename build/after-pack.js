// electron-builder `afterPack` hook.
//
// Apple Silicon refuses to launch unsigned binaries — the kernel's
// AMFI (Apple Mobile File Integrity) enforces a mandatory code signature
// on every executable, even if it's just ad-hoc. When no Developer ID
// is configured, electron-builder ships the .app completely unsigned,
// which produces macOS's infamous "X is damaged and can't be opened"
// error on arm64 (Intel Macs don't have AMFI so they tolerate it).
//
// This hook ad-hoc-signs the packaged .app with codesign's self-identity
// (`-`). That's enough to pass AMFI. Users still need to clear the
// quarantine flag on first launch (right-click → Open, or
// `xattr -cr /Applications/Ketcher\ Desktop.app`) because we're not
// notarized — but at least the app can *run*.
//
// Skipped on non-mac builds and when a real Developer ID is present
// (in which case electron-builder's own signing pipeline takes over).

const { execFileSync } = require('child_process');
const path = require('path');

module.exports = async function afterPack(context) {
  if (context.electronPlatformName !== 'darwin') return;

  // If a real Developer ID is configured, electron-builder already
  // signed the app — don't overwrite that with an ad-hoc signature.
  if (process.env.CSC_LINK || process.env.CSC_NAME) return;

  const appPath = path.join(
    context.appOutDir,
    `${context.packager.appInfo.productFilename}.app`
  );

  console.log(`  • ad-hoc signing ${appPath}`);
  try {
    execFileSync(
      'codesign',
      ['--force', '--deep', '--sign', '-', appPath],
      { stdio: 'inherit' }
    );
  } catch (err) {
    console.warn('  • ad-hoc signing failed:', err.message);
    // Don't abort the build — the DMG still works on Intel, and M-series
    // users can run the codesign command manually as a workaround.
  }
};
