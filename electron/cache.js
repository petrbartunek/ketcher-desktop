// On-disk cache of exported structures, keyed by a short content hash.
//
// Why this exists:
//   The ketcher:// deep links we embed in .docx files used to carry the full
//   structure as a base64url-encoded query parameter, e.g.
//     ketcher://open?format=cdxml&data=<~4000 chars>
//   Word for Mac silently refuses to follow custom-protocol URLs that long
//   when clicked from inside a document — you get "An unexpected error has
//   occurred" with no further diagnostic. macOS's `open` command at the
//   shell has no such limit, which is why the Terminal test round-trips
//   fine but the in-document click fails.
//
// Fix: stash the structure on disk and put only a short reference in the URL:
//     ketcher://open?ref=<12-hex-hash>
//   The main process resolves the ref back to its content and feeds the
//   renderer the same { format, data } pair it already knows how to handle.
//
// Trade-off: an emailed .docx clicked on a DIFFERENT machine won't find the
// ref in the local cache. For that case the user sees a friendly error and
// can re-export on their side. We can add a cross-machine fallback later if
// it becomes a real need (would require embedding the payload inline for
// small molecules — which defeats the URL-length fix anyway).
//
// Files live in Electron's per-user app-data dir, one file per hash. Cache
// size is unbounded; a menu item under File lets the user nuke it.

'use strict';

const crypto = require('crypto');
const fs = require('fs/promises');
const path = require('path');
const { app } = require('electron');

// 12 hex chars = 48 bits ≈ 2.8 × 10^14 distinct refs, plenty for a personal
// cache even over years of use. Collisions would silently map one structure
// onto another, which would be confusing, but the risk is negligible unless
// we start tracking millions of unique structures per user.
const HASH_LEN = 12;

// The formats we round-trip through the cache, in preference order. CDXML
// is richer (stereochemistry, atom maps, etc.) than molfile; both are safe
// text. We don't cache binary CDX — CDXML is close enough and the UTF-8
// hash is stable across platforms.
const FORMATS = [
  { key: 'cdxml', ext: 'cdxml' },
  { key: 'mol',   ext: 'mol'   },
];

function dir() {
  return path.join(app.getPath('userData'), 'structures');
}

async function ensureDir() {
  await fs.mkdir(dir(), { recursive: true });
}

function shortHash(content) {
  return crypto.createHash('sha256').update(content, 'utf8').digest('hex').slice(0, HASH_LEN);
}

// Write the best-available representation to the cache and return a short
// ketcher:// URL that points back at it. Returns null if nothing to store.
async function put({ cdxml, molfile }) {
  if (!cdxml && !molfile) return null;

  const primary = cdxml
    ? { key: 'cdxml', ext: 'cdxml', content: cdxml }
    : { key: 'mol',   ext: 'mol',   content: molfile };

  await ensureDir();
  const hash = shortHash(primary.content);
  const file = path.join(dir(), `${hash}.${primary.ext}`);

  // Dedup: skip the write if we already have this content.
  try { await fs.access(file); }
  catch { await fs.writeFile(file, primary.content, 'utf8'); }

  return {
    url: `ketcher://open?ref=${hash}`,
    hash,
    format: primary.key,
  };
}

// Look up a hash and return { format, data } or null if missing.
async function resolveRef(hash) {
  if (!/^[0-9a-f]{1,40}$/i.test(hash)) return null;
  for (const { key, ext } of FORMATS) {
    const file = path.join(dir(), `${hash}.${ext}`);
    try {
      const data = await fs.readFile(file, 'utf8');
      return { format: key, data };
    } catch { /* try next format */ }
  }
  return null;
}

// Delete everything in the cache dir. Returns the count removed.
async function clear() {
  try {
    const files = await fs.readdir(dir());
    await Promise.all(files.map((f) => fs.unlink(path.join(dir(), f)).catch(() => {})));
    return files.length;
  } catch {
    return 0;
  }
}

async function size() {
  try { return (await fs.readdir(dir())).length; }
  catch { return 0; }
}

module.exports = { put, resolveRef, clear, size, dir };
