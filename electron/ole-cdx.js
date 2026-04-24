// OLE Compound Binary File writer for embedding CDX in Word.
//
// A ChemDraw OLE object inside a .docx lives at word/embeddings/oleObjectN.bin
// and is an OLE CFB (Compound Binary File, aka Microsoft structured storage)
// with three streams:
//
//   \x01Ole       — minimal header identifying this as an OLE object
//   \x01CompObj   — "CompObj" metadata: CLSID, user name, ProgID
//   CONTENTS      — raw binary CDX payload (the bytes of a .cdx file)
//
// For CDXML (no binary CDX available), Word still needs a CFB wrapper —
// we put the XML in CONTENTS and set the ProgID to ChemDraw.Document.
//
// ChemDraw's CLSID and ProgID vary by version. We target the most common
// modern one; users with older ChemDraw may need to adjust CLSID_CHEMDRAW
// below. The DOCX will still render the preview image regardless.
//
// References consulted mentally:
//   [MS-CFB] Compound File Binary File Format
//   [MS-OLEDS] Object Linking and Embedding (OLE) Data Structures
//
// This module has no external dependencies. Writing CFB by hand keeps the
// Electron bundle small and avoids version issues with cfb packages.

'use strict';

// ChemDraw 15+ CLSID. If your recipients run a different ChemDraw version,
// substitute their CLSID here. This value is used in both the CompObj stream
// and the storage's CLSID field.
const CLSID_CHEMDRAW = '4A5E2C77-9BFB-11D4-AE36-0050DA4C6752';
const PROGID_CHEMDRAW = 'ChemDraw.Document.15';
const USERNAME_CHEMDRAW = 'ChemDraw Document';

// ---- low-level helpers ----------------------------------------------------

function u16(n) {
  const b = Buffer.alloc(2);
  b.writeUInt16LE(n >>> 0, 0);
  return b;
}
function u32(n) {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n >>> 0, 0);
  return b;
}
function i32(n) {
  const b = Buffer.alloc(4);
  b.writeInt32LE(n | 0, 0);
  return b;
}

// UTF-16LE string with no NUL terminator.
function utf16le(s) {
  return Buffer.from(s, 'utf16le');
}

// Convert a hyphenated GUID to the 16-byte CLSID layout used on disk.
// GUID text:   XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
// On disk: [uint32 LE][uint16 LE][uint16 LE][8 bytes big-endian]
function clsidBytes(guid) {
  const hex = guid.replace(/-/g, '');
  const buf = Buffer.alloc(16);
  buf.writeUInt32LE(parseInt(hex.slice(0, 8), 16), 0);
  buf.writeUInt16LE(parseInt(hex.slice(8, 12), 16), 4);
  buf.writeUInt16LE(parseInt(hex.slice(12, 16), 16), 6);
  Buffer.from(hex.slice(16), 'hex').copy(buf, 8);
  return buf;
}

// ---- CompObj stream body --------------------------------------------------
// Layout:
//   u32  reserved1 = 0xFFFFFFFE
//   u32  version   = 0x0A03 (or similar)
//   u16  byte order = 0xFFFE
//   u16  format version = 0x0009
//   u32  OS type  = 0x00020000 (Windows NT)
//   u32  reserved2 = 0x00000000
//   CLSID (16 bytes)
//   LengthPrefixedAnsiString  user type (null terminated)
//   LengthPrefixedAnsiString  clipboard format (0 means none)
//   LengthPrefixedAnsiString  ProgID (ANSI)
//   u32 UnicodeMarker = 0x71B239F4
//   LengthPrefixedUnicodeString user type
//   LengthPrefixedUnicodeString clipboard format (0)
//   LengthPrefixedUnicodeString ProgID
function ansiLenPrefix(str) {
  // Includes trailing NUL.
  const body = Buffer.from(str + '\0', 'latin1');
  return Buffer.concat([u32(body.length), body]);
}
function unicodeLenPrefix(str) {
  if (!str) return u32(0);
  const body = Buffer.from(str + '\0', 'utf16le'); // character count incl. NUL
  const charCount = str.length + 1;
  return Buffer.concat([u32(charCount), body]);
}

function buildCompObjStream() {
  const header = Buffer.concat([
    u32(0xFFFFFFFE),       // reserved1
    u32(0x0000000A),       // version
    u16(0xFFFE),           // byte order mark
    u16(0x0009),           // format version
    u32(0x00000002),       // OS type (Windows)
    u32(0x00000000),       // reserved2
    clsidBytes(CLSID_CHEMDRAW),
  ]);

  const ansiBlock = Buffer.concat([
    ansiLenPrefix(USERNAME_CHEMDRAW),
    u32(0),                // clipboard format (none)
    ansiLenPrefix(PROGID_CHEMDRAW),
  ]);

  const unicodeMarker = u32(0x71B239F4);
  const unicodeBlock = Buffer.concat([
    unicodeMarker,
    unicodeLenPrefix(USERNAME_CHEMDRAW),
    u32(0),
    unicodeLenPrefix(PROGID_CHEMDRAW),
  ]);

  return Buffer.concat([header, ansiBlock, unicodeBlock]);
}

// ---- Ole stream -----------------------------------------------------------
// Minimal \x01Ole stream: 20 bytes header.
function buildOleStream() {
  return Buffer.concat([
    u32(0x02000001),  // Version
    u32(0x00000000),  // Flags
    u32(0xFFFFFFFF),  // LinkUpdateOption (none)
    u32(0x00000000),  // Reserved1
    u32(0x00000000),  // Reserved2 (moniker stream size = 0)
  ]);
}

// ---- CFB writer -----------------------------------------------------------
//
// We write a simple CFB with a single sector size of 512 bytes (Minor version 3).
// Small streams (< 4096 bytes) normally go into the mini-stream; for simplicity
// and robustness, we pad everything to full sectors and store in the regular FAT.

const SECTOR_SIZE = 512;
const DIFSECT = 0xFFFFFFFC;
const FATSECT = 0xFFFFFFFD;
const ENDOFCHAIN = 0xFFFFFFFE;
const FREESECT = 0xFFFFFFFF;

function padToSector(buf) {
  const rem = buf.length % SECTOR_SIZE;
  if (rem === 0) return buf;
  return Buffer.concat([buf, Buffer.alloc(SECTOR_SIZE - rem)]);
}

function sectorsNeeded(size) {
  return Math.ceil(size / SECTOR_SIZE);
}

/**
 * Build an OLE Compound File containing a ChemDraw object.
 *
 * @param {Object}   args
 * @param {Buffer}   args.contents  raw bytes for the CONTENTS stream (CDX binary or CDXML text)
 * @returns {Buffer} oleObject.bin
 */
function buildOleContainer({ contents }) {
  if (!Buffer.isBuffer(contents)) contents = Buffer.from(contents);

  const compObj = buildCompObjStream();
  const oleHdr = buildOleStream();

  // Allocate sectors for each stream. Sector 0 is the FAT (just one for
  // simplicity), sector 1 is the directory, then streams follow.
  const streams = [
    { name: '\x01CompObj', data: compObj },
    { name: '\x01Ole',      data: oleHdr  },
    { name: 'CONTENTS',     data: contents },
  ];

  // Assign contiguous sector ranges to each stream.
  let nextSector = 1;               // sector 0 = FAT
  const dirStartSector = nextSector++;
  streams.forEach((s) => {
    s.startSector = nextSector;
    s.sectorCount = Math.max(1, sectorsNeeded(s.data.length));
    nextSector += s.sectorCount;
  });
  const totalSectors = nextSector;

  // ---- Build FAT -------------------------------------------------------
  const fatEntries = new Uint32Array(SECTOR_SIZE / 4).fill(FREESECT);
  fatEntries[0] = FATSECT;          // FAT itself
  // Directory is a single sector → ENDOFCHAIN.
  fatEntries[dirStartSector] = ENDOFCHAIN;
  // Stream chains:
  streams.forEach((s) => {
    for (let i = 0; i < s.sectorCount; i++) {
      const sec = s.startSector + i;
      fatEntries[sec] = (i === s.sectorCount - 1) ? ENDOFCHAIN : sec + 1;
    }
  });
  const fatSector = Buffer.from(fatEntries.buffer);

  // ---- Build directory ------------------------------------------------
  // Four directory entries per sector of 512 bytes (each entry = 128 bytes).
  // We need: [0] Root, [1] CompObj, [2] Ole, [3] CONTENTS
  function dirEntry({ name, type, color, leftSid, rightSid, childSid, clsid, startSector, streamSize }) {
    const entry = Buffer.alloc(128);
    // Name: UTF-16LE, max 31 chars + NUL = 32 chars × 2 bytes = 64 bytes
    const nameBuf = utf16le(name);
    nameBuf.copy(entry, 0, 0, Math.min(nameBuf.length, 62));
    entry.writeUInt16LE(Math.min(nameBuf.length + 2, 64), 64); // length in bytes incl. NUL
    entry.writeUInt8(type, 66);        // 1=storage,2=stream,5=root
    entry.writeUInt8(color, 67);       // 0=red,1=black
    entry.writeInt32LE(leftSid, 68);
    entry.writeInt32LE(rightSid, 72);
    entry.writeInt32LE(childSid, 76);
    (clsid || Buffer.alloc(16)).copy(entry, 80);
    entry.writeUInt32LE(0, 96);        // state bits
    // CreationTime/ModifiedTime at 100,108 — leave as zero
    entry.writeInt32LE(startSector, 116);
    entry.writeUInt32LE(streamSize >>> 0, 120);
    entry.writeUInt32LE(0, 124);       // high 32 bits of stream size (0 for <4GB)
    return entry;
  }

  // Build a balanced tree; for 3 entries, child = middle, left/right are siblings.
  // We'll keep it simple: CompObj as root's child, with CONTENTS to its right and Ole to its left.
  // Directory sids: 0 root, 1 CompObj, 2 Ole, 3 CONTENTS
  const dir = Buffer.concat([
    dirEntry({
      name: 'Root Entry', type: 5, color: 1,
      leftSid: -1, rightSid: -1, childSid: 1,
      clsid: clsidBytes(CLSID_CHEMDRAW),
      startSector: -2, // no mini-stream
      streamSize: 0,
    }),
    dirEntry({
      name: streams[0].name, type: 2, color: 1,
      leftSid: 2, rightSid: 3, childSid: -1,
      startSector: streams[0].startSector,
      streamSize: streams[0].data.length,
    }),
    dirEntry({
      name: streams[1].name, type: 2, color: 1,
      leftSid: -1, rightSid: -1, childSid: -1,
      startSector: streams[1].startSector,
      streamSize: streams[1].data.length,
    }),
    dirEntry({
      name: streams[2].name, type: 2, color: 1,
      leftSid: -1, rightSid: -1, childSid: -1,
      startSector: streams[2].startSector,
      streamSize: streams[2].data.length,
    }),
  ]);
  // Pad directory to full sector.
  const dirSector = Buffer.concat([dir, Buffer.alloc(SECTOR_SIZE - dir.length)]);

  // ---- Build header ---------------------------------------------------
  const header = Buffer.alloc(SECTOR_SIZE);
  // Signature
  Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]).copy(header, 0);
  // CLSID (0)
  // Minor version 0x003E, Major version 0x0003 (512-byte sectors)
  header.writeUInt16LE(0x003E, 24);
  header.writeUInt16LE(0x0003, 26);
  header.writeUInt16LE(0xFFFE, 28);       // byte order
  header.writeUInt16LE(0x0009, 30);       // sector shift (2^9 = 512)
  header.writeUInt16LE(0x0006, 32);       // mini sector shift (2^6 = 64)
  // bytes 34..39 reserved
  header.writeUInt32LE(0, 40);            // number of directory sectors (0 for v3)
  header.writeUInt32LE(1, 44);            // number of FAT sectors
  header.writeInt32LE(dirStartSector, 48); // first directory sector
  header.writeUInt32LE(0, 52);            // transaction signature
  header.writeUInt32LE(4096, 56);         // mini-stream cutoff
  header.writeInt32LE(-2, 60);            // first mini-FAT sector (none)
  header.writeUInt32LE(0, 64);            // number of mini-FAT sectors
  header.writeInt32LE(-2, 68);            // first DIFAT sector (none)
  header.writeUInt32LE(0, 72);            // number of DIFAT sectors
  // DIFAT (109 entries × 4 bytes = 436 bytes starting at offset 76)
  for (let i = 0; i < 109; i++) {
    header.writeInt32LE(i === 0 ? 0 : -1, 76 + i * 4);
  }

  // ---- Assemble file --------------------------------------------------
  const parts = [header, fatSector, dirSector];
  streams.forEach((s) => parts.push(padToSector(s.data)));
  return Buffer.concat(parts);
}

/**
 * Build an OLE container holding binary CDX.
 * @param {Buffer} cdxBuffer
 * @returns {Buffer}
 */
function buildCdxOle(cdxBuffer) {
  return buildOleContainer({ contents: cdxBuffer });
}

/**
 * Build an OLE container holding CDXML (as UTF-8 text).
 * @param {string} cdxml
 * @returns {Buffer}
 */
function buildCdxmlOle(cdxml) {
  return buildOleContainer({ contents: Buffer.from(cdxml, 'utf8') });
}

module.exports = {
  buildCdxOle,
  buildCdxmlOle,
  CLSID_CHEMDRAW,
  PROGID_CHEMDRAW,
};
