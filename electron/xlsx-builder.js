// Excel catalog writer — appends one row per exported structure to an
// .xlsx file. Each row gets:
//
//   A: ID         (auto-incrementing integer)
//   B: Name       (optional — empty when the user hasn't supplied one)
//   C: Structure  (embedded PNG, scaled to fit while preserving aspect ratio)
//   D: SMILES     (clickable ketcher:// hyperlink to re-open in Ketcher Desktop)
//   E: Formula
//   F: InChI
//   G: InChIKey
//   H: PubChem    (clickable "Similarity" link into PubChem's structure search)
//
// The first row is a bold, frozen header. If the target file doesn't
// exist we create it; if it does, we load it and append to the first
// matching worksheet.
//
// Aspect-ratio handling:
//   Earlier versions forced every image into a fixed 200×150 px box,
//   which stretched wide molecules (indole, benzene rings, etc.) into
//   4:3 no matter what. Now we read the PNG's intrinsic dimensions
//   from its IHDR chunk, fit them inside a 220×150 max box preserving
//   ratio, and size the row height to match the image height. This
//   yields clean, undistorted renderings regardless of molecule shape.
//
// Schema migration:
//   Three historical layouts exist:
//     v1: 5 cols — Structure, SMILES, Formula, InChI, InChIKey
//     v2: 7 cols — ID, Name, Structure, SMILES, Formula, InChI, InChIKey
//     v3: 8 cols — v2 + trailing PubChem column (current)
//   On first append we detect which layout the file uses and upgrade to
//   v3 in place. v1→v3 requires a full rebuild (new sheet with images
//   shifted two columns right, sequential IDs assigned). v2→v3 only
//   needs the PubChem header cell plus per-row hyperlink. Either way,
//   existing data is preserved.
//
// Why exceljs instead of SheetJS / our-own-XML approach like docx-builder:
//   - Embedded images in .xlsx live in a separate ZIP part (xl/media/...)
//     plus drawing XML plus anchor XML, and they get re-keyed every time
//     we add one. That machinery is a lot of code we'd be re-implementing.
//   - exceljs round-trips the whole file correctly including images, so
//     opening the same catalog ten times and appending rows doesn't
//     break anything or lose earlier images.

'use strict';

const ExcelJS = require('exceljs');
const fs = require('fs/promises');

const SHEET_NAME = 'Structures';
const HEADERS = ['ID', 'Name', 'Structure', 'SMILES', 'Formula', 'InChI', 'InChIKey', 'PubChem'];

// Column widths in Excel's "default character count" units. Rough rule:
// one unit ≈ 7 px with default font. These values produce an ID column
// wide enough for 4-5 digits, a Name column sized for typical compound
// names, a Structure column wide enough for a 220 px image plus margin,
// SMILES / InChI columns wide enough to show most strings, and a compact
// PubChem column that just holds a short "Similarity" link.
const COL_WIDTHS = [8, 24, 32, 50, 18, 60, 32, 14];

// Zero-based column index of the Structure column (where images anchor).
// ID(0), Name(1), Structure(2) → idx 2.
const STRUCT_COL_IDX = 2;

// Maximum image display box in pixels. Images are scaled to fit entirely
// inside this box while preserving their intrinsic aspect ratio.
const IMG_MAX_W = 220;
const IMG_MAX_H = 150;

// Minimum row height so very thin molecules (horizontal chains) still
// leave breathing room. In Excel points (1 pt = 1/72 inch, 1 px at 96
// DPI ≈ 0.75 pt).
const MIN_ROW_HEIGHT_PT = 60;

// Convert pixels at 96 DPI to Excel points.
function pxToPt(px) { return px * 0.75; }

// Parse PNG IHDR width / height out of a base64-encoded PNG string.
// PNG layout: 8-byte signature, then a sequence of chunks. The IHDR
// chunk always comes first; width is bytes 16-19 (big-endian), height
// is 20-23.
function pngDimensions(base64) {
  if (!base64) return null;
  try {
    const buf = Buffer.from(base64, 'base64');
    if (buf.length < 24) return null;
    // Verify PNG signature \x89PNG\r\n\x1a\n.
    if (buf[0] !== 0x89 || buf[1] !== 0x50 || buf[2] !== 0x4e || buf[3] !== 0x47) {
      return null;
    }
    const w = buf.readUInt32BE(16);
    const h = buf.readUInt32BE(20);
    if (!w || !h) return null;
    return { w, h };
  } catch {
    return null;
  }
}

// Same as pngDimensions but for a raw Node Buffer (no base64 decoding
// needed). Used during migration, where exceljs hands us image buffers
// directly.
function pngDimensionsFromBuffer(buf) {
  if (!buf || buf.length < 24) return null;
  try {
    if (buf[0] !== 0x89 || buf[1] !== 0x50 || buf[2] !== 0x4e || buf[3] !== 0x47) {
      return null;
    }
    return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  } catch {
    return null;
  }
}

// Fit (w, h) inside (maxW, maxH) preserving aspect ratio. Returns
// whole pixels, minimum 1×1. Never upscales past the native size —
// Ketcher already rasterizes at ~3× display, so the native PNG is
// already plenty sharp; upscaling further just blurs the picture.
function fitWithin(w, h, maxW, maxH) {
  if (!w || !h) return { w: maxW, h: maxH };
  const scale = Math.min(maxW / w, maxH / h, 1);
  return {
    w: Math.max(1, Math.round(w * scale)),
    h: Math.max(1, Math.round(h * scale)),
  };
}

// Pick the worksheet to append to. Prefer our named sheet, fall back to
// the first worksheet in the file (handles catalogs the user renamed).
function pickOrCreateSheet(workbook) {
  let sheet = workbook.getWorksheet(SHEET_NAME);
  if (sheet) return { sheet, fresh: false };
  if (workbook.worksheets.length > 0) {
    return { sheet: workbook.worksheets[0], fresh: false };
  }
  sheet = workbook.addWorksheet(SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  return { sheet, fresh: true };
}

function applyHeaderStyling(sheet) {
  const header = sheet.getRow(1);
  header.font = { bold: true, size: 11 };
  header.alignment = { vertical: 'middle', horizontal: 'center' };
  header.height = 22;
  header.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFEFEFEF' },
  };
  COL_WIDTHS.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
}

function applyDataRowAlignment(row) {
  row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' }; // ID
  row.getCell(2).alignment = { wrapText: true, vertical: 'top' };          // Name
  row.getCell(3).alignment = { vertical: 'middle', horizontal: 'center' }; // Structure
  row.getCell(4).alignment = { wrapText: true, vertical: 'top' };          // SMILES
  row.getCell(5).alignment = { vertical: 'top' };                          // Formula
  row.getCell(6).alignment = { wrapText: true, vertical: 'top' };          // InChI
  row.getCell(7).alignment = { vertical: 'top' };                          // InChIKey
  row.getCell(8).alignment = { vertical: 'middle', horizontal: 'center' }; // PubChem
}

// Build a PubChem similarity-search URL keyed on the row's SMILES.
// This is the URL PubChem's own web UI uses for in-page hash routing,
// so it keeps working even when PubChem tweaks their front end.
function pubchemSimilarityUrl(smiles) {
  if (!smiles) return null;
  const s = String(smiles).trim();
  if (!s) return null;
  return `https://pubchem.ncbi.nlm.nih.gov/#query=${encodeURIComponent(s)}` +
         '&input_type=smiles&tab=similarity';
}

// Write the PubChem cell for a row. Leaves it blank when SMILES is empty.
function setPubchemLinkCell(cell, smiles) {
  const url = pubchemSimilarityUrl(smiles);
  if (!url) {
    cell.value = '';
    return;
  }
  cell.value = {
    text: 'Similarity',
    hyperlink: url,
    tooltip: 'Search PubChem for structures similar to this SMILES',
  };
  cell.font = LINK_FONT;
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
}

function ensureHeader(sheet) {
  // A sheet with 0 rows hasn't been written to at all. A sheet where row 1
  // doesn't start with our headers is assumed to be someone else's layout
  // and we don't stomp on it — we just append below whatever's there.
  if (sheet.rowCount === 0) {
    sheet.addRow(HEADERS);
    applyHeaderStyling(sheet);
  }
}

// Detect which header schema (if any) the sheet currently has.
// Returns:
//   'current' — 8-col layout with PubChem (today's schema)
//   'v2'      — 7-col layout (ID/Name/Structure/...InChIKey, no PubChem)
//   'v1'      — 5-col layout (Structure/SMILES/Formula/InChI/InChIKey)
//   null      — empty sheet or a foreign layout we leave alone
function detectSchema(sheet) {
  if (!sheet || sheet.rowCount === 0) return null;
  const header = sheet.getRow(1);
  const a = String(header.getCell(1).value || '').trim();
  const b = String(header.getCell(2).value || '').trim();
  const c = String(header.getCell(3).value || '').trim();
  const h = String(header.getCell(8).value || '').trim();
  if (a === 'ID' && b === 'Name' && c === 'Structure' && h === 'PubChem') return 'current';
  if (a === 'ID' && b === 'Name' && c === 'Structure') return 'v2';
  if (a === 'Structure' && b === 'SMILES') return 'v1';
  return null;
}

// Upgrade a v2 (ID/Name/Structure/.../InChIKey) sheet to current by
// adding the PubChem header + per-row similarity links. In-place —
// no need to rebuild the sheet since no images are moving.
function migrateV2ToCurrent(sheet) {
  const header = sheet.getRow(1);
  header.getCell(8).value = 'PubChem';
  applyHeaderStyling(sheet);
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const smiles = String(row.getCell(4).value && row.getCell(4).value.text
      ? row.getCell(4).value.text               // SMILES cell may be a hyperlink object
      : (row.getCell(4).value || '')).trim();
    setPubchemLinkCell(row.getCell(8), smiles);
  }
}

// Compute the next auto-increment ID by scanning column A of existing
// data rows. Skips the header and any non-integer values. Returns 1 if
// nothing valid is found.
function nextId(sheet) {
  let max = 0;
  for (let r = 2; r <= sheet.rowCount; r++) {
    const v = sheet.getRow(r).getCell(1).value;
    const n = typeof v === 'number' ? v : parseInt(String(v), 10);
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max + 1;
}

// Rebuild a legacy 5-column workbook under the new 7-column schema.
// Reads existing rows + images, creates a fresh sheet with the new
// headers, re-inserts data shifted 2 columns right, assigns sequential
// IDs, and re-anchors images to the new Structure column. Mutates the
// workbook in place.
function migrateOldCatalog(workbook) {
  const oldSheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];
  if (!oldSheet) return;

  // Collect per-row image anchors.
  // exceljs anchors images by fractional row/col with a zero-based
  // "native" coordinate; sheet rows are 1-based, so we +1.
  const imagesByRow = new Map();
  const images = typeof oldSheet.getImages === 'function' ? oldSheet.getImages() : [];
  for (const img of images) {
    if (!img || !img.range || !img.range.tl) continue;
    const nativeRow = img.range.tl.nativeRow != null
      ? img.range.tl.nativeRow
      : (img.range.tl.row || 0);
    const rowNum = Math.round(nativeRow) + 1;
    // If multiple images share a row (shouldn't happen in our catalogs),
    // the first one wins.
    if (!imagesByRow.has(rowNum)) imagesByRow.set(rowNum, img);
  }

  const migrated = [];
  const lastRow = oldSheet.rowCount;
  for (let r = 2; r <= lastRow; r++) {
    const row = oldSheet.getRow(r);
    // Old schema: A=Structure(placeholder), B=SMILES, C=Formula, D=InChI, E=InChIKey.
    const smiles   = row.getCell(2).value;
    const formula  = row.getCell(3).value;
    const inchi    = row.getCell(4).value;
    const inchiKey = row.getCell(5).value;

    let imageBuffer = null;
    let imageExt = 'png';
    const img = imagesByRow.get(r);
    if (img && img.imageId != null) {
      try {
        const media = workbook.getImage(img.imageId);
        if (media && media.buffer) {
          imageBuffer = media.buffer;
          imageExt = media.extension || 'png';
        }
      } catch {
        // If we can't extract the buffer, the migrated row keeps its
        // text data but loses the image; acceptable degradation.
      }
    }

    migrated.push({
      smiles:   smiles   == null ? '' : String(smiles),
      formula:  formula  == null ? '' : String(formula),
      inchi:    inchi    == null ? '' : String(inchi),
      inchiKey: inchiKey == null ? '' : String(inchiKey),
      imageBuffer,
      imageExt,
    });
  }

  // Strip the old sheet and start fresh with the same name so the
  // user's references to it (filters, named ranges) keep working.
  const oldId = oldSheet.id;
  const oldName = oldSheet.name;
  workbook.removeWorksheet(oldId);
  const newSheet = workbook.addWorksheet(oldName || SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  newSheet.addRow(HEADERS);
  applyHeaderStyling(newSheet);

  migrated.forEach((rec, i) => {
    const id = i + 1;
    const dataRow = newSheet.addRow([
      id,
      '',            // Name (empty for legacy rows — user can fill in)
      '',            // Structure placeholder; image anchored below
      rec.smiles,
      rec.formula,
      rec.inchi,
      rec.inchiKey,
      '',            // PubChem placeholder; filled by setPubchemLinkCell below
    ]);

    // Figure out display size from the native buffer dimensions when
    // possible so migrated rows also use correct aspect ratios.
    let displayW = IMG_MAX_W;
    let displayH = IMG_MAX_H;
    if (rec.imageBuffer) {
      const dims = pngDimensionsFromBuffer(rec.imageBuffer);
      if (dims) {
        const fit = fitWithin(dims.w, dims.h, IMG_MAX_W, IMG_MAX_H);
        displayW = fit.w;
        displayH = fit.h;
      }
    }
    dataRow.height = Math.max(MIN_ROW_HEIGHT_PT, pxToPt(displayH) + 6);
    applyDataRowAlignment(dataRow);
    setPubchemLinkCell(dataRow.getCell(8), rec.smiles);

    if (rec.imageBuffer) {
      const newImageId = workbook.addImage({
        buffer: rec.imageBuffer,
        extension: rec.imageExt,
      });
      newSheet.addImage(newImageId, {
        tl: { col: STRUCT_COL_IDX, row: dataRow.number - 1 },
        ext: { width: displayW, height: displayH },
        editAs: 'oneCell',
      });
    }
  });

  return newSheet;
}

// Style applied to SMILES cells that carry a Ketcher deep link. Matches
// Excel's default hyperlink look so it's obvious the cell is clickable.
const LINK_FONT = { color: { argb: 'FF0563C1' }, underline: true };

// Append one structure to the catalog at `filePath`. Creates the file
// if it doesn't exist yet.
//
// row: {
//   pngBase64, smiles, formula, inchi, inchiKey, name?,
//   deepLink?,     // optional ketcher://open?ref=<hash> URL
// }
// Returns: { rowNumber, created, migrated }
async function appendToXlsxCatalog(filePath, row) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Ketcher Desktop';
  workbook.modified = new Date();

  let created = false;
  try {
    await fs.access(filePath);
    await workbook.xlsx.readFile(filePath);
  } catch {
    created = true;
    workbook.created = new Date();
  }

  // If this is a legacy catalog, migrate before appending so we don't
  // end up with a half-old / half-new file.
  let migrated = false;
  if (!created) {
    const existingSheet = workbook.getWorksheet(SHEET_NAME) || workbook.worksheets[0];
    const schema = existingSheet ? detectSchema(existingSheet) : null;
    if (schema === 'v1') {
      // 5-col → 8-col (full rebuild; images shift right, IDs assigned,
      // PubChem column populated).
      migrateOldCatalog(workbook);
      migrated = true;
    } else if (schema === 'v2') {
      // 7-col → 8-col (in-place PubChem column).
      migrateV2ToCurrent(existingSheet);
      migrated = true;
    }
  }

  const { sheet } = pickOrCreateSheet(workbook);
  ensureHeader(sheet);

  // Figure out image dimensions + row height before inserting, so the
  // row height matches the rasterized PNG exactly.
  let displayW = IMG_MAX_W;
  let displayH = IMG_MAX_H;
  if (row.pngBase64) {
    const dims = pngDimensions(row.pngBase64);
    if (dims) {
      const fit = fitWithin(dims.w, dims.h, IMG_MAX_W, IMG_MAX_H);
      displayW = fit.w;
      displayH = fit.h;
    }
  }

  const id = nextId(sheet);
  const smilesText = row.smiles || '';
  const dataRow = sheet.addRow([
    id,
    row.name || '',
    '',                   // Structure placeholder; image anchored below
    smilesText,           // may be replaced below with a hyperlink cell
    row.formula  || '',
    row.inchi    || '',
    row.inchiKey || '',
    '',                   // PubChem placeholder; filled by setPubchemLinkCell
  ]);
  // +6 pt padding so the image doesn't sit flush against the row border.
  dataRow.height = Math.max(MIN_ROW_HEIGHT_PT, pxToPt(displayH) + 6);
  applyDataRowAlignment(dataRow);

  // If the caller gave us a deep link (ketcher://open?ref=...), turn
  // the SMILES cell into a clickable hyperlink. On the same machine
  // Ketcher Desktop will re-open the structure; on a different machine
  // the `?ref=` resolver in main.js falls back to a friendly dialog.
  if (row.deepLink && smilesText) {
    const smilesCell = dataRow.getCell(4);
    smilesCell.value = {
      text: smilesText,
      hyperlink: row.deepLink,
      tooltip: 'Open this structure in Ketcher Desktop',
    };
    smilesCell.font = LINK_FONT;
  }

  // PubChem similarity link — does nothing when SMILES is empty.
  setPubchemLinkCell(dataRow.getCell(8), smilesText);

  if (row.pngBase64) {
    const imageId = workbook.addImage({
      base64: row.pngBase64,
      extension: 'png',
    });
    // tl is zero-based; dataRow.number is one-based, so subtract one.
    // editAs:'oneCell' keeps the image glued to the row if someone
    // resizes columns or inserts rows above it.
    sheet.addImage(imageId, {
      tl: { col: STRUCT_COL_IDX, row: dataRow.number - 1 },
      ext: { width: displayW, height: displayH },
      editAs: 'oneCell',
    });
  }

  try {
    await workbook.xlsx.writeFile(filePath);
  } catch (err) {
    // Windows and Excel both love to hold exclusive write locks on open
    // .xlsx files. Re-throw with a clearer message so main.js surfaces
    // something useful in the error dialog.
    if (err && /EBUSY|EACCES|locked/i.test(String(err.message))) {
      throw new Error(
        `Could not write to ${filePath}. Close the file in Excel (or any ` +
        `other app that has it open) and try again.`
      );
    }
    throw err;
  }

  return { rowNumber: dataRow.number, created, migrated };
}

module.exports = { appendToXlsxCatalog };
