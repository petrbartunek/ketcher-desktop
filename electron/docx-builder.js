// Build (or modify) a .docx that embeds a ChemDraw OLE object.
//
// Two public functions:
//
//   buildDocxWithCdx(bundle) -> Buffer
//      Creates a brand-new minimal Word document containing a single
//      embedded structure (with PNG preview) and an optional caption.
//
//   insertCdxIntoDocx(input, bundle) -> Buffer
//      Takes an existing .docx and replaces the first occurrence of
//      "{{CDX:1}}" in document.xml with an embedded OLE object.
//
// The OLE object is a ChemDraw.Document.15 OLEObject referencing
// word/embeddings/oleObjectN.bin — which is the CFB produced by ole-cdx.js.
// The preview image is stored at word/media/imageN.png (rasterized PNG
// supplied by the caller — Ketcher's generateImage('png') output).
//
// Depends on: jszip (for ZIP assembly).

'use strict';

const JSZip = require('jszip');
const { buildCdxOle, buildCdxmlOle, PROGID_CHEMDRAW } = require('./ole-cdx');

// Preferred order: binary CDX (if Ketcher produced it), fall back to CDXML.
function buildOleBinary({ cdxBase64, cdxml }) {
  if (cdxBase64) {
    return { bin: buildCdxOle(Buffer.from(cdxBase64, 'base64')), kind: 'cdx' };
  }
  if (cdxml) {
    return { bin: buildCdxmlOle(cdxml), kind: 'cdxml' };
  }
  throw new Error('Neither CDX nor CDXML was provided');
}

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Read width/height from a PNG buffer. PNG has a fixed header: 8-byte
// signature, then a 4-byte chunk length, then "IHDR", then 4 bytes width,
// 4 bytes height — all big-endian.
function pngDimensions(buf) {
  if (!buf || buf.length < 24) return null;
  if (buf[0] !== 0x89 || buf[1] !== 0x50 || buf[2] !== 0x4E || buf[3] !== 0x47) return null;
  return {
    width:  buf.readUInt32BE(16),
    height: buf.readUInt32BE(20),
  };
}

// Compute display size in points, capping the longer dimension at maxPt.
// Default cap: 3.5 inches = 252 pt.
function displaySize(pngBuf, { maxPt = 252 } = {}) {
  const d = pngDimensions(pngBuf) || { width: 600, height: 450 };
  const ratio = d.width / d.height;
  let w, h;
  if (d.width >= d.height) {
    w = maxPt;
    h = Math.round(maxPt / ratio);
  } else {
    h = maxPt;
    w = Math.round(maxPt * ratio);
  }
  return { widthPt: w, heightPt: h };
}

// VML shape + OLEObject snippet that Word uses to render an embedded object
// with a bitmap fallback. Shape dimensions follow the PNG's own aspect
// ratio so Word doesn't stretch the preview.
function oleObjectXml({ imageRid, oleRid, shapeId, widthPt, heightPt }) {
  return `
    <w:object w:dxaOrig="${widthPt * 20}" w:dyaOrig="${heightPt * 20}">
      <v:shape id="${xmlEscape(shapeId)}" type="#_x0000_t75"
               style="width:${widthPt}pt;height:${heightPt}pt" o:ole="">
        <v:imagedata r:id="${xmlEscape(imageRid)}" o:title=""/>
      </v:shape>
      <o:OLEObject Type="Embed" ProgID="${xmlEscape(PROGID_CHEMDRAW)}"
                   ShapeID="${xmlEscape(shapeId)}" DrawAspect="Content"
                   ObjectID="_${Math.floor(Math.random() * 1e9)}" r:id="${xmlEscape(oleRid)}"/>
    </w:object>`;
}

// ---------------------------------------------------------------------------
// Placeholder replacement that tolerates Word's "split across runs" habit.
//
// When you type `{{CDX:1}}` into Word, Word sometimes splits the characters
// into multiple <w:r>…<w:t>…</w:t>…</w:r> elements (auto-correct, spell-check,
// inline edits, different formatting spans, etc.). A plain string search for
// `{{CDX:1}}` then misses it even though the user can see it clearly.
//
// This function walks through the document's runs, concatenates their visible
// text, and replaces the first span of runs whose joint text contains the
// placeholder. The `replacement` is injected as-is (it can be plain text, a
// run, or a sequence of runs).
//
// Returns the patched XML, or null if the placeholder really isn't anywhere.
// ---------------------------------------------------------------------------
function replacePlaceholderAcrossRuns(docXml, placeholder, replacement) {
  const decodeXml = (s) => s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');

  const runRegex = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
  const runs = [];
  let m;
  while ((m = runRegex.exec(docXml)) !== null) {
    const runXml = m[0];
    let text = '';
    const tRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let tm;
    while ((tm = tRegex.exec(runXml)) !== null) text += decodeXml(tm[1]);
    runs.push({ start: m.index, end: m.index + runXml.length, text });
  }

  for (let i = 0; i < runs.length; i++) {
    let combined = '';
    for (let j = i; j < runs.length; j++) {
      combined += runs[j].text;
      const idx = combined.indexOf(placeholder);
      if (idx !== -1) {
        const before = combined.slice(0, idx);
        const after  = combined.slice(idx + placeholder.length);
        const startPos = runs[i].start;
        const endPos   = runs[j].end;
        const beforeRun = before
          ? `<w:r><w:t xml:space="preserve">${xmlEscape(before)}</w:t></w:r>`
          : '';
        const afterRun = after
          ? `<w:r><w:t xml:space="preserve">${xmlEscape(after)}</w:t></w:r>`
          : '';
        return docXml.slice(0, startPos) + beforeRun + replacement + afterRun + docXml.slice(endPos);
      }
    }
  }
  return null;
}

// base64url: URL-safe base64 without padding.
function base64url(input) {
  const buf = Buffer.isBuffer(input) ? input : Buffer.from(input, 'utf8');
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Fallback deep-link builder — embeds the full payload in the URL.
// Used only when the caller didn't pre-compute a short-ref URL via the
// cache module. Word for Mac silently rejects URLs longer than ~2000
// chars when clicked from inside a document, so anything beyond a tiny
// molecule fails to round-trip via this path. The preferred flow is for
// main.js to call cache.put() and pass the resulting short URL in as
// `deepLink`.
function buildInlineKetcherDeepLink({ cdxml, molfile }) {
  const MAX = 50_000;
  if (cdxml && cdxml.length * 1.4 < MAX) {
    return `ketcher://open?format=cdxml&data=${base64url(cdxml)}`;
  }
  if (molfile && molfile.length * 1.4 < MAX) {
    return `ketcher://open?format=mol&data=${base64url(molfile)}`;
  }
  return null;
}

// A small blue-underlined text hyperlink. Rendered as a separate paragraph
// (or inline run) below the structure — we intentionally do NOT wrap the
// <w:object> itself in a <w:hyperlink> because Word for Mac throws "An
// unexpected error has occurred" when double-click OLE activation collides
// with hyperlink navigation on the same shape.
//
// `italic` makes the text italic (matches the SMILES caption style).
function hyperlinkRun({ linkRid, text = 'Open in Ketcher Desktop', italic = false }) {
  if (!linkRid) return '';
  const italicTag = italic ? '<w:i/>' : '';
  return `<w:hyperlink r:id="${xmlEscape(linkRid)}" w:history="1">
    <w:r>
      <w:rPr>
        ${italicTag}
        <w:color w:val="0563C1"/>
        <w:u w:val="single"/>
      </w:rPr>
      <w:t xml:space="preserve">${xmlEscape(text)}</w:t>
    </w:r>
  </w:hyperlink>`;
}

// Minimal document.xml for the "new doc" flow.
// Layout: structure paragraph (OLE, no hyperlink wrap), plus a caption
// paragraph. If linkRid is supplied, the caption text itself becomes the
// clickable hyperlink back to Ketcher Desktop. If no caption is supplied
// but a link is available, we fall back to "Open in Ketcher Desktop".
function minimalDocumentXml({ imageRid, oleRid, linkRid, caption, widthPt, heightPt }) {
  const obj = oleObjectXml({ imageRid, oleRid, shapeId: '_x0000_s1026', widthPt, heightPt });

  let captionPara = '';
  if (caption && linkRid) {
    captionPara = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
      ${hyperlinkRun({ linkRid, text: caption, italic: true })}
    </w:p>`;
  } else if (caption) {
    captionPara = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${xmlEscape(caption)}</w:t></w:r>
    </w:p>`;
  } else if (linkRid) {
    captionPara = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
      ${hyperlinkRun({ linkRid })}
    </w:p>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>${obj}</w:r>
    </w:p>
    ${captionPara}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function minimalContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="bin"  ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"
            ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function docRelsXml({ imageRid, oleRid, imageTarget, oleTarget, linkRid, linkTarget }) {
  const linkRel = linkRid && linkTarget
    ? `<Relationship Id="${linkRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(linkTarget)}" TargetMode="External"/>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${imageTarget}"/>
  <Relationship Id="${oleRid}"   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="${oleTarget}"/>
  ${linkRel}
</Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>`;
}

function coreXml() {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Ketcher Structure</dc:title>
  <dc:creator>Ketcher Desktop</dc:creator>
  <cp:lastModifiedBy>Ketcher Desktop</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function appXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Ketcher Desktop</Application>
</Properties>`;
}

// --- PNG fallback ----------------------------------------------------------
// If the renderer couldn't produce a PNG, we still need SOME image for Word's
// OLE fallback. Ship a tiny 1×1 transparent PNG in that case so the file opens.
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

function resolvePreviewPng(pngBase64) {
  return Buffer.from(pngBase64 || TINY_PNG_BASE64, 'base64');
}

// --- Public: build a new .docx --------------------------------------------

async function buildDocxWithCdx({ cdxBase64, cdxml, molfile, pngBase64, caption, deepLink }) {
  const { bin: oleBin } = buildOleBinary({ cdxBase64, cdxml });
  const preview = resolvePreviewPng(pngBase64);
  const { widthPt, heightPt } = displaySize(preview);

  // Caller may supply a ready-made short-ref URL (preferred — see cache.js).
  // If not, fall back to stuffing the whole structure into the URL, which
  // is fine only for very small molecules because Word rejects long URLs.
  if (deepLink == null) deepLink = buildInlineKetcherDeepLink({ cdxml, molfile });
  const imageRid = 'rId100';
  const oleRid   = 'rId101';
  const linkRid  = deepLink ? 'rId102' : null;

  const zip = new JSZip();

  zip.file('[Content_Types].xml', minimalContentTypesXml());
  zip.file('_rels/.rels', rootRelsXml());
  zip.file('docProps/core.xml', coreXml());
  zip.file('docProps/app.xml', appXml());

  zip.file('word/styles.xml', stylesXml());

  zip.file('word/_rels/document.xml.rels', docRelsXml({
    imageRid, oleRid,
    imageTarget: 'media/image1.png',
    oleTarget:   'embeddings/oleObject1.bin',
    linkRid, linkTarget: deepLink,
  }));

  zip.file('word/document.xml', minimalDocumentXml({
    imageRid, oleRid, linkRid, caption, widthPt, heightPt,
  }));
  zip.file('word/media/image1.png', preview);
  zip.file('word/embeddings/oleObject1.bin', oleBin);

  return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// --- Public: insert into an existing .docx --------------------------------

async function insertCdxIntoDocx(inputBuf, { placeholder, cdxBase64, cdxml, molfile, pngBase64, smiles, deepLink }) {
  const { bin: oleBin } = buildOleBinary({ cdxBase64, cdxml });
  const preview = resolvePreviewPng(pngBase64);
  const { widthPt, heightPt } = displaySize(preview);
  // Same pattern as buildDocxWithCdx: prefer caller-supplied short-ref URL;
  // fall back to inline payload for tiny molecules when no ref is given.
  if (deepLink == null) deepLink = buildInlineKetcherDeepLink({ cdxml, molfile });

  const zip = await JSZip.loadAsync(inputBuf);

  // Find next free index for embeddings / media to avoid collisions.
  function nextIndex(prefix, ext) {
    let i = 1;
    while (zip.file(`${prefix}${i}.${ext}`) != null) i++;
    return i;
  }
  const mediaIdx = nextIndex('word/media/image', 'png');
  const embedIdx = nextIndex('word/embeddings/oleObject', 'bin');

  const imagePath = `word/media/image${mediaIdx}.png`;
  const oleRelPath = `embeddings/oleObject${embedIdx}.bin`;
  const olePath = `word/${oleRelPath}`;
  const imageRelPath = `media/image${mediaIdx}.png`;

  zip.file(imagePath, preview);
  zip.file(olePath, oleBin);

  // Patch word/_rels/document.xml.rels — add image + oleObject (+ optional hyperlink) relationships.
  const relsPath = 'word/_rels/document.xml.rels';
  let relsXml = await zip.file(relsPath).async('string');

  // Pick rIds that don't already exist.
  const existingIds = new Set(Array.from(relsXml.matchAll(/Id="(rId\d+)"/g)).map((m) => m[1]));
  function claimId(startN) {
    let n = startN;
    while (existingIds.has(`rId${n}`)) n++;
    existingIds.add(`rId${n}`);
    return `rId${n}`;
  }
  const imageRid = claimId(500);
  const oleRid   = claimId(501);
  const linkRid  = deepLink ? claimId(502) : null;

  const extraRels =
    `<Relationship Id="${imageRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${imageRelPath}"/>` +
    `<Relationship Id="${oleRid}"   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="${oleRelPath}"/>` +
    (linkRid
      ? `<Relationship Id="${linkRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(deepLink)}" TargetMode="External"/>`
      : '');
  relsXml = relsXml.replace('</Relationships>', `${extraRels}</Relationships>`);
  zip.file(relsPath, relsXml);

  // Patch [Content_Types].xml — make sure .png and .bin Defaults exist.
  const ctPath = '[Content_Types].xml';
  let ctXml = await zip.file(ctPath).async('string');
  if (!/Extension="png"/i.test(ctXml)) {
    ctXml = ctXml.replace('</Types>',
      '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  if (!/Extension="bin"/i.test(ctXml)) {
    ctXml = ctXml.replace('</Types>',
      '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>');
  }
  zip.file(ctPath, ctXml);

  // Patch document.xml — replace placeholder with OLE object XML.
  const docPath = 'word/document.xml';
  let docXml = await zip.file(docPath).async('string');

  const objXml = oleObjectXml({
    imageRid, oleRid,
    shapeId: `_x0000_s${1000 + embedIdx}`,
    widthPt, heightPt,
  });
  // The OLE shape itself is NOT wrapped in a hyperlink — that combination
  // breaks Word for Mac. Instead, the deep-link shows as a small underlined
  // text link right after the object, separated by a soft line break so the
  // structure gets its own visual line. When we know the SMILES we use it
  // as the link text (same as the new-doc flow); otherwise we fall back to
  // a plain "Open in Ketcher Desktop" label.
  const linkLabel = smiles ? `SMILES: ${smiles}` : 'Open in Ketcher Desktop';
  const linkInline = linkRid
    ? `<w:r><w:br/></w:r>${hyperlinkRun({ linkRid, text: linkLabel, italic: Boolean(smiles) })}`
    : '';
  const combined = `<w:r>${objXml}</w:r>${linkInline}`;

  // The replacement is sequence of runs (not a fragment of a run), so the
  // cross-run replacer wraps it cleanly between the surviving before/after
  // text pieces.
  const patched = replacePlaceholderAcrossRuns(docXml, placeholder, combined);
  if (!patched) {
    throw new Error(
      `Placeholder "${placeholder}" not found in document.\n\n` +
      `Open the source .docx, type ${placeholder} as plain text where you want ` +
      `the structure to appear, save, and try again.\n\n` +
      `Tip: avoid retyping any of the braces or colons in a different ` +
      `formatting span — Word treats each edit as a separate run, which is ` +
      `what this tool stitches back together. If you're sure the text is in ` +
      `the document, try Cmd-A, Cut, Paste in place — that often re-unifies ` +
      `split runs.`
    );
  }
  docXml = patched;
  zip.file(docPath, docXml);

  return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { buildDocxWithCdx, insertCdxIntoDocx };
