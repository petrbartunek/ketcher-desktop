#!/usr/bin/env node
// Standalone CLI: insert a CDX / CDXML structure into an existing .docx
// at a "{{CDX:1}}" placeholder. Useful if you already have the structure
// exported as a file and want to merge it into a pre-formatted template
// without opening Ketcher Desktop.
//
// Usage:
//   node tools/insert-cdx.js \
//        --in  template.docx \
//        --out report.docx \
//        --cdx structure.cdx                  # or --cdxml structure.cdxml
//        [--png preview.png]                  # optional preview bitmap
//        [--placeholder '{{CDX:1}}']
//
// Extract a bundle exported by the "Export to Word" menu and reuse its
// CDX + PNG from disk.

'use strict';

const fs = require('fs');
const path = require('path');
const { insertCdxIntoDocx } = require('../electron/docx-builder');

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k.startsWith('--')) {
      const key = k.slice(2);
      const val = argv[i + 1];
      out[key] = val;
      i++;
    }
  }
  return out;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const required = ['in', 'out'];
  for (const r of required) {
    if (!args[r]) {
      console.error(`Missing --${r}`);
      printUsage();
      process.exit(2);
    }
  }
  if (!args.cdx && !args.cdxml) {
    console.error('Provide at least one of --cdx or --cdxml');
    printUsage();
    process.exit(2);
  }

  const inBuf = await fs.promises.readFile(args.in);
  const cdxBase64 = args.cdx
    ? (await fs.promises.readFile(args.cdx)).toString('base64')
    : null;
  const cdxml = args.cdxml
    ? await fs.promises.readFile(args.cdxml, 'utf8')
    : null;
  const pngBase64 = args.png
    ? (await fs.promises.readFile(args.png)).toString('base64')
    : null;

  const placeholder = args.placeholder || '{{CDX:1}}';

  const outBuf = await insertCdxIntoDocx(inBuf, {
    placeholder, cdxBase64, cdxml, pngBase64,
  });

  await fs.promises.writeFile(args.out, outBuf);
  console.log(`Wrote ${path.resolve(args.out)}`);
}

function printUsage() {
  console.error(`
Usage:
  node tools/insert-cdx.js --in template.docx --out result.docx \\
       --cdx structure.cdx [--cdxml structure.cdxml] [--png preview.png] \\
       [--placeholder "{{CDX:1}}"]

The placeholder must appear as plain text somewhere in the source document
(ideally on its own in a paragraph). The tool replaces it with a ChemDraw
OLE object that recipients running ChemDraw can double-click to edit.
`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
