# Ketcher Desktop

A desktop chemical structure editor built by wrapping EPAM's
[Ketcher](https://github.com/epam/ketcher) in an Electron shell, with
extra workflows for embedding structures into Microsoft Word as
clickable ChemDraw OLE objects.

You get the real Ketcher UI — atom and bond tools, ring templates
(benzene, cyclohexane, cyclopentane, and the full template library),
stereochemistry, SMILES / MOL / SDF / RXN / InChI / CDXML / CDX import
and export — plus a File menu that reads and writes files on your
disk and can generate Word documents.

Ketcher itself runs entirely in the renderer via `ketcher-standalone`
(an Indigo WebAssembly bundle), so there is no server and no network
dependency at runtime.

---

## Requirements

- Node.js 18 or newer
- npm 9 or newer
- Microsoft Word on any recipient machine that needs to *view* the
  embedded structure
- ChemDraw (any recent version) on any recipient machine that needs to
  *double-click and edit* the embedded structure

## Install

```bash
cd ketcher-desktop
npm install
```

## Run

```bash
npm run dev            # hot-reload dev mode (Vite + Electron)
npm start              # production build, launched in Electron
npm run package        # native installer(s) for the current OS, in release/
npm run package:mac    # force macOS build (arm64 + x64 DMGs)
npm run package:win    # force Windows NSIS installer (needs Wine on non-Win)
npm run package:linux  # force Linux AppImage
npm run package:all    # all three at once (Windows step needs Wine)
```

## Distribution across platforms

Installers land in `release/` with arch-tagged filenames like
`Ketcher Desktop-0.1.0-arm64.dmg`. The matrix you can actually build
depends on where you run electron-builder:

| Host OS  | macOS arm64 | macOS x64 | Windows x64 | Linux x64 |
| -------- | :---------: | :-------: | :---------: | :-------: |
| macOS    |      ✅      |     ✅     |    ✅ (Wine) |    ✅      |
| Windows  |      ❌      |     ❌     |     ✅       |    ✅      |
| Linux    |      ❌      |     ❌     |    ✅ (Wine) |    ✅      |

macOS builds require an Apple toolchain, so you can't produce a `.dmg`
from Windows or Linux — full stop. For a one-command build of every
platform, use the included GitHub Actions workflow
(`.github/workflows/release.yml`): push a tag like `v0.1.0` and CI
runs the matrix on native `macos-latest` / `windows-latest` /
`ubuntu-latest` runners and attaches the three installers to a draft
release. The workflow reads code-signing secrets (`APPLE_ID`,
`APPLE_TEAM_ID`, `APPLE_APP_SPECIFIC_PASSWORD`, `MAC_CSC_LINK`,
`WIN_CSC_LINK`, etc.) if you set them; without them it produces
unsigned artifacts that are fine for internal distribution.

Apple Silicon note: the mac target builds both `arm64` and `x64` DMGs
so M-series and Intel Macs each get a native binary. If everyone at
the lab is on M-series, drop `"x64"` from `build.mac.target[0].arch`
to halve the build time.

---

## File menu

| Command                                  | Shortcut           | What it does                                    |
| ---------------------------------------- | ------------------ | ----------------------------------------------- |
| New                                      | Ctrl/Cmd-N         | Clear the canvas                                |
| Open…                                    | Ctrl/Cmd-O         | Read .mol, .sdf, .smi, .rxn, .ket, .cdxml, .cdx |
| Save As…                                 | Ctrl/Cmd-S         | Serialize in the format picked by extension     |
| Paste SMILES…                            | Ctrl/Cmd-Shift-V   | Prompt for a SMILES string and parse it         |
| Export to Word (new document)…           | Ctrl/Cmd-Shift-E   | Create a fresh .docx with the current structure |
| Insert into existing Word document…      | —                  | Replace `{{CDX:1}}` in a source .docx           |
| Append to Excel catalog…                 | Ctrl/Cmd-Shift-X   | Add one row per structure to an .xlsx catalog   |
| Change Excel Catalog…                    | —                  | Switch to a different catalog file              |

Save format is picked from the extension you type in the Save dialog:
`.mol` / `.sdf` → MDL Molfile V2000, `.smi` → SMILES, `.rxn` → RXN,
`.ket` → Ketcher native JSON, `.cdxml` → ChemDraw XML, `.cdx` → binary
CDX (if your Ketcher build supports it — see the caveats below),
`.cml` → CML, `.inchi` → InChI.

---

## Embedding structures in Word

### Quick path: Export to Word (new document)

File → **Export to Word (new document)…** produces a `.docx` that
contains:

- A rasterized preview of the structure (Ketcher's PNG export) so the
  document renders identically everywhere.
- An embedded **ChemDraw OLE object** (`ProgID=ChemDraw.Document.15`)
  so recipients with ChemDraw can double-click the picture to open
  the real editable structure.
- An optional SMILES caption below the structure.

The OLE object prefers **binary CDX** if your Ketcher build exposes
`getCDX()`; otherwise it falls back to **CDXML** packed into the same
OLE container. ChemDraw accepts both.

### Templated path: Insert into existing Word document

1. In your source `.docx`, type the literal placeholder `{{CDX:1}}`
   somewhere on its own line where the structure should go.
2. Draw the structure in Ketcher Desktop.
3. File → **Insert into existing Word document…**, pick the source
   document, then pick an output path.

The tool leaves your formatting, headers, styles, and everything else
intact. It only adds three things to the zip: a new image at
`word/media/imageN.png`, a new OLE object at
`word/embeddings/oleObjectN.bin`, and two relationship entries in
`word/_rels/document.xml.rels`.

### Headless CLI (no Ketcher Desktop needed)

If you already have a `.cdx` or `.cdxml` on disk (for example from a
batch pipeline) you can drop it straight into a template:

```bash
node tools/insert-cdx.js \
     --in  template.docx \
     --out report.docx \
     --cdx structure.cdx \
     --png preview.png \
     --placeholder '{{CDX:1}}'
```

`--cdxml` works in place of `--cdx`; `--png` is optional (a 1×1
transparent pixel is used if omitted, which is ugly but valid).

---

## Building an Excel catalog of structures

File → **Append to Excel catalog…** adds one row per export to an
`.xlsx` file — eight columns per row:

| A   | B    | C         | D      | E       | F     | G        | H       |
| --- | ---- | --------- | ------ | ------- | ----- | -------- | ------- |
| ID  | Name | Structure | SMILES | Formula | InChI | InChIKey | PubChem |

ID auto-increments (starting from 1) so every structure has a stable
handle you can reference in notes. Name is left empty at export time
for you to fill in. The Structure column carries a rasterized PNG
anchored to the cell, sized to preserve the molecule's native aspect
ratio inside a 220×150 px box; row height is adjusted to match, so
wide molecules (like indole) render without the vertical stretching
older versions produced. The header row is bold, lightly shaded, and
frozen so it stays visible as the catalog grows.

**Click-to-edit**: the SMILES cell is a live `ketcher://open?ref=<hash>`
hyperlink. Clicking it re-opens the structure in Ketcher Desktop — the
same mechanism used by the Word export. Rows created on a different
machine show a friendly "structure not cached here" dialog when
clicked, since the on-disk cache is per-user.

**PubChem similarity**: column H shows a "Similarity" link that opens
PubChem's structure-search results for the row's SMILES. It's a plain
`https://pubchem.ncbi.nlm.nih.gov/#query=<smiles>&tab=similarity` URL,
so any app that respects cell hyperlinks will follow it into a browser.

Two older catalog layouts exist in the wild — the very first version
had 5 columns (Structure/SMILES/Formula/InChI/InChIKey), and the
interim version had 7 columns (everything above through InChIKey but
without PubChem). The first time you append to such a file under the
current version, the workbook is transparently upgraded:

- **5 → 8 cols** is a full rebuild: existing rows receive sequential
  IDs, embedded images shift two columns right to the new Structure
  column, Name is left blank, and PubChem is populated from the
  existing SMILES.
- **7 → 8 cols** just adds the PubChem header and per-row Similarity
  links; nothing else is touched.

A dialog confirms the upgrade so nothing happens silently.

On first use, Ketcher Desktop asks **Create new… / Use existing… / Cancel**.
Pick "Create new…" to open a Save dialog (type a filename anywhere on
disk); pick "Use existing…" to open a file browser and point at an
`.xlsx` you already have — its current rows are preserved, new rows
append at the bottom. The chosen path is remembered for next time, so
every subsequent click appends silently to the same file with no
dialog. Use **File → Change Excel Catalog…** to switch catalogs or
re-point the setting after you've moved the file.

The remembered path lives in Electron's app-data dir:

| OS      | Path |
| ------- | ---- |
| macOS   | `~/Library/Application Support/Ketcher Desktop/config.json` |
| Windows | `%APPDATA%\Ketcher Desktop\config.json`                     |
| Linux   | `~/.config/Ketcher Desktop/config.json`                     |

### Caveats

- **InChIKey may be blank.** Ketcher's standalone JS build doesn't
  always expose `getInchiKey()`. If your build doesn't, that column
  stays empty — everything else still works.
- **Formula** prefers `getGrossFormula()` if present, otherwise it's
  pulled from the first layer of the InChI string (e.g. `C9H8O4`). Both
  are correct; occasionally their ordering of elements differs.
- **Excel must be closed** when you append — Excel holds a write lock on
  open `.xlsx` files. If it's open you'll get a friendly error; just
  close it and retry.

---

## Round-trip editing: clickable "Open in Ketcher Desktop" links

Every structure that Export-to-Word embeds is also wrapped in a
hyperlink pointing back to Ketcher Desktop via a custom URL scheme:

```
ketcher://open?ref=<12-hex-hash>
```

Clicking the structure in Word (Ctrl/Cmd-click in Word for Mac, plain
click in Word for Windows if it honors hyperlinks on images) launches
Ketcher Desktop, which resolves the hash against a local cache and
loads the structure onto the canvas — no file roundtrip, no ChemDraw
needed.

### Why a short hash and not the full structure?

Earlier versions put the whole CDXML in the URL, base64-encoded
(`?format=cdxml&data=...`). It works from the shell (`open ketcher://…`
round-trips fine on macOS) but Word for Mac **silently refuses** clicks
on custom-scheme URLs longer than ~2000 chars, surfacing only a generic
"An unexpected error has occurred". Real molecules easily blow past that
limit, so we moved the structure itself into a cache keyed by a short
content hash.

The cache lives in Electron's per-user app-data dir:

| OS      | Path |
| ------- | ---- |
| macOS   | `~/Library/Application Support/Ketcher Desktop/structures/` |
| Windows | `%APPDATA%\Ketcher Desktop\structures\`                     |
| Linux   | `~/.config/Ketcher Desktop/structures/`                     |

One small text file per structure (CDXML or MOL, a few KB each). The
cache grows unbounded; **File → Clear Structure Cache…** wipes it.

### Email caveat

Because the structure lives in the sender's local cache, a `.docx`
forwarded to a colleague will have a link that doesn't resolve on their
machine. Clicking it pops a friendly dialog pointing them back to the
OLE object — which *is* portable, so they can still double-click the
structure image to edit it in ChemDraw. If you need full portability
across machines, send the `.cdxml` / `.mol` file separately, or have
the recipient re-export from their own copy of Ketcher Desktop.

### Registering the scheme

- **Dev mode (`npm run dev`), macOS.** `app.setAsDefaultProtocolClient('ketcher')`
  is called at startup, but macOS Launch Services will silently ignore
  the request because the stock `node_modules/electron/dist/Electron.app`
  has no `CFBundleURLTypes` entry in its `Info.plist`. You'll see Word
  prompt the "Hyperlinks can be harmful…" dialog, click **Open**, and
  then get a generic "An unexpected error has occurred" — that is macOS
  telling Word there's no handler. Run this once to fix it:

  ```bash
  npm run register-scheme
  ```

  This patches the dev Electron.app's Info.plist with the scheme and
  refreshes Launch Services. `npm run unregister-scheme` undoes it.
  You need to rerun `register-scheme` after `npm install`, `npm update`,
  or any other action that rewrites `node_modules`.

- **Dev mode, Windows / Linux.** `setAsDefaultProtocolClient` writes
  registry / `.desktop` entries at runtime; no extra step needed.

- **Packaged app (`npm run package`).** electron-builder writes the
  scheme into the platform metadata (macOS `Info.plist`
  `CFBundleURLTypes`, Windows registry entry, Linux `.desktop` file)
  via the `build.protocols` block in `package.json`. This is the
  reliable path — once you install the DMG / EXE / AppImage,
  `ketcher://` URLs route to it from anywhere in the OS with no manual
  registration dance.

### Quick diagnostic

Whether registration is working on your machine:

```bash
open 'ketcher://open?format=cdxml&data=SGVsbG8='   # macOS
```

If Ketcher Desktop jumps to the front, you're wired up. If macOS says
"There is no application set to open the URL…", re-run
`npm run register-scheme`. Windows users can run `start ketcher://…`,
Linux users `xdg-open ketcher://…`.

### Multiple windows / cold launches

- If Ketcher Desktop is already running, the incoming URL is delivered
  to the existing window and the structure replaces the canvas.
- If the app is not running, the URL is queued and the structure is
  loaded as soon as the renderer has finished booting.
- On Windows and Linux, a second-instance lock forwards the URL to the
  first instance; macOS uses the native `open-url` event.

---

## Caveats and limitations (please read)

This is the honest part. OLE embedding of ChemDraw objects works
reliably only when a few things line up, and you should know where the
moving parts are.

1. **Binary CDX export from Ketcher is version-dependent.**
   `ketcher-standalone` started exposing `getCDX()` around release 2.20.
   If the desktop app logs `"CDX export failed"` in the console, you
   are likely on an older build. `getCDXml()` has been around for much
   longer and is the safer choice if you don't need binary CDX
   specifically. The export flow automatically falls back to CDXML.

2. **ChemDraw's OLE layout is not publicly documented.**
   `electron/ole-cdx.js` writes a standard OLE Compound Binary File
   with the three streams Word expects (`\x01CompObj`, `\x01Ole`,
   `CONTENTS`) and tags it with ChemDraw's ProgID and CLSID. This is
   the same layout Word itself produces when it stores ChemDraw
   objects, as far as we've been able to verify. If your recipients
   use an older ChemDraw whose CLSID differs, edit
   `CLSID_CHEMDRAW` / `PROGID_CHEMDRAW` in `electron/ole-cdx.js`
   accordingly — the most common alternate is
   `ChemDraw.Document.10` on older installs.

3. **No ChemDraw? Word still shows the picture.**
   Recipients without ChemDraw installed see the PNG preview and can't
   edit. Double-clicking pops a "the program for this type is not
   installed" dialog. This is a limitation of OLE itself, not our
   wrapper.

4. **Tables and inline text placeholders.**
   The placeholder replacement is designed for `{{CDX:1}}` sitting on
   its own in a paragraph. Placeholders inside more complex runs
   (mixed formatting, inside a table cell with footnote references,
   etc.) may still work, but the OLE object will land as the nearest
   valid location Word allows. Test with your real template.

5. **Only one structure per export.**
   The current flow embeds the *current* sketch. If you want multiple
   structures (`{{CDX:1}}`, `{{CDX:2}}`, …), extend `insert-cdx.js` to
   loop — the underlying `insertCdxIntoDocx` supports being called
   repeatedly on the same buffer.

6. **CSP.**
   The renderer enables `'unsafe-eval'` because Ketcher's WASM loader
   needs it. This is fine in a local Electron shell; if you're ever
   tempted to deploy the renderer as a real web page, rethink that.

---

## Project layout

```
ketcher-desktop/
├── electron/
│   ├── main.js           Electron main process + menus + IPC
│   ├── preload.js        contextBridge exposing `window.desktop`
│   ├── ole-cdx.js        OLE CFB writer for ChemDraw objects
│   └── docx-builder.js   DOCX assembler + template patcher
├── src/
│   ├── index.html        Renderer entry HTML
│   ├── main.jsx          React bootstrap
│   └── App.jsx           Mounts <Editor> from ketcher-react
├── tools/
│   └── insert-cdx.js     Headless CLI for injecting CDX into a .docx
├── vite.config.js
├── package.json
└── README.md
```

## License

Ketcher is Apache-2.0. This shell is published under the same license.
