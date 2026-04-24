// Electron main process — creates the window, File menu, and handles
// open/save dialogs for .mol / .sdf / .smi / .smiles. The renderer embeds
// EPAM Ketcher via ketcher-react + ketcher-standalone (all client-side,
// no server needed).

const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const { URL } = require('url');
const { buildDocxWithCdx, insertCdxIntoDocx } = require('./docx-builder');
const { appendToXlsxCatalog } = require('./xlsx-builder');
const structureCache = require('./cache');
const appConfig = require('./app-config');

// --- Deep-link support -----------------------------------------------------
// Register ketcher:// so clicking a hyperlink in Word (or anywhere else)
// routes here, launching / focusing the app with a pre-loaded structure.
// On macOS the URL arrives via the 'open-url' event; on Windows/Linux it
// arrives as a command-line argument handled via the single-instance lock.

const PROTOCOL = 'ketcher';

if (process.defaultApp) {
  // Electron in dev mode: register with the node executable that's running us.
  if (process.argv.length >= 2) {
    app.setAsDefaultProtocolClient(PROTOCOL, process.execPath, [path.resolve(process.argv[1])]);
  }
} else {
  app.setAsDefaultProtocolClient(PROTOCOL);
}

// Pending URL in case the app launched cold from a ketcher:// click and
// the renderer isn't ready yet.
let pendingDeepLink = null;

// base64url-encode a UTF-8 string (used when materializing a short ref back
// into the inline URL shape the renderer already knows how to parse).
function base64url(str) {
  return Buffer.from(str, 'utf8').toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Short-ref URLs look like `ketcher://open?ref=<hash>`. Resolve the hash
// against the on-disk cache and rewrite to the inline form the renderer
// understands (`?format=<f>&data=<base64url>`).
//
// Returns:
//   { url, missingRef?: string }
//     `url` is the URL to send to the renderer (possibly rewritten), or
//     the original. `missingRef` is set when the URL carried a ref that
//     we couldn't resolve locally — typically because the doc was
//     exported on a different machine. Callers should surface a dialog
//     in that case instead of sending a broken URL to the renderer.
async function materializeRefUrl(url) {
  try {
    const u = new URL(url);
    const host = u.hostname || u.pathname.replace(/^\/+/, '');
    if (host !== 'open') return { url };
    const ref = u.searchParams.get('ref');
    if (!ref) return { url }; // already inline; pass through
    const hit = await structureCache.resolveRef(ref);
    if (!hit) return { url, missingRef: ref };
    return { url: `ketcher://open?format=${encodeURIComponent(hit.format)}&data=${base64url(hit.data)}` };
  } catch (err) {
    console.warn('[deep-link] failed to materialize ref URL:', err);
    return { url };
  }
}

async function deliverDeepLink(url) {
  if (!url || !url.startsWith(`${PROTOCOL}://`)) return;
  const { url: resolved, missingRef } = await materializeRefUrl(url);

  // Bring the window forward regardless of whether we can open the structure
  // — the user clicked something, so *something* should happen on screen.
  // On macOS the window can be closed while the app is still running (its
  // reference lingers but is destroyed), so guard with isDestroyed().
  if (mainWindow && !mainWindow.isDestroyed()) {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  } else if (process.platform === 'darwin') {
    createWindow();
  }

  if (missingRef) {
    // Doc was exported on a different machine, or the cache was cleared.
    // The OLE object in Word still works (double-click for ChemDraw);
    // we just can't re-hydrate the Ketcher canvas from here.
    if (mainWindow) {
      dialog.showMessageBox(mainWindow, {
        type: 'info',
        title: 'Structure not in local cache',
        message: 'This link points to a structure that was exported on another machine.',
        detail: `Reference: ${missingRef}\n\n` +
                'Double-click the structure image in Word to edit it in ChemDraw, ' +
                'or ask the document\u2019s author to re-send their Ketcher source.',
      });
    }
    return;
  }

  if (!mainWindow || mainWindow.isDestroyed()) { pendingDeepLink = resolved; return; }
  if (mainWindow.webContents.isLoading()) {
    pendingDeepLink = resolved;
    mainWindow.webContents.once('did-finish-load', () => {
      const u = pendingDeepLink;
      pendingDeepLink = null;
      if (u) mainWindow.webContents.send('deep-link', u);
    });
    return;
  }
  mainWindow.webContents.send('deep-link', resolved);
}

app.on('open-url', (event, url) => {
  event.preventDefault();
  deliverDeepLink(url);
});

// Single-instance lock for Windows/Linux deep-link support.
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    const url = argv.find((a) => typeof a === 'string' && a.startsWith(`${PROTOCOL}://`));
    if (url) deliverDeepLink(url);
    // The window may have been closed (its reference lingers after close on
    // macOS, since the app stays alive in the Dock). Re-create in that case.
    if (mainWindow && !mainWindow.isDestroyed()) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    } else {
      createWindow();
    }
  });
}

const isDev = !app.isPackaged;

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: 'Ketcher Desktop',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  // Null out the reference when the window closes so any late event
  // handler ('second-instance', 'open-url') falls through the guard
  // instead of dereferencing a destroyed BrowserWindow.
  mainWindow.on('closed', () => { mainWindow = null; });

  // Deliver any pending deep-link once the renderer finishes loading.
  mainWindow.webContents.once('did-finish-load', () => {
    if (pendingDeepLink) {
      const u = pendingDeepLink;
      pendingDeepLink = null;
      mainWindow.webContents.send('deep-link', u);
    }
  });
}

// --- File menu -------------------------------------------------------------

function buildMenu() {
  const isMac = process.platform === 'darwin';

  const template = [
    ...(isMac
      ? [{
          label: app.name,
          submenu: [
            { role: 'about' },
            { type: 'separator' },
            { role: 'services' },
            { type: 'separator' },
            { role: 'hide' },
            { role: 'hideOthers' },
            { role: 'unhide' },
            { type: 'separator' },
            { role: 'quit' },
          ],
        }]
      : []),
    {
      label: 'File',
      submenu: [
        {
          label: 'New',
          accelerator: 'CmdOrCtrl+N',
          click: () => mainWindow?.webContents.send('menu:new'),
        },
        { type: 'separator' },
        {
          label: 'Open\u2026',
          accelerator: 'CmdOrCtrl+O',
          click: handleOpen,
        },
        {
          label: 'Save As\u2026',
          accelerator: 'CmdOrCtrl+S',
          click: handleSave,
        },
        { type: 'separator' },
        {
          label: 'Paste SMILES\u2026',
          accelerator: 'CmdOrCtrl+Shift+V',
          click: () => mainWindow?.webContents.send('menu:paste-smiles'),
        },
        { type: 'separator' },
        {
          label: 'Export to Word (new document)\u2026',
          accelerator: 'CmdOrCtrl+Shift+E',
          click: handleExportToWordNew,
        },
        {
          label: 'Insert into existing Word document\u2026',
          click: handleInsertIntoExistingDocx,
        },
        { type: 'separator' },
        {
          label: 'Append to Excel catalog\u2026',
          accelerator: 'CmdOrCtrl+Shift+X',
          click: handleAppendToExcelCatalog,
        },
        {
          label: 'Change Excel Catalog\u2026',
          click: handleChangeExcelCatalog,
        },
        { type: 'separator' },
        {
          label: 'Clear Structure Cache\u2026',
          click: handleClearStructureCache,
        },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'selectAll' },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'forceReload' },
        { role: 'toggleDevTools' },
        { type: 'separator' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
    {
      role: 'help',
      submenu: [
        {
          label: 'Ketcher on GitHub',
          click: () => shell.openExternal('https://github.com/epam/ketcher'),
        },
      ],
    },
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// --- Open / Save handlers --------------------------------------------------

const FILE_FILTERS = [
  { name: 'Chemical files', extensions: ['mol', 'sdf', 'smi', 'smiles', 'rxn', 'ket'] },
  { name: 'MOL / SDF', extensions: ['mol', 'sdf'] },
  { name: 'SMILES', extensions: ['smi', 'smiles'] },
  { name: 'Ketcher native', extensions: ['ket'] },
  { name: 'All files', extensions: ['*'] },
];

function extOf(filePath) {
  return path.extname(filePath).toLowerCase().replace(/^\./, '');
}

async function handleOpen() {
  if (!mainWindow) return;
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: 'Open structure',
    properties: ['openFile'],
    filters: FILE_FILTERS,
  });
  if (canceled || !filePaths.length) return;

  const filePath = filePaths[0];
  try {
    const content = await fs.readFile(filePath, 'utf8');
    mainWindow.webContents.send('file:opened', {
      path: filePath,
      ext: extOf(filePath),
      content,
    });
  } catch (err) {
    dialog.showErrorBox('Open failed', String(err.message || err));
  }
}

async function handleSave() {
  if (!mainWindow) return;
  // Ask renderer for the current structure in the format we want to save.
  // The renderer returns { content, extHint } based on user choice.
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Save structure',
    defaultPath: 'structure.mol',
    filters: FILE_FILTERS,
  });
  if (canceled || !filePath) return;

  const ext = extOf(filePath) || 'mol';
  try {
    // Main → renderer request/reply: send on a well-known channel with a
    // unique reply channel, renderer answers on that unique channel.
    const content = await new Promise((resolve, reject) => {
      const replyChannel = `renderer:structure-reply:${Date.now()}:${Math.random().toString(36).slice(2)}`;
      const timeout = setTimeout(() => {
        ipcMain.removeAllListeners(replyChannel);
        reject(new Error('Timed out waiting for renderer to serialize structure'));
      }, 15000);
      ipcMain.once(replyChannel, (_e, payload) => {
        clearTimeout(timeout);
        if (payload && payload.error) reject(new Error(payload.error));
        else resolve(payload.content);
      });
      mainWindow.webContents.send('renderer:get-structure', { ext, replyChannel });
    });

    await fs.writeFile(filePath, content, 'utf8');
  } catch (err) {
    dialog.showErrorBox('Save failed', String(err.message || err));
  }
}

// --- Export to Word --------------------------------------------------------

// Ask renderer for the current structure in every format we want to embed.
async function fetchExportBundle() {
  return await new Promise((resolve, reject) => {
    const replyChannel = `renderer:export-reply:${Date.now()}:${Math.random().toString(36).slice(2)}`;
    const timeout = setTimeout(() => {
      ipcMain.removeAllListeners(replyChannel);
      reject(new Error('Timed out waiting for renderer'));
    }, 20000);
    ipcMain.once(replyChannel, (_e, payload) => {
      clearTimeout(timeout);
      if (payload?.error) reject(new Error(payload.error));
      else resolve(payload.bundle);
    });
    mainWindow.webContents.send('renderer:get-export-bundle', { replyChannel });
  });
}

async function handleExportToWordNew() {
  if (!mainWindow) return;
  try {
    const bundle = await fetchExportBundle();
    if (!bundle.cdxBase64 && !bundle.cdxml) {
      dialog.showErrorBox('Export failed',
        'Neither CDX nor CDXML could be produced by this Ketcher build. ' +
        'Check that ketcher-standalone is new enough to expose getCDX() / getCDXml().');
      return;
    }

    const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
      title: 'Export to Word',
      defaultPath: 'structure.docx',
      filters: [{ name: 'Word document', extensions: ['docx'] }],
    });
    if (canceled || !filePath) return;

    // Stash the structure in the on-disk cache and embed only the short
    // ref in the docx hyperlink — Word for Mac silently rejects clickable
    // custom-protocol URLs longer than ~2000 chars, which the old inline
    // payload easily exceeded.
    const cacheEntry = await structureCache.put({
      cdxml: bundle.cdxml,
      molfile: bundle.molfile,
    });
    const deepLink = cacheEntry?.url ?? null;

    const buffer = await buildDocxWithCdx({
      cdxBase64: bundle.cdxBase64,
      cdxml: bundle.cdxml,
      molfile: bundle.molfile,
      pngBase64: bundle.pngBase64,
      svg: bundle.svg,
      caption: bundle.smiles ? `SMILES: ${bundle.smiles}` : null,
      deepLink,
    });
    await fs.writeFile(filePath, buffer);
    shell.showItemInFolder(filePath);
  } catch (err) {
    dialog.showErrorBox('Export failed', String(err.message || err));
  }
}

async function handleInsertIntoExistingDocx() {
  if (!mainWindow) return;
  try {
    const bundle = await fetchExportBundle();
    if (!bundle.cdxBase64 && !bundle.cdxml) {
      dialog.showErrorBox('Export failed',
        'Neither CDX nor CDXML could be produced by this Ketcher build.');
      return;
    }

    const { canceled: cOpen, filePaths } = await dialog.showOpenDialog(mainWindow, {
      title: 'Choose Word document to insert into',
      properties: ['openFile'],
      filters: [{ name: 'Word document', extensions: ['docx'] }],
    });
    if (cOpen || !filePaths.length) return;
    const sourcePath = filePaths[0];

    const { canceled: cSave, filePath: destPath } = await dialog.showSaveDialog(mainWindow, {
      title: 'Save updated document as',
      defaultPath: sourcePath.replace(/\.docx$/i, ' (with structure).docx'),
      filters: [{ name: 'Word document', extensions: ['docx'] }],
    });
    if (cSave || !destPath) return;

    const placeholder = '{{CDX:1}}'; // Users put this in the source doc where they want the structure.
    const input = await fs.readFile(sourcePath);

    // Same short-ref dance as the "new doc" flow — see handleExportToWordNew
    // for why this matters (Word for Mac URL-length limit).
    const cacheEntry = await structureCache.put({
      cdxml: bundle.cdxml,
      molfile: bundle.molfile,
    });
    const deepLink = cacheEntry?.url ?? null;

    const output = await insertCdxIntoDocx(input, {
      placeholder,
      cdxBase64: bundle.cdxBase64,
      cdxml: bundle.cdxml,
      molfile: bundle.molfile,
      pngBase64: bundle.pngBase64,
      svg: bundle.svg,
      smiles: bundle.smiles,
      deepLink,
    });
    await fs.writeFile(destPath, output);
    shell.showItemInFolder(destPath);
  } catch (err) {
    dialog.showErrorBox('Insert failed', String(err.message || err));
  }
}

// --- Export to Excel catalog ----------------------------------------------
//
// UX model: the user picks a catalog file once (the first time they click
// Append to Excel catalog…), we persist that path in app-config.json, and
// every subsequent append lands in the same file silently. A separate
// "Change Excel Catalog…" entry lets them switch catalogs or reset the
// path if the file was moved / deleted.
//
// This keeps the menu click from feeling like a full Save-As dialog every
// time, which was the whole point of the "catalog" framing — you end up
// with one long spreadsheet of structures accumulated over time.

// Two-step picker: first ask whether the user wants a brand-new file or
// an existing one, then show the matching dialog. This replaces a single
// showSaveDialog, which on macOS pops a misleading "Replace?" prompt when
// the user picks an existing file to append to — we don't replace, we
// append, and the OS-level prompt is wrong about what's happening.
async function promptForCatalog(title = 'Choose Excel catalog') {
  if (!mainWindow) return null;

  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    title,
    message: 'Where should structures be saved?',
    detail:
      'Create a brand-new .xlsx catalog, or append to one you already have.\n\n' +
      'Tip: an existing file keeps all its current rows — we just add new ' +
      'rows at the bottom.',
    buttons: ['Create new\u2026', 'Use existing\u2026', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  });

  if (response === 2) return null;

  if (response === 0) {
    const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
      title: 'Create new Excel catalog',
      defaultPath: 'ketcher-catalog.xlsx',
      filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
      properties: ['createDirectory'],
    });
    return canceled || !filePath ? null : filePath;
  }

  // response === 1 — use existing
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: 'Open existing Excel catalog',
    filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });
  return canceled || !filePaths.length ? null : filePaths[0];
}

async function getOrPickCatalogPath() {
  let filePath = await appConfig.get('xlsxCatalogPath');
  if (filePath) {
    // Forget a path whose parent directory no longer exists (e.g. external
    // drive unplugged, folder moved). If the *file* itself is gone but the
    // directory is there, we keep the setting — the next append will
    // recreate the file at that location.
    try {
      await fs.access(path.dirname(filePath));
    } catch {
      filePath = null;
    }
  }
  if (!filePath) {
    filePath = await promptForCatalog('Choose an Excel catalog');
    if (!filePath) return null;
    await appConfig.set('xlsxCatalogPath', filePath);
  }
  return filePath;
}

async function handleAppendToExcelCatalog() {
  if (!mainWindow) return;
  try {
    const bundle = await fetchExportBundle();
    if (!bundle.pngBase64 && !bundle.smiles) {
      dialog.showErrorBox('Excel export failed',
        'The canvas is empty. Draw a structure before appending to the catalog.');
      return;
    }

    const filePath = await getOrPickCatalogPath();
    if (!filePath) return;

    // Same trick as the Word export: stash the structure in the on-disk
    // cache and keep only a short `ketcher://open?ref=<hash>` URL in the
    // spreadsheet. Clicking the SMILES cell then re-opens the structure
    // in Ketcher Desktop.
    let deepLink = null;
    try {
      const cacheEntry = await structureCache.put({
        cdxml:   bundle.cdxml,
        molfile: bundle.molfile,
      });
      deepLink = cacheEntry?.url ?? null;
    } catch (err) {
      console.warn('[xlsx] failed to stash structure in cache:', err);
    }

    const { rowNumber, created, migrated } = await appendToXlsxCatalog(filePath, {
      pngBase64: bundle.pngBase64,
      smiles:    bundle.smiles,
      formula:   bundle.formula,
      inchi:     bundle.inchi,
      inchiKey:  bundle.inchiKey,
      deepLink,
    });

    // On first creation, reveal the new file so the user knows where it
    // landed. On subsequent appends, just show a quiet toast-ish info
    // dialog mentioning the row number so they can jump to it. When we
    // auto-migrated from the old 5-column schema, mention that so the
    // user isn't surprised by the new ID / Name columns.
    if (created) {
      shell.showItemInFolder(filePath);
    } else {
      const detailLines = [filePath];
      if (migrated) {
        detailLines.push('');
        detailLines.push(
          'This catalog used the older 5-column layout; it has been upgraded ' +
          'to include ID and Name columns. Existing rows received sequential ' +
          'IDs and an empty Name — feel free to fill them in.'
        );
      }
      await dialog.showMessageBox(mainWindow, {
        type: 'info',
        title: migrated ? 'Catalog upgraded and appended' : 'Appended to catalog',
        message: `Row ${rowNumber} added to ${path.basename(filePath)}.`,
        detail: detailLines.join('\n'),
        buttons: ['OK', 'Reveal in Finder'],
        defaultId: 0,
        cancelId: 0,
      }).then((r) => { if (r.response === 1) shell.showItemInFolder(filePath); });
    }
  } catch (err) {
    dialog.showErrorBox('Excel export failed', String(err.message || err));
  }
}

async function handleChangeExcelCatalog() {
  try {
    const current = await appConfig.get('xlsxCatalogPath');
    const picked = await promptForCatalog(
      current ? `Change catalog (current: ${path.basename(current)})` : 'Choose Excel catalog'
    );
    if (!picked) return;
    await appConfig.set('xlsxCatalogPath', picked);
    if (mainWindow) {
      dialog.showMessageBox(mainWindow, {
        type: 'info',
        message: 'Excel catalog updated.',
        detail: picked,
      });
    }
  } catch (err) {
    dialog.showErrorBox('Change catalog failed', String(err.message || err));
  }
}

// --- Structure cache management -------------------------------------------

// Lets the user nuke every structure stashed in the deep-link cache. Links
// in already-distributed .docx files will stop resolving on this machine
// after this; the OLE object (ChemDraw double-click) keeps working either
// way, so no data is lost — the cache only ever held a round-trip
// convenience, not the source of truth.
async function handleClearStructureCache() {
  if (!mainWindow) return;
  try {
    const count = await structureCache.size();
    const { response } = await dialog.showMessageBox(mainWindow, {
      type: 'question',
      buttons: ['Cancel', 'Clear'],
      defaultId: 0,
      cancelId: 0,
      title: 'Clear Structure Cache',
      message: count === 0
        ? 'The structure cache is already empty.'
        : `Delete ${count} cached structure${count === 1 ? '' : 's'}?`,
      detail: 'Already-exported .docx files will keep their embedded ChemDraw ' +
              'objects — only the "click SMILES to re-open in Ketcher Desktop" ' +
              'links for those documents will stop working on this machine.',
    });
    if (response !== 1) return;
    const removed = await structureCache.clear();
    await dialog.showMessageBox(mainWindow, {
      type: 'info',
      message: `Removed ${removed} cached structure${removed === 1 ? '' : 's'}.`,
    });
  } catch (err) {
    dialog.showErrorBox('Clear cache failed', String(err.message || err));
  }
}

// --- App lifecycle ---------------------------------------------------------

app.whenReady().then(() => {
  createWindow();
  buildMenu();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Allow the renderer to trigger dialogs from in-page toolbar buttons too.
ipcMain.handle('dialog:open', handleOpen);
ipcMain.handle('dialog:save', handleSave);
