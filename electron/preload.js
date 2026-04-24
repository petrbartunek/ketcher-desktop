// Preload — exposes a tiny, safe API to the renderer via contextBridge.
// The renderer can:
//   - listen for File-menu events from the main process
//   - ask main to show Open/Save dialogs
//   - reply with the current structure when main asks for it during Save.

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktop', {
  // Events coming FROM main (File menu)
  onNew: (cb) => ipcRenderer.on('menu:new', () => cb()),
  onPasteSmiles: (cb) => ipcRenderer.on('menu:paste-smiles', () => cb()),
  onFileOpened: (cb) =>
    ipcRenderer.on('file:opened', (_e, payload) => cb(payload)),

  // Request-reply: main asks renderer for the current structure when saving.
  onStructureRequest: (handler) => {
    ipcRenderer.on('renderer:get-structure', async (_e, { ext, replyChannel }) => {
      try {
        const content = await handler(ext);
        ipcRenderer.send(replyChannel, { content });
      } catch (err) {
        ipcRenderer.send(replyChannel, { error: String(err.message || err) });
      }
    });
  },

  // Renderer-initiated dialogs (toolbar buttons)
  openDialog: () => ipcRenderer.invoke('dialog:open'),
  saveDialog: () => ipcRenderer.invoke('dialog:save'),

  // Export-to-Word flow: main asks renderer for a bundle (CDX + CDXML + PNG + SVG).
  onExportBundleRequest: (handler) => {
    ipcRenderer.on('renderer:get-export-bundle', async (_e, { replyChannel }) => {
      try {
        const bundle = await handler();
        ipcRenderer.send(replyChannel, { bundle });
      } catch (err) {
        ipcRenderer.send(replyChannel, { error: String(err.message || err) });
      }
    });
  },

  // Deep-link arrivals from main (ketcher://open?format=...&data=...).
  onDeepLink: (cb) => ipcRenderer.on('deep-link', (_e, url) => cb(url)),
});
