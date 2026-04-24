import React, { useEffect, useRef, useState } from 'react';
import { Editor } from 'ketcher-react';
import { StandaloneStructServiceProvider } from 'ketcher-standalone';
import 'ketcher-react/dist/index.css';

const structServiceProvider = new StandaloneStructServiceProvider();

function formatForExt(ext) {
  switch ((ext || '').toLowerCase()) {
    case 'mol':    return 'mol';
    case 'sdf':    return 'sdf';
    case 'smi':
    case 'smiles': return 'smiles';
    case 'rxn':    return 'rxn';
    case 'ket':    return 'ket';
    case 'cml':    return 'cml';
    case 'inchi':  return 'inchi';
    case 'cdx':    return 'cdx';
    case 'cdxml':  return 'cdxml';
    default:       return 'mol';
  }
}

// Ketcher API helpers — the exact method names differ slightly across
// versions, so we fall back to the generic getStructure(format) path.
async function getStructureIn(ketcher, format) {
  if (!ketcher) throw new Error('Editor not ready');

  try {
    switch (format) {
      case 'smiles': return await ketcher.getSmiles();
      case 'inchi':  return await ketcher.getInchi();
      case 'ket':    return await ketcher.getKet();
      case 'rxn':    return await ketcher.getRxn();
      case 'cml':    return await ketcher.getCml();
      case 'cdxml':
        if (typeof ketcher.getCDXml === 'function') return await ketcher.getCDXml();
        break;
      case 'cdx':
        // Binary CDX: ketcher-standalone 2.20+ exposes getCDX() returning
        // a base64 string. Older builds don't have it.
        if (typeof ketcher.getCDX === 'function') return await ketcher.getCDX();
        break;
      case 'mol':
      case 'sdf':
        return await ketcher.getMolfile();
    }
  } catch (err) {
    // Fall through to generic path below.
    console.warn('[ketcher] typed getter failed, trying generic', err);
  }

  // Generic fallback: Ketcher exposes a low-level structure service.
  if (typeof ketcher.getStructure === 'function') {
    return await ketcher.getStructure(format);
  }
  throw new Error(`Format "${format}" is not supported by this Ketcher build`);
}

// Render the current sketch as SVG and PNG. We use PNG for Word's OLE
// fallback thumbnail and keep SVG around for potential future use.
//
// IMPORTANT: ketcher.generateImage(data, options) takes the actual
// structure string as its first argument, not a format name. Ketcher's
// SMILES parser will happily choke on "mol" or "png" if those get passed
// in by mistake. We fetch the current molfile first and use that.
async function getPreviewImages(ketcher) {
  // Get the current structure as a molfile; generateImage will layout & render it.
  let molfile = '';
  try {
    molfile = await ketcher.getMolfile();
  } catch (_) {
    // Empty canvas — no structure drawn. Use a tiny empty placeholder.
    molfile = '\n  Ketcher\n\n  0  0  0  0  0  0  0  0  0  0999 V2000\nM  END\n';
  }

  // SVG: use the dedicated getter if present, fall back to generateImage.
  let svgText = '';
  try {
    if (typeof ketcher.getSvg === 'function') {
      svgText = await ketcher.getSvg();
    } else if (typeof ketcher.generateImage === 'function') {
      const svgBlob = await ketcher.generateImage(molfile, { outputFormat: 'svg' });
      if (svgBlob && typeof svgBlob.text === 'function') {
        svgText = await svgBlob.text();
      }
    }
  } catch (err) {
    console.warn('[ketcher] SVG generation failed', err);
  }

  // PNG: rasterized preview for Word. We want the picture to stay crisp
  // even when Word downscales to ~3.5 inches, so we rasterize the SVG
  // ourselves at ~3× the display size instead of relying on Ketcher's default.
  //
  // Strategy:
  //   1. Start from the SVG (which is resolution-independent).
  //   2. Parse its intrinsic size from the width/height attrs or viewBox.
  //   3. Draw it onto a canvas scaled up 3×, then export PNG from the canvas.
  //
  // If anything goes sideways we fall back to Ketcher's built-in PNG.
  async function rasterizeSvgToPng(svg, scale = 3) {
    // Extract natural dimensions from the SVG root.
    const parser = new DOMParser();
    const doc = parser.parseFromString(svg, 'image/svg+xml');
    const root = doc.documentElement;

    let w = parseFloat(root.getAttribute('width'));
    let h = parseFloat(root.getAttribute('height'));
    if (!w || !h) {
      const vb = (root.getAttribute('viewBox') || '').split(/\s+/).map(Number);
      if (vb.length === 4) { w = vb[2]; h = vb[3]; }
    }
    if (!w || !h) { w = 600; h = 450; }

    // Make sure the SVG has an explicit size so the browser honors scale.
    root.setAttribute('width',  String(w));
    root.setAttribute('height', String(h));
    const serialized = new XMLSerializer().serializeToString(root);

    // Use a data URL (avoids blob-URL same-origin quirks in Electron).
    const encoded = encodeURIComponent(serialized)
      .replace(/'/g, '%27').replace(/"/g, '%22');
    const src = `data:image/svg+xml;charset=utf-8,${encoded}`;

    const img = await new Promise((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = reject;
      el.src = src;
    });

    const canvas = document.createElement('canvas');
    canvas.width  = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    // Convert to blob → base64.
    return await new Promise((resolve, reject) => {
      canvas.toBlob(async (blob) => {
        if (!blob) return reject(new Error('toBlob returned null'));
        const buf = new Uint8Array(await blob.arrayBuffer());
        let bin = '';
        for (let i = 0; i < buf.length; i++) bin += String.fromCharCode(buf[i]);
        resolve(btoa(bin));
      }, 'image/png');
    });
  }

  let pngBase64 = '';
  try {
    if (svgText) {
      pngBase64 = await rasterizeSvgToPng(svgText, 3);
    }
  } catch (err) {
    console.warn('[ketcher] SVG→PNG rasterization failed, falling back', err);
  }
  if (!pngBase64) {
    try {
      if (typeof ketcher.generateImage === 'function') {
        const pngBlob = await ketcher.generateImage(molfile, { outputFormat: 'png' });
        if (pngBlob) {
          const buf = new Uint8Array(await pngBlob.arrayBuffer());
          let bin = '';
          for (let i = 0; i < buf.length; i++) bin += String.fromCharCode(buf[i]);
          pngBase64 = btoa(bin);
        }
      }
    } catch (err) {
      console.warn('[ketcher] PNG generation failed', err);
    }
  }

  return { svg: svgText, pngBase64 };
}

export default function App() {
  const ketcherRef = useRef(null);
  const [ready, setReady] = useState(false);

  const onInit = (ketcher) => {
    ketcherRef.current = ketcher;
    window.ketcher = ketcher;
    setReady(true);
  };

  useEffect(() => {
    if (!window.desktop) return;

    window.desktop.onNew(async () => {
      await ketcherRef.current?.setMolecule('');
    });

    window.desktop.onFileOpened(async ({ path, content }) => {
      try {
        await ketcherRef.current.setMolecule(content);
        document.title = `Ketcher Desktop — ${path}`;
      } catch (err) {
        alert(`Could not parse file:\n${err.message || err}`);
      }
    });

    window.desktop.onPasteSmiles(async () => {
      const smi = prompt('Paste a SMILES string:');
      if (!smi) return;
      try {
        await ketcherRef.current.setMolecule(smi.trim());
      } catch (err) {
        alert(`Invalid SMILES:\n${err.message || err}`);
      }
    });

    // Main → renderer request/reply for Save As.
    window.desktop.onStructureRequest(async (ext) => {
      return await getStructureIn(ketcherRef.current, formatForExt(ext));
    });

    // Main → renderer for Export to Word / Excel: gather every
    // representation downstream builders might want.
    window.desktop.onExportBundleRequest?.(async () => {
      const k = ketcherRef.current;
      if (!k) throw new Error('Editor not ready');

      let cdxBase64 = null;
      let cdxml = null;
      let molfile = null;
      try { cdxml = await getStructureIn(k, 'cdxml'); } catch (e) { console.warn('CDXML export failed:', e); }
      try { cdxBase64 = await getStructureIn(k, 'cdx'); } catch (e) { console.warn('CDX export failed:', e); }
      try { molfile = await k.getMolfile(); } catch (_) {}

      const { svg, pngBase64 } = await getPreviewImages(k);

      // Textual identifiers — each wrapped in its own try/catch because
      // different Ketcher builds expose different subsets of these APIs.
      let smiles = '';
      let inchi = '';
      let inchiKey = '';
      let formula = '';
      try { smiles = await k.getSmiles(); } catch (_) {}
      try { if (typeof k.getInchi === 'function') inchi = await k.getInchi(); } catch (e) { console.warn('InChI export failed:', e); }
      try {
        // The exact method name moved around across ketcher-standalone
        // releases — 2.22+ uses getInchiKey (lowercase "chi"), 2.26+
        // renamed it to getInChIKey (Proper-case InChI). Newer builds
        // also expose both. We try every known spelling before giving up.
        const inchiKeyFn =
          (typeof k.getInChIKey === 'function' && k.getInChIKey) ||
          (typeof k.getInchiKey === 'function' && k.getInchiKey) ||
          (typeof k.getInchIKey === 'function' && k.getInchIKey) ||
          null;
        if (inchiKeyFn) {
          inchiKey = await inchiKeyFn.call(k);
        } else if (k.indigo) {
          // Fallback: ask Indigo (the struct service backing ketcher-standalone)
          // to convert the current structure to an InChIKey string. The exact
          // method name on indigo also varies by version, so probe a few.
          const ind = k.indigo;
          const indFn =
            (typeof ind.getInChIKey === 'function' && ind.getInChIKey) ||
            (typeof ind.getInchiKey === 'function' && ind.getInchiKey) ||
            (typeof ind.inchiKey    === 'function' && ind.inchiKey)    ||
            null;
          if (indFn) inchiKey = await indFn.call(ind);
        }
      } catch (e) { console.warn('InChIKey export failed:', e); }

      // One last resort: derive InChIKey from the aux InChI string if
      // Ketcher returned it there. ketcher.getInchi() can accept an
      // options object in newer builds that asks for "inchi-aux", which
      // includes the key after a "InChIKey=" prefix.
      if (!inchiKey) {
        try {
          if (typeof k.getInchi === 'function') {
            const aux = await k.getInchi({ 'output-format': 'inchi-aux' });
            if (typeof aux === 'string') {
              const m = aux.match(/InChIKey=([A-Z0-9-]+)/);
              if (m) inchiKey = m[1];
            }
          }
        } catch (_) { /* ignore — column simply stays empty */ }
      }

      // Gross formula: prefer the Ketcher/Indigo getter if present; fall
      // back to pulling the formula layer out of the InChI string
      // (which starts right after "InChI=1S/" or "InChI=1/").
      try {
        if (typeof k.getGrossFormula === 'function') {
          formula = await k.getGrossFormula();
        } else if (inchi) {
          const m = inchi.match(/^InChI=1S?\/([^\/]+)/);
          if (m) formula = m[1];
        }
      } catch (e) { console.warn('Formula extraction failed:', e); }

      return { cdxBase64, cdxml, molfile, svg, pngBase64, smiles, inchi, inchiKey, formula };
    });

    // Deep link handler: ketcher://open?format=cdxml&data=<base64url>
    window.desktop.onDeepLink?.(async (url) => {
      try {
        const u = new URL(url);
        if (u.hostname !== 'open' && u.pathname.replace(/^\//, '') !== 'open') return;
        const format = u.searchParams.get('format') || 'cdxml';
        const data = u.searchParams.get('data') || '';
        // base64url decode
        const b64 = data.replace(/-/g, '+').replace(/_/g, '/');
        const padded = b64 + '==='.slice((b64.length + 3) % 4);
        const bin = atob(padded);
        let content = '';
        // Decode UTF-8
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        content = new TextDecoder('utf-8').decode(bytes);
        await ketcherRef.current?.setMolecule(content);
        document.title = `Ketcher Desktop — (from deep link, ${format})`;
      } catch (err) {
        alert(`Failed to open structure from link:\n${err.message || err}`);
      }
    });
  }, []);

  return (
    <div style={{ width: '100vw', height: '100vh', position: 'relative' }}>
      <Editor
        staticResourcesUrl=""
        structServiceProvider={structServiceProvider}
        errorHandler={(msg) => console.error('[Ketcher]', msg)}
        onInit={onInit}
      />
      {!ready && (
        <div style={{
          position: 'absolute', inset: 0, display: 'flex',
          alignItems: 'center', justifyContent: 'center',
          background: 'rgba(255,255,255,0.85)', fontSize: 18, color: '#333',
        }}>
          Loading Ketcher…
        </div>
      )}
    </div>
  );
}
