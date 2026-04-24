import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { nodePolyfills } from 'vite-plugin-node-polyfills';

// Vite config for the renderer.
//
// Ketcher's internals (via ketcher-core / ketcher-standalone) reference
// Node.js built-ins — process, util, assert, buffer, stream. These have no
// native browser equivalent, so Vite needs to inject polyfills. Without
// this plugin the renderer dies at load time with "process is not defined".
//
// base: './' keeps asset paths relative so Electron can load the built
// bundle via file:// in production.
export default defineConfig({
  plugins: [
    react(),
    nodePolyfills({
      // Provide globals Ketcher expects.
      globals: {
        Buffer: true,
        global: true,
        process: true,
      },
      protocolImports: true,
    }),
  ],
  base: './',
  root: 'src',
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    // Ketcher bundles can be big; raise the warning threshold.
    chunkSizeWarningLimit: 4000,
    // ketcher-standalone ships a mix of CJS and ESM; without this, stray
    // require(...) calls land verbatim in the output bundle and crash at
    // load time with "require is not defined".
    commonjsOptions: {
      transformMixedEsModules: true,
      include: [/node_modules/],
    },
  },
  server: {
    port: 5173,
    strictPort: true,
  },
  define: {
    // Some Ketcher code checks process.env.NODE_ENV directly.
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'development'),
  },
});
