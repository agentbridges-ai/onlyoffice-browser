import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

const root = path.dirname(fileURLToPath(import.meta.url));
const outDir = process.env.ONLYOFFICE_SW_OUT_DIR || path.resolve(root, 'public');

export default defineConfig({
  publicDir: false,
  define: {
    'process.env.NODE_ENV': JSON.stringify('production'),
  },
  build: {
    outDir,
    emptyOutDir: false,
    minify: true,
    lib: {
      entry: path.resolve(root, 'src/service-worker.js'),
      name: 'OnlyOfficeBrowserServiceWorker',
      formats: ['iife'],
      fileName: () => 'sw.js',
    },
  },
});
