import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import packageJson from './package.json' with { type: 'json' };

export default defineConfig({
  publicDir: false,
  define: {
    __ONLYOFFICE_BROWSER_VERSION__: JSON.stringify(packageJson.version),
  },
  build: {
    emptyOutDir: false,
    lib: {
      entry: resolve(__dirname, 'src/public-api.ts'),
      fileName: () => 'public-api.js',
      formats: ['es'],
    },
    outDir: resolve(__dirname, 'dist/npm'),
    sourcemap: false,
    rollupOptions: {
      output: {
        exports: 'named',
      },
    },
  },
});
