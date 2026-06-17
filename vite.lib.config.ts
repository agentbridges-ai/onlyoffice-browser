import { resolve } from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  publicDir: false,
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
