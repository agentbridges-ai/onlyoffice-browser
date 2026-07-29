#!/usr/bin/env node

import path from 'node:path';
import { injectManifest } from 'workbox-build';

const root = process.cwd();
const { count, size, warnings } = await injectManifest({
  swSrc: path.join(root, '.onlyoffice-sw', 'sw.js'),
  swDest: path.join(root, 'dist', 'sw.js'),
  globDirectory: path.join(root, 'dist'),
  globPatterns: ['index.html', 'plugins.json', 'themes.json', 'assets/main-*.js', 'assets/base-*.css'],
  maximumFileSizeToCacheInBytes: 2 * 1024 * 1024,
  dontCacheBustURLsMatching: /-[a-zA-Z0-9_-]{8,}\.(?:css|js)$/,
});

for (const warning of warnings) console.warn(warning);
console.log(`Injected ${count} PWA shell assets (${size} bytes) into dist/sw.js`);
