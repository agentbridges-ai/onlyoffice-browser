#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';

const file = process.argv[2] || 'dist/sw.js';
const source = fs.readFileSync(file, 'utf8');
const revision = crypto.createHash('sha256').update(source).digest('hex').slice(0, 24);
if (!source.includes('SW_VERSION_PLACEHOLDER')) {
  throw new Error(`${file} does not contain SW_VERSION_PLACEHOLDER`);
}
fs.writeFileSync(file, source.replaceAll('SW_VERSION_PLACEHOLDER', revision));
console.log(`Service Worker revision: ${revision}`);
