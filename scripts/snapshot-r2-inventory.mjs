#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { inspectRcloneInventory } from './verify-release-publication.mjs';

function usage() {
  return 'Usage: node scripts/snapshot-r2-inventory.mjs <rclone-remote> <output-json>';
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const remote = process.argv[2];
  const output = process.argv[3];
  if (!remote || !output) {
    console.error(usage());
    process.exitCode = 2;
  } else {
    const inventory = await inspectRcloneInventory(remote);
    const entries = [...inventory.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, value]) => ({ Path: key, Size: value.bytes }));
    fs.mkdirSync(path.dirname(path.resolve(output)), { recursive: true });
    fs.writeFileSync(output, `${JSON.stringify(entries)}\n`);
    console.log(`R2 inventory complete: ${entries.length} objects`);
  }
}
