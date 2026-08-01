#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

function walk(root) {
  const files = [];
  for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
    const absolute = path.join(root, entry.name);
    if (entry.isDirectory()) files.push(...walk(absolute));
    else if (entry.isFile()) files.push(absolute);
  }
  return files;
}

export function verifyBuiltWorkers(root) {
  const absoluteRoot = path.resolve(root);
  const files = walk(absoluteRoot);
  const workerArtifactPattern = /^(?:conversion-worker|startup-heartbeat-worker)(?:-|\.|$)/;
  const sourceArtifacts = files.filter(
    (file) =>
      workerArtifactPattern.test(path.basename(file)) && /\.[cm]?[jt]sx?$/.test(file) && !/\.[cm]?js$/.test(file),
  );
  if (sourceArtifacts.length > 0) {
    throw new Error(
      `deployable output contains uncompiled source artifacts: ${sourceArtifacts
        .map((file) => path.relative(absoluteRoot, file))
        .join(', ')}`,
    );
  }

  for (const workerName of ['conversion-worker', 'startup-heartbeat-worker']) {
    const matches = files.filter(
      (file) => path.extname(file) === '.js' && path.basename(file).startsWith(`${workerName}-`),
    );
    if (matches.length !== 1) {
      throw new Error(`expected one executable ${workerName} JavaScript artifact, found ${matches.length}`);
    }
    const source = fs.readFileSync(matches[0], 'utf8');
    if (
      source.includes('/// <reference') ||
      /\bimport\s+type\b/.test(source) ||
      /\bas\s+DedicatedWorkerGlobalScope\b/.test(source)
    ) {
      throw new Error(`${workerName} output still contains TypeScript syntax`);
    }
  }
}

const isCli = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) {
  const root = process.argv[2] || 'dist';
  verifyBuiltWorkers(root);
  console.log(`Verified executable worker artifacts in ${path.resolve(root)}`);
}
