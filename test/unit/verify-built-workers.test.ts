import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { verifyBuiltWorkers } from '../../scripts/verify-built-workers.mjs';

const roots: string[] = [];

function fixture(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-worker-build-'));
  roots.push(root);
  fs.mkdirSync(path.join(root, 'wasm', 'x2t'), { recursive: true });
  fs.writeFileSync(path.join(root, 'wasm', 'x2t', 'conversion-worker-a.js'), 'self.onmessage=()=>{};');
  fs.writeFileSync(path.join(root, 'wasm', 'x2t', 'startup-heartbeat-worker-a.js'), 'self.onmessage=()=>{};');
  return root;
}

afterEach(() => {
  for (const root of roots.splice(0)) fs.rmSync(root, { recursive: true, force: true });
});

describe('verifyBuiltWorkers', () => {
  it('accepts exactly one executable JavaScript artifact for each Worker', () => {
    expect(() => verifyBuiltWorkers(fixture())).not.toThrow();
  });

  it('rejects copied TypeScript source even when a JavaScript worker also exists', () => {
    const root = fixture();
    fs.mkdirSync(path.join(root, 'assets'), { recursive: true });
    fs.writeFileSync(path.join(root, 'assets', 'conversion-worker.ts'), 'import type { X } from "./x";', {
      flag: 'w',
    });
    expect(() => verifyBuiltWorkers(root)).toThrow(/uncompiled source artifacts/);
  });

  it('allows declaration files that are intentional third-party runtime assets', () => {
    const root = fixture();
    const vendor = path.join(root, 'web-apps', 'vendor', 'monaco');
    fs.mkdirSync(vendor, { recursive: true });
    fs.writeFileSync(path.join(vendor, 'monaco.d.ts'), 'declare const monaco: unknown;');
    expect(() => verifyBuiltWorkers(root)).not.toThrow();
  });
});
