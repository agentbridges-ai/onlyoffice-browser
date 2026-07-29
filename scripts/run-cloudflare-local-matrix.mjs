#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';

const root = process.cwd();
const wranglerVersion = '4.114.0';
const seederPort = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_SEED_PORT || '8790', 10);
const workerPort = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_PORT || '8787', 10);
const persistDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-cloudflare-matrix-'));
const token = crypto.randomBytes(24).toString('hex');
const children = new Set();

function command(program, args, options = {}) {
  const child = spawn(program, args, {
    cwd: root,
    env: { ...process.env, ...options.env },
    stdio: options.stdio || 'inherit',
  });
  children.add(child);
  child.once('exit', () => children.delete(child));
  return child;
}

async function run(program, args, options = {}) {
  const child = command(program, args, options);
  const [code, signal] = await new Promise((resolve) => {
    child.once('exit', (exitCode, exitSignal) => resolve([exitCode, exitSignal]));
  });
  if (code !== 0) {
    throw new Error(`${program} ${args.join(' ')} failed (${signal || code})`);
  }
}

function wranglerArgs(config, port, extra = []) {
  return [
    '--yes',
    `wrangler@${wranglerVersion}`,
    'dev',
    '--config',
    config,
    '--local',
    '--ip',
    '127.0.0.1',
    '--port',
    String(port),
    '--persist-to',
    persistDirectory,
    '--log-level',
    'error',
    '--show-interactive-dev-session=false',
    ...extra,
  ];
}

function request({ port, method = 'GET', pathname = '/', headers = {}, body }) {
  return new Promise((resolve, reject) => {
    const outgoing = http.request(
      {
        hostname: '127.0.0.1',
        port,
        path: pathname,
        method,
        headers,
      },
      (response) => {
        const chunks = [];
        response.on('data', (chunk) => chunks.push(chunk));
        response.on('end', () =>
          resolve({
            status: response.statusCode || 0,
            headers: response.headers,
            body: Buffer.concat(chunks),
          }),
        );
      },
    );
    outgoing.once('error', reject);
    if (body) outgoing.end(body);
    else outgoing.end();
  });
}

async function waitForServer(check, label) {
  let lastError;
  for (let attempt = 0; attempt < 120; attempt += 1) {
    try {
      const result = await check();
      if (result) return;
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error(`${label} did not become ready${lastError ? `: ${lastError}` : ''}`);
}

function stop(child) {
  if (!child || child.exitCode !== null || child.signalCode !== null) return Promise.resolve();
  child.kill('SIGTERM');
  return Promise.race([
    new Promise((resolve) => child.once('exit', resolve)),
    new Promise((resolve) =>
      setTimeout(() => {
        child.kill('SIGKILL');
        resolve();
      }, 5_000),
    ),
  ]);
}

function walk(directory, prefix = '') {
  const files = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...walk(absolute, relative));
    else if (entry.isFile()) files.push({ key: relative, absolute });
  }
  return files;
}

function mimeFor(file) {
  const extension = path.extname(file).toLowerCase();
  return (
    {
      '.css': 'text/css; charset=utf-8',
      '.html': 'text/html; charset=utf-8',
      '.js': 'text/javascript; charset=utf-8',
      '.json': 'application/json; charset=utf-8',
      '.png': 'image/png',
      '.svg': 'image/svg+xml',
      '.wasm': 'application/wasm',
      '.woff': 'font/woff',
      '.woff2': 'font/woff2',
      '.otf': 'font/otf',
      '.ttf': 'font/ttf',
    }[extension] || 'application/octet-stream'
  );
}

async function uploadObject(file) {
  const bytes = fs.readFileSync(file.absolute);
  const response = await request({
    port: seederPort,
    method: 'PUT',
    pathname: '/__matrix__/object',
    headers: {
      'Content-Length': String(bytes.byteLength),
      'Content-Type': mimeFor(file.key),
      'X-Matrix-R2-Key': file.key,
      'X-Matrix-Seed-Token': token,
    },
    body: bytes,
  });
  if (response.status !== 204) {
    throw new Error(`Unable to seed ${file.key}: HTTP ${response.status} ${response.body.toString()}`);
  }
}

async function uploadFiles(files) {
  let cursor = 0;
  let completed = 0;
  const workers = Array.from({ length: 12 }, async () => {
    while (cursor < files.length) {
      const file = files[cursor++];
      await uploadObject(file);
      completed += 1;
      if (completed % 100 === 0 || completed === files.length) {
        process.stdout.write(`Seeded ${completed}/${files.length} local R2 objects\n`);
      }
    }
  });
  await Promise.all(workers);
}

async function main() {
  process.stdout.write(`Cloudflare matrix state: ${persistDirectory}\n`);
  await run('pnpm', ['build']);
  await run('node', ['scripts/hydrate-cloudflare-matrix-fonts.mjs']);
  await run('pnpm', ['release:build']);

  const releasePointer = JSON.parse(fs.readFileSync(path.join(root, '.onlyoffice-release/channels/stable.json')));
  const releaseManifest = JSON.parse(
    fs.readFileSync(path.join(root, `.onlyoffice-release/releases/${releasePointer.releaseId}/manifest.json`)),
  );
  const packageManifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json')));
  if (releaseManifest.version !== 3 || releaseManifest.releaseId !== releasePointer.releaseId) {
    throw new Error('Release v3 pointer and manifest do not match');
  }

  const seeder = command(
    'npx',
    wranglerArgs('wrangler.local-r2-seeder.jsonc', seederPort, ['--var', `MATRIX_SEED_TOKEN:${token}`]),
  );
  await waitForServer(async () => {
    const response = await request({
      port: seederPort,
      pathname: '/__matrix__/ready',
      headers: { 'X-Matrix-Seed-Token': token },
    });
    return response.status === 200;
  }, 'Local R2 seeder');

  const distFiles = walk(path.join(root, 'dist'));
  const releaseFiles = walk(path.join(root, '.onlyoffice-release'));
  await uploadFiles([...distFiles, ...releaseFiles]);
  await stop(seeder);

  const worker = command(
    'npx',
    wranglerArgs('wrangler.jsonc', workerPort, [
      '--var',
      `ASSET_VERSION:${releasePointer.releaseId}`,
      '--var',
      'LOCAL_MATRIX_MODE:1',
    ]),
  );
  await waitForServer(async () => {
    const response = await request({
      port: workerPort,
      pathname: '/channels/stable.json',
      headers: { Host: 'onlyoffice.localhost' },
    });
    return response.status === 200 && response.body.includes(releasePointer.releaseId);
  }, 'OnlyOffice production Worker');

  const hostResponse = await request({
    port: workerPort,
    method: 'HEAD',
    pathname: `/r/${releasePointer.releaseId}/office-host.html`,
    headers: { Host: 'office-editor-matrix.localhost' },
  });
  if (hostResponse.status !== 200 || hostResponse.headers['content-type'] !== 'text/html; charset=utf-8') {
    throw new Error(
      `Immutable Host preflight failed: ${hostResponse.status} ${hostResponse.headers['content-type'] || 'missing MIME'}`,
    );
  }

  const playwrightArgs = ['exec', 'playwright', 'test', '--config', 'playwright.cloudflare.config.ts'];
  if (process.env.ONLYOFFICE_CF_MATRIX_GREP) {
    playwrightArgs.push('--grep', process.env.ONLYOFFICE_CF_MATRIX_GREP);
  }
  await run('pnpm', playwrightArgs, {
    env: {
      ONLYOFFICE_CF_MATRIX_PORT: String(workerPort),
      ONLYOFFICE_CF_MATRIX_RELEASE_ID: releasePointer.releaseId,
      ONLYOFFICE_CF_MATRIX_PACKAGE_VERSION: packageManifest.version,
    },
  });
  await stop(worker);
}

async function cleanup() {
  await Promise.all([...children].map(stop));
  fs.rmSync(persistDirectory, { recursive: true, force: true });
}

main()
  .catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
    process.exitCode = 1;
  })
  .finally(cleanup);
