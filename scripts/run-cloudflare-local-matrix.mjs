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
const workerInternalPort = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_INTERNAL_PORT || '8788', 10);
const reusedPersistDirectory = process.env.ONLYOFFICE_CF_MATRIX_REUSE_STATE
  ? path.resolve(process.env.ONLYOFFICE_CF_MATRIX_REUSE_STATE)
  : '';
const persistDirectory =
  reusedPersistDirectory || fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-cloudflare-matrix-'));
const refreshReusedState = reusedPersistDirectory && process.env.ONLYOFFICE_CF_MATRIX_REFRESH_STATE === '1';
const token = crypto.randomBytes(24).toString('hex');
const children = new Set();
const servers = new Set();
const brokerOnly =
  process.env.ONLYOFFICE_CF_MATRIX_GREP?.includes('canonical broker') &&
  process.env.ONLYOFFICE_CF_MATRIX_FORCE_FULL !== '1';

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
    if (body && typeof body.pipe === 'function') {
      body.once('error', reject);
      body.pipe(outgoing);
    } else if (body) outgoing.end(body);
    else outgoing.end();
  });
}

function matrixHostname(authority) {
  if (typeof authority !== 'string' || !authority || authority.length > 253 || /[/\\@]/.test(authority)) return null;
  try {
    const parsed = new URL(`http://${authority}`);
    if (parsed.username || parsed.password || parsed.pathname !== '/' || parsed.search || parsed.hash) return null;
    return parsed.hostname.toLowerCase();
  } catch {
    return null;
  }
}

async function startMatrixGateway() {
  if (workerInternalPort === workerPort) {
    throw new Error('ONLYOFFICE_CF_MATRIX_INTERNAL_PORT must differ from ONLYOFFICE_CF_MATRIX_PORT');
  }
  const server = http.createServer((incoming, outgoing) => {
    const browserHostname = matrixHostname(incoming.headers.host);
    if (!browserHostname) {
      outgoing.writeHead(400, { 'content-type': 'text/plain; charset=utf-8' });
      outgoing.end('Invalid matrix Host');
      return;
    }
    const headers = {
      ...incoming.headers,
      host: 'onlyoffice.getpi.work',
      'x-onlyoffice-matrix-host': browserHostname,
    };
    delete headers['content-length'];
    if (incoming.headers['content-length']) headers['content-length'] = incoming.headers['content-length'];
    const upstream = http.request(
      {
        hostname: '127.0.0.1',
        port: workerInternalPort,
        path: incoming.url || '/',
        method: incoming.method,
        headers,
      },
      (response) => {
        outgoing.writeHead(response.statusCode || 502, response.rawHeaders);
        response.pipe(outgoing);
      },
    );
    upstream.once('error', (error) => {
      if (!outgoing.headersSent) {
        outgoing.writeHead(502, { 'content-type': 'text/plain; charset=utf-8' });
      }
      outgoing.end(`Matrix gateway upstream failed: ${error.message}`);
    });
    incoming.once('aborted', () => upstream.destroy());
    outgoing.once('close', () => {
      if (!outgoing.writableEnded) upstream.destroy();
    });
    incoming.pipe(upstream);
  });
  servers.add(server);
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(workerPort, '127.0.0.1', resolve);
  });
  return server;
}

function stopServer(server) {
  if (!server?.listening) return Promise.resolve();
  return new Promise((resolve) => server.close(resolve));
}

async function waitForServer(check, label, process) {
  let lastError;
  for (let attempt = 0; attempt < 360; attempt += 1) {
    if (process && (process.exitCode !== null || process.signalCode !== null)) {
      throw new Error(
        `${label} exited before becoming ready (${process.signalCode || process.exitCode || 'unknown status'})`,
      );
    }
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
    else if (entry.isFile()) files.push({ key: relative, absolute, bytes: fs.statSync(absolute).size });
  }
  return files;
}

function writeFixture(rootDirectory, key, bytes) {
  const absolute = path.join(rootDirectory, key);
  fs.mkdirSync(path.dirname(absolute), { recursive: true });
  fs.writeFileSync(absolute, bytes);
  return { key, absolute, bytes: fs.statSync(absolute).size };
}

function prepareBrokerOnlyRelease() {
  const sourceDirectory = path.join(persistDirectory, 'broker-only-source');
  const segment = Buffer.allocUnsafe(2 * 1024 * 1024);
  for (let index = 0; index < segment.length; index += 1) segment[index] = (index * 31 + 17) & 0xff;
  const sha256 = crypto.createHash('sha256').update(segment).digest('hex');
  const releaseId = `broker-experiment-${sha256.slice(0, 16)}`;
  const changedSegment = Buffer.allocUnsafe(1024 * 1024);
  for (let index = 0; index < changedSegment.length; index += 1) changedSegment[index] = (index * 17 + 29) & 0xff;
  const changedSha256 = crypto.createHash('sha256').update(changedSegment).digest('hex');
  const nextReleaseId = `broker-experiment-next-${changedSha256.slice(0, 16)}`;
  const packageManifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json')));
  const matrixFixtures = walk(path.join(root, 'test/cloudflare-e2e/fixtures')).map((file) => {
    const digest = crypto.createHash('sha256').update(fs.readFileSync(file.absolute)).digest('hex');
    return { ...file, digest, key: `__matrix__/${file.key}` };
  });
  const pwaShellTopLevel = new Set([
    'sw.js',
    'index.html',
    'editor-shell-prime.html',
    'resource-broker.html',
    'resource-installer.html',
    'plugins.json',
    'themes.json',
  ]);
  const pwaShellAssetPrefixes = [
    'assets/main-',
    'assets/editor-shell-cache-',
    'assets/editorShellPrime-',
    'assets/office-origin-pool-',
    'assets/resource-broker-frame-client-',
    'assets/resourceBroker-',
    'assets/resource-broker-protocol-',
    'assets/resourceInstaller-',
    'assets/resource-installer-frame-protocol-',
    'assets/base-',
  ];
  const pwaShellFiles = walk(path.join(root, 'dist')).filter(
    (file) => pwaShellTopLevel.has(file.key) || pwaShellAssetPrefixes.some((prefix) => file.key.startsWith(prefix)),
  );
  const releaseManifest = {
    version: 4,
    releaseId,
    packageVersion: packageManifest.version,
    assets: matrixFixtures.map((file) => ({
      path: file.key,
      bytes: file.bytes,
      mime: mimeFor(file.key),
      sha256: file.digest,
    })),
    package: {
      format: 'onlyoffice-pack-v1',
      path: 'office-resources.oobpack',
      bytes: segment.byteLength,
      sha256,
      segments: [{ id: 'segment-0001', offset: 0, bytes: segment.byteLength, sha256 }],
    },
  };
  const nextReleaseManifest = {
    ...releaseManifest,
    releaseId: nextReleaseId,
    package: {
      ...releaseManifest.package,
      bytes: segment.byteLength + changedSegment.byteLength,
      sha256: crypto.createHash('sha256').update(segment).update(changedSegment).digest('hex'),
      segments: [
        ...releaseManifest.package.segments,
        {
          id: 'segment-0002',
          offset: segment.byteLength,
          bytes: changedSegment.byteLength,
          sha256: changedSha256,
        },
      ],
    },
  };
  const files = [
    writeFixture(sourceDirectory, 'channels/stable.json', JSON.stringify({ version: 1, releaseId }, null, 2)),
    writeFixture(sourceDirectory, `releases/${releaseId}/manifest.json`, JSON.stringify(releaseManifest, null, 2)),
    writeFixture(sourceDirectory, `segments/sha256/${sha256}`, segment),
    writeFixture(
      sourceDirectory,
      `releases/${nextReleaseId}/manifest.json`,
      JSON.stringify(nextReleaseManifest, null, 2),
    ),
    writeFixture(sourceDirectory, `segments/sha256/${changedSha256}`, changedSegment),
    ...pwaShellFiles,
    ...matrixFixtures,
    ...matrixFixtures.map((file) => ({ ...file, key: `blobs/sha256/${file.digest}` })),
  ];
  return {
    files,
    releasePointer: { version: 1, releaseId },
    releaseManifest,
    packageManifest,
    nextReleaseId,
  };
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
      '.oobpack': 'application/vnd.onlyoffice.browser-pack',
      '.ttf': 'font/ttf',
    }[extension] || 'application/octet-stream'
  );
}

async function prepareFullReleaseVariants(baseManifest, packageManifest) {
  const changedPath = 'onlyoffice-icon.svg';
  if (!baseManifest.assets?.some((asset) => asset.path === changedPath)) {
    throw new Error(`Full-v5 matrix release is missing mutation sentinel ${changedPath}`);
  }
  const buildVariant = async (label, marker) => {
    const variantRoot = path.join(persistDirectory, `${label}-dist`);
    const variantOutput = path.join(persistDirectory, `${label}-release`);
    fs.rmSync(variantRoot, { recursive: true, force: true });
    fs.rmSync(variantOutput, { recursive: true, force: true });
    fs.cpSync(path.join(root, 'dist'), variantRoot, {
      recursive: true,
      mode: fs.constants.COPYFILE_FICLONE,
    });
    fs.appendFileSync(path.join(variantRoot, changedPath), `\n<!-- ${marker} -->\n`);
    await run('node', [
      'scripts/build-release-manifest.mjs',
      '--root',
      variantRoot,
      '--output',
      variantOutput,
      '--package-version',
      packageManifest.version,
    ]);
    const pointer = JSON.parse(fs.readFileSync(path.join(variantOutput, 'channels/stable-v5.json')));
    const manifest = JSON.parse(
      fs.readFileSync(path.join(variantOutput, `releases/${pointer.releaseId}/manifest.json`)),
    );
    const changed = manifest.assets.find((asset) => asset.path === changedPath);
    if (!changed || changed.sha256 === baseManifest.assets.find((asset) => asset.path === changedPath)?.sha256) {
      throw new Error(`Full-v5 matrix ${label} release did not mutate ${changedPath}`);
    }
    return {
      changed,
      manifest,
      output: variantOutput,
      pointer,
    };
  };

  const next = await buildVariant('next', 'onlyoffice full-v5 incremental release B');
  const failed = await buildVariant('failed', 'onlyoffice full-v5 intentionally incomplete release');
  const variantFiles = [
    ...walk(next.output),
    ...walk(failed.output).filter((file) => file.key !== `blobs/sha256/${failed.changed.sha256}`),
  ].filter((file) => !file.key.startsWith('channels/') && !file.key.startsWith('packages/sha256/'));
  return {
    files: variantFiles,
    next,
    failed,
    changedPath,
  };
}

async function uploadObject(file) {
  if (/^(?:blobs|segments)\/sha256\/[a-f0-9]{64}$|^packages\/sha256\/[a-f0-9]{64}\.oobpack$/.test(file.key)) {
    const existing = await request({
      port: seederPort,
      method: 'HEAD',
      pathname: '/__matrix__/object',
      headers: {
        'X-Matrix-R2-Key': file.key,
        'X-Matrix-Seed-Token': token,
      },
    });
    if (existing.status === 200 && Number(existing.headers['content-length']) === file.bytes) return;
  }
  let lastFailure = '';
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      const response = await request({
        port: seederPort,
        method: 'PUT',
        pathname: '/__matrix__/object',
        headers: {
          'Content-Length': String(file.bytes),
          'Content-Type': mimeFor(file.key),
          'X-Matrix-R2-Key': file.key,
          'X-Matrix-Seed-Token': token,
        },
        body: fs.createReadStream(file.absolute),
      });
      if (response.status === 204) return;
      lastFailure = `HTTP ${response.status} ${response.body.toString()}`;
    } catch (error) {
      lastFailure = error instanceof Error ? error.message : String(error);
    }
    if (attempt < 2) await new Promise((resolve) => setTimeout(resolve, [100, 500][attempt]));
  }
  throw new Error(`Unable to seed ${file.key} after 3 attempts: ${lastFailure}`);
}

async function uploadMultipartObject(file) {
  if (/^(?:blobs|segments)\/sha256\/[a-f0-9]{64}$|^packages\/sha256\/[a-f0-9]{64}\.oobpack$/.test(file.key)) {
    const existing = await request({
      port: seederPort,
      method: 'HEAD',
      pathname: '/__matrix__/object',
      headers: {
        'X-Matrix-R2-Key': file.key,
        'X-Matrix-Seed-Token': token,
      },
    });
    if (existing.status === 200 && Number(existing.headers['content-length']) === file.bytes) return;
  }
  const commonHeaders = {
    'Content-Type': mimeFor(file.key),
    'X-Matrix-R2-Key': file.key,
    'X-Matrix-Seed-Token': token,
  };
  const started = await request({
    port: seederPort,
    method: 'POST',
    pathname: '/__matrix__/multipart/start',
    headers: commonHeaders,
  });
  if (started.status !== 200) {
    throw new Error(`Unable to start multipart seed for ${file.key}: HTTP ${started.status} ${started.body}`);
  }
  const uploadId = JSON.parse(started.body.toString()).uploadId;
  if (typeof uploadId !== 'string' || !uploadId) {
    throw new Error(`Local R2 returned an invalid multipart upload ID for ${file.key}`);
  }
  const partBytes = 8 * 1024 * 1024;
  const parts = [];
  try {
    for (let offset = 0, partNumber = 1; offset < file.bytes; offset += partBytes, partNumber += 1) {
      const bytes = Math.min(partBytes, file.bytes - offset);
      let uploaded = null;
      let lastFailure = '';
      for (let attempt = 0; attempt < 3 && !uploaded; attempt += 1) {
        try {
          const response = await request({
            port: seederPort,
            method: 'PUT',
            pathname: '/__matrix__/multipart/part',
            headers: {
              ...commonHeaders,
              'Content-Length': String(bytes),
              'X-Matrix-R2-Upload-Id': uploadId,
              'X-Matrix-R2-Part-Number': String(partNumber),
            },
            body: fs.createReadStream(file.absolute, { start: offset, end: offset + bytes - 1 }),
          });
          if (response.status === 200) {
            uploaded = JSON.parse(response.body.toString());
            break;
          }
          lastFailure = `HTTP ${response.status} ${response.body.toString()}`;
        } catch (error) {
          lastFailure = error instanceof Error ? error.message : String(error);
        }
        if (attempt < 2) await new Promise((resolve) => setTimeout(resolve, [100, 500][attempt]));
      }
      if (!uploaded || uploaded.partNumber !== partNumber || typeof uploaded.etag !== 'string' || !uploaded.etag) {
        throw new Error(`Unable to seed multipart part ${partNumber} for ${file.key}: ${lastFailure}`);
      }
      parts.push({ partNumber, etag: uploaded.etag });
    }
    const completionBody = Buffer.from(JSON.stringify({ parts }));
    const completed = await request({
      port: seederPort,
      method: 'POST',
      pathname: '/__matrix__/multipart/complete',
      headers: {
        ...commonHeaders,
        'Content-Type': 'application/json',
        'Content-Length': String(completionBody.byteLength),
        'X-Matrix-R2-Upload-Id': uploadId,
      },
      body: completionBody,
    });
    if (completed.status !== 204) {
      throw new Error(`Unable to complete multipart seed for ${file.key}: HTTP ${completed.status} ${completed.body}`);
    }
  } catch (error) {
    await request({
      port: seederPort,
      method: 'DELETE',
      pathname: '/__matrix__/multipart',
      headers: {
        ...commonHeaders,
        'X-Matrix-R2-Upload-Id': uploadId,
      },
    }).catch(() => undefined);
    throw error;
  }
}

async function uploadFiles(files) {
  let completed = 0;
  const largeThreshold = 8 * 1024 * 1024;
  const smallFiles = files.filter((file) => file.bytes <= largeThreshold);
  const largeFiles = files.filter((file) => file.bytes > largeThreshold);
  let cursor = 0;
  const markCompleted = () => {
    completed += 1;
    if (completed % 100 === 0 || completed === files.length) {
      process.stdout.write(`Seeded ${completed}/${files.length} local R2 objects\n`);
    }
  };
  const workers = Array.from({ length: 8 }, async () => {
    while (cursor < smallFiles.length) {
      const file = smallFiles[cursor++];
      await uploadObject(file);
      markCompleted();
    }
  });
  await Promise.all(workers);
  for (const file of largeFiles) {
    await uploadMultipartObject(file);
    markCompleted();
  }
}

async function main() {
  process.stdout.write(`Cloudflare matrix state: ${persistDirectory}\n`);
  let releasePointer;
  let releaseManifest;
  let packageManifest;
  let nextReleaseId;
  let nextManifestSha256;
  let nextChangedPath;
  let nextChangedSha256;
  let nextChangedBytes;
  let failedReleaseId;
  let failedManifestSha256;
  let failedChangedSha256;
  let files;
  if (brokerOnly) {
    ({ files, releasePointer, releaseManifest, packageManifest, nextReleaseId } = prepareBrokerOnlyRelease());
    process.stdout.write(`Using minimal broker release ${releasePointer.releaseId} (${files.length} R2 objects)\n`);
  } else {
    if (!reusedPersistDirectory || refreshReusedState) {
      await run('pnpm', ['build']);
      await run('node', ['scripts/hydrate-cloudflare-matrix-fonts.mjs']);
      await run('pnpm', ['release:build']);
    } else if (!fs.existsSync(reusedPersistDirectory)) {
      throw new Error(`Requested Cloudflare matrix state does not exist: ${reusedPersistDirectory}`);
    }
    releasePointer = JSON.parse(fs.readFileSync(path.join(root, '.onlyoffice-release/channels/stable-v5.json')));
    releaseManifest = JSON.parse(
      fs.readFileSync(path.join(root, `.onlyoffice-release/releases/${releasePointer.releaseId}/manifest.json`)),
    );
    packageManifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json')));
    const variants =
      reusedPersistDirectory && !refreshReusedState
        ? (() => {
            const readVariant = (label) => {
              const output = path.join(persistDirectory, `${label}-release`);
              const pointer = JSON.parse(fs.readFileSync(path.join(output, 'channels/stable-v5.json')));
              const manifest = JSON.parse(
                fs.readFileSync(path.join(output, `releases/${pointer.releaseId}/manifest.json`)),
              );
              const changed = manifest.assets.find((asset) => {
                const base = releaseManifest.assets.find((candidate) => candidate.path === asset.path);
                return base && base.sha256 !== asset.sha256;
              });
              if (!changed) throw new Error(`Reused ${label} release has no changed asset`);
              return { changed, manifest, output, pointer };
            };
            const next = readVariant('next');
            const failed = readVariant('failed');
            return { files: [], next, failed, changedPath: next.changed.path };
          })()
        : await prepareFullReleaseVariants(releaseManifest, packageManifest);
    nextReleaseId = variants.next.pointer.releaseId;
    nextManifestSha256 = variants.next.pointer.manifestSha256;
    nextChangedPath = variants.changedPath;
    nextChangedSha256 = variants.next.changed.sha256;
    nextChangedBytes = variants.next.changed.bytes;
    failedReleaseId = variants.failed.pointer.releaseId;
    failedManifestSha256 = variants.failed.pointer.manifestSha256;
    failedChangedSha256 = variants.failed.changed.sha256;
    files =
      reusedPersistDirectory && !refreshReusedState
        ? []
        : [
            ...walk(path.join(root, 'dist')),
            ...walk(path.join(root, '.onlyoffice-release')),
            ...walk(path.join(root, 'test/cloudflare-e2e/fixtures')).map((file) => ({
              ...file,
              key: `__matrix__/${file.key}`,
            })),
            ...variants.files,
          ];
  }
  files = [...new Map(files.map((file) => [file.key, file])).values()];
  const routeOnly = releaseManifest.version === 5;
  const validLegacyPack = releaseManifest.version === 4 && releaseManifest.package?.format === 'onlyoffice-pack-v1';
  const validContentRelease =
    releaseManifest.version === 5 && Array.isArray(releaseManifest.assets) && Array.isArray(releaseManifest.chunks);
  if (releaseManifest.releaseId !== releasePointer.releaseId || (!validLegacyPack && !validContentRelease)) {
    throw new Error('Release pointer and immutable manifest do not match a supported v4/v5 release');
  }

  if (files.length > 0) {
    const seeder = command(
      'npx',
      wranglerArgs('wrangler.local-r2-seeder.jsonc', seederPort, ['--var', `MATRIX_SEED_TOKEN:${token}`]),
    );
    await waitForServer(
      async () => {
        const response = await request({
          port: seederPort,
          pathname: '/__matrix__/ready',
          headers: { 'X-Matrix-Seed-Token': token },
        });
        return response.status === 200;
      },
      'Local R2 seeder',
      seeder,
    );

    await uploadFiles(files);
    await stop(seeder);
  } else {
    process.stdout.write(`Reusing seeded local R2 state from ${persistDirectory}\n`);
  }

  const worker = command(
    'npx',
    wranglerArgs('wrangler.jsonc', workerInternalPort, [
      '--var',
      `ASSET_VERSION:${releasePointer.releaseId}`,
      '--var',
      'LOCAL_MATRIX_MODE:1',
      '--var',
      `LOCAL_MATRIX_PORT:${workerPort}`,
      '--var',
      `LOCAL_MATRIX_CONTROL_TOKEN:${token}`,
      ...(process.env.ONLYOFFICE_CF_MATRIX_R2_DELAY_MS
        ? ['--var', `LOCAL_MATRIX_R2_DELAY_MS:${process.env.ONLYOFFICE_CF_MATRIX_R2_DELAY_MS}`]
        : []),
      ...(process.env.ONLYOFFICE_CF_MATRIX_STALL_TEST === '1' || process.env.ONLYOFFICE_CF_MATRIX_STALL_SEGMENT
        ? [
            '--var',
            `LOCAL_MATRIX_R2_STALL_KEY:${
              process.env.ONLYOFFICE_CF_MATRIX_STALL_SEGMENT || releaseManifest.package.segments[0].sha256
            }`,
            '--var',
            `LOCAL_MATRIX_R2_STALL_AFTER_BYTES:${process.env.ONLYOFFICE_CF_MATRIX_STALL_AFTER_BYTES || '1'}`,
            '--var',
            `LOCAL_MATRIX_R2_STALL_MS:${process.env.ONLYOFFICE_CF_MATRIX_STALL_MS || '5000'}`,
            '--var',
            `LOCAL_MATRIX_R2_STALL_ONCE:${process.env.ONLYOFFICE_CF_MATRIX_STALL_ONCE || '1'}`,
          ]
        : []),
    ]),
  );
  await waitForServer(
    async () => {
      const response = await request({
        port: workerInternalPort,
        pathname: routeOnly ? '/channels/stable-v5.json' : '/channels/stable.json',
        headers: {
          Host: 'onlyoffice.getpi.work',
          'X-OnlyOffice-Matrix-Host': 'onlyoffice.localhost',
        },
      });
      return response.status === 200;
    },
    'OnlyOffice production Worker',
    worker,
  );
  if (routeOnly) {
    const pointerBody = Buffer.from(`${JSON.stringify(releasePointer, null, 2)}\n`);
    const resetResponse = await request({
      port: workerInternalPort,
      method: 'POST',
      pathname: '/__matrix__/stable-v5',
      headers: {
        Host: 'onlyoffice.getpi.work',
        'X-OnlyOffice-Matrix-Host': 'onlyoffice.localhost',
        'X-OnlyOffice-Matrix-Control-Token': token,
        'Content-Type': 'application/json',
        'Content-Length': String(pointerBody.byteLength),
      },
      body: pointerBody,
    });
    if (resetResponse.status !== 200) {
      throw new Error(`Could not reset the local stable-v5 pointer: ${resetResponse.status} ${resetResponse.body}`);
    }
  }
  await startMatrixGateway();

  if (!brokerOnly) {
    const hostResponse = await request({
      port: workerPort,
      method: 'HEAD',
      pathname: `/r/${releasePointer.releaseId}/office-host.html`,
      headers: { Host: 'aries.localhost' },
    });
    if (hostResponse.status !== 200 || hostResponse.headers['content-type'] !== 'text/html; charset=utf-8') {
      throw new Error(
        `Immutable Host preflight failed: ${hostResponse.status} ${
          hostResponse.headers['content-type'] || 'missing MIME'
        }`,
      );
    }
    const packResponse = await request({
      port: workerPort,
      pathname: `/p/${releasePointer.releaseId}/office-resources.oobpack`,
      headers: {
        Host: 'onlyoffice.localhost',
        Range: 'bytes=0-7',
      },
    });
    if (
      packResponse.status !== 206 ||
      packResponse.body.toString() !== 'OOBPACK1' ||
      packResponse.headers['content-type'] !== 'application/vnd.onlyoffice.browser-pack'
    ) {
      throw new Error(`Office Pack preflight failed: ${packResponse.status} ${packResponse.body.toString('hex')}`);
    }
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
      ...(nextReleaseId ? { ONLYOFFICE_CF_MATRIX_NEXT_RELEASE_ID: nextReleaseId } : {}),
      ...(nextManifestSha256 ? { ONLYOFFICE_CF_MATRIX_NEXT_MANIFEST_SHA256: nextManifestSha256 } : {}),
      ...(nextChangedPath ? { ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_PATH: nextChangedPath } : {}),
      ...(nextChangedSha256 ? { ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_SHA256: nextChangedSha256 } : {}),
      ...(nextChangedBytes ? { ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_BYTES: String(nextChangedBytes) } : {}),
      ...(failedReleaseId ? { ONLYOFFICE_CF_MATRIX_FAILED_RELEASE_ID: failedReleaseId } : {}),
      ...(failedManifestSha256 ? { ONLYOFFICE_CF_MATRIX_FAILED_MANIFEST_SHA256: failedManifestSha256 } : {}),
      ...(failedChangedSha256 ? { ONLYOFFICE_CF_MATRIX_FAILED_CHANGED_SHA256: failedChangedSha256 } : {}),
      ONLYOFFICE_CF_MATRIX_CONTROL_TOKEN: token,
      ONLYOFFICE_CF_MATRIX_MODE: brokerOnly ? 'synthetic-broker' : 'full-v5',
      ...(process.platform === 'darwin' ? { ONLYOFFICE_CF_MATRIX_BROWSER_CHANNEL: 'chrome' } : {}),
      ...(process.env.ONLYOFFICE_CF_MATRIX_STALL_TEST === '1'
        ? {
            ONLYOFFICE_CF_MATRIX_STALL_SEGMENT:
              process.env.ONLYOFFICE_CF_MATRIX_STALL_SEGMENT || releaseManifest.package.segments[0].sha256,
          }
        : {}),
    },
  });
  await stop(worker);
}

async function cleanup(removeState) {
  await Promise.all([...servers].map(stopServer));
  await Promise.all([...children].map(stop));
  if (removeState && !reusedPersistDirectory) {
    fs.rmSync(persistDirectory, { recursive: true, force: true });
  } else {
    process.stderr.write(`Cloudflare matrix state retained for diagnosis: ${persistDirectory}\n`);
  }
}

let succeeded = false;
main()
  .then(() => {
    succeeded = true;
  })
  .catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
    process.exitCode = 1;
  })
  .finally(() => cleanup(succeeded && process.env.ONLYOFFICE_CF_MATRIX_KEEP_STATE !== '1'));
