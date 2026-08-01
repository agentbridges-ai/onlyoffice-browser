#!/usr/bin/env node

import crypto from 'node:crypto';
import http from 'node:http';
import { chromium } from '@playwright/test';

const SEGMENT_BYTES = Buffer.allocUnsafe(2 * 1024 * 1024);
for (let index = 0; index < SEGMENT_BYTES.length; index += 1) SEGMENT_BYTES[index] = (index * 31 + 17) & 0xff;
const SEGMENT_SHA256 = crypto.createHash('sha256').update(SEGMENT_BYTES).digest('hex');
const UPDATED_SEGMENT_BYTES = Buffer.from(SEGMENT_BYTES);
UPDATED_SEGMENT_BYTES.fill(197, 768 * 1024, 832 * 1024);

function brokerHtml(port) {
  return `<!doctype html>
<meta charset="utf-8">
<script>
const CACHE_NAME = 'onlyoffice-canonical-broker-experiment-v1';
const parentOrigin = (() => {
  try {
    return new URL(document.referrer).origin;
  } catch {
    return '';
  }
})();
const allowedParent = (origin) => {
  try {
    const url = new URL(origin);
    return url.protocol === 'http:' &&
      url.port === '${port}' &&
      ['editor-a.localhost', 'editor-b.localhost'].includes(url.hostname);
  } catch {
    return false;
  }
};
const hex = (bytes) => Array.from(bytes, (value) => value.toString(16).padStart(2, '0')).join('');
const reply = (port, value, transfer = []) => port.postMessage(value, transfer);
addEventListener('message', async (event) => {
  const port = event.ports[0];
  if (!port || event.source !== parent || !allowedParent(event.origin)) return;
  try {
    const cache = await caches.open(CACHE_NAME);
    if (event.data?.type === 'INSTALL') {
      const cached = await cache.match(event.data.url);
      if (
        cached?.ok &&
        cached.headers.get('x-content-sha256') === event.data.sha256 &&
        Number(cached.headers.get('content-length')) === event.data.bytes
      ) {
        reply(port, { ok: true, bytes: event.data.bytes, sha256: event.data.sha256, reused: true });
        return;
      }
      const response = await fetch(event.data.url, { cache: 'no-store', credentials: 'omit' });
      if (!response.ok) throw new Error('network ' + response.status);
      const bytes = await response.arrayBuffer();
      const digest = hex(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)));
      if (digest !== event.data.sha256 || bytes.byteLength !== event.data.bytes) throw new Error('integrity');
      await cache.put(event.data.url, new Response(bytes, {
        headers: {
          'content-length': String(bytes.byteLength),
          'content-type': 'application/vnd.onlyoffice.browser-pack-segment',
          'x-content-sha256': digest,
        },
      }));
      reply(port, { ok: true, bytes: bytes.byteLength, sha256: digest, reused: false });
      return;
    }
    if (event.data?.type === 'GET') {
      const response = await cache.match(event.data.url);
      if (!response) throw new Error('missing');
      const bytes = await response.arrayBuffer();
      reply(port, { ok: true, bytes }, [bytes]);
      return;
    }
    if (event.data?.type === 'STATS') {
      const keys = await cache.keys();
      reply(port, { ok: true, keys: keys.map((request) => request.url) });
      return;
    }
    throw new Error('unsupported');
  } catch (error) {
    reply(port, { ok: false, error: error instanceof Error ? error.message : String(error) });
  }
});
if (allowedParent(parentOrigin)) {
  parent.postMessage({ type: 'ONLYOFFICE_BROKER_READY' }, parentOrigin);
}
</script>`;
}

function editorHtml(port) {
  return `<!doctype html>
<meta charset="utf-8">
<iframe id="broker" hidden src="http://canonical.localhost:${port}/broker.html"></iframe>
<script>
const brokerOrigin = 'http://canonical.localhost:${port}';
const brokerFrame = document.querySelector('#broker');
let readyResolve;
globalThis.__BROKER_READY__ = new Promise((resolve) => { readyResolve = resolve; });
addEventListener('message', (event) => {
  if (event.source === brokerFrame.contentWindow && event.origin === brokerOrigin &&
      event.data?.type === 'ONLYOFFICE_BROKER_READY') readyResolve();
});
globalThis.__BROKER_REQUEST__ = async (message) => {
  await globalThis.__BROKER_READY__;
  const channel = new MessageChannel();
  const result = new Promise((resolve, reject) => {
    channel.port1.onmessage = (event) => event.data?.ok ? resolve(event.data) : reject(new Error(event.data?.error));
    channel.port1.start();
  });
  brokerFrame.contentWindow.postMessage(message, brokerOrigin, [channel.port2]);
  return result;
};
globalThis.__BROKER_VERIFY__ = async (url, expected) => {
  const result = await globalThis.__BROKER_REQUEST__({ type: 'GET', url });
  const digest = Array.from(
    new Uint8Array(await crypto.subtle.digest('SHA-256', result.bytes)),
    (value) => value.toString(16).padStart(2, '0'),
  ).join('');
  return { bytes: result.bytes.byteLength, sha256: digest, matches: digest === expected };
};
</script>`;
}

async function listen(server) {
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve());
  });
  return server.address().port;
}

async function run() {
  let segmentRequests = 0;
  let port = 0;
  const server = http.createServer((request, response) => {
    const hostname = String(request.headers.host || '').split(':')[0];
    if (hostname === 'canonical.localhost' && request.url === '/broker.html') {
      response.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
      response.end(brokerHtml(port));
      return;
    }
    if (hostname === 'canonical.localhost' && request.url === `/segments/sha256/${SEGMENT_SHA256}`) {
      segmentRequests += 1;
      response.writeHead(200, {
        'content-type': 'application/vnd.onlyoffice.browser-pack-segment',
        'content-length': String(SEGMENT_BYTES.byteLength),
        'cache-control': 'no-store',
      });
      response.end(SEGMENT_BYTES);
      return;
    }
    if ((hostname === 'editor-a.localhost' || hostname === 'editor-b.localhost') && request.url === '/') {
      response.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
      response.end(editorHtml(port));
      return;
    }
    response.writeHead(404);
    response.end('not found');
  });

  port = await listen(server);
  const browser = await chromium.launch({ headless: true, args: ['--no-proxy-server'] });
  const context = await browser.newContext({ serviceWorkers: 'allow' });
  try {
    const segmentUrl = `http://canonical.localhost:${port}/segments/sha256/${SEGMENT_SHA256}`;
    const editorA = await context.newPage();
    await editorA.goto(`http://editor-a.localhost:${port}/`);
    await editorA.evaluate(() => globalThis.__BROKER_READY__);
    const installed = await editorA.evaluate(
      ({ url, sha256, bytes }) => globalThis.__BROKER_REQUEST__({ type: 'INSTALL', url, sha256, bytes }),
      { url: segmentUrl, sha256: SEGMENT_SHA256, bytes: SEGMENT_BYTES.byteLength },
    );
    const firstRead = await editorA.evaluate(({ url, sha256 }) => globalThis.__BROKER_VERIFY__(url, sha256), {
      url: segmentUrl,
      sha256: SEGMENT_SHA256,
    });
    const editorACaches = await editorA.evaluate(() => caches.keys());

    const editorB = await context.newPage();
    await editorB.goto(`http://editor-b.localhost:${port}/`);
    await editorB.evaluate(() => globalThis.__BROKER_READY__);
    const secondRead = await editorB.evaluate(({ url, sha256 }) => globalThis.__BROKER_VERIFY__(url, sha256), {
      url: segmentUrl,
      sha256: SEGMENT_SHA256,
    });
    const editorBCaches = await editorB.evaluate(() => caches.keys());

    await editorA.reload();
    await editorA.evaluate(() => globalThis.__BROKER_READY__);
    const restartRead = await editorA.evaluate(({ url, sha256 }) => globalThis.__BROKER_VERIFY__(url, sha256), {
      url: segmentUrl,
      sha256: SEGMENT_SHA256,
    });
    const stats = await editorA.evaluate(() => globalThis.__BROKER_REQUEST__({ type: 'STATS' }));

    const report = {
      version: 1,
      networkDownloads: segmentRequests,
      canonicalPersistentObjects: stats.keys.length,
      editorAPersistentCaches: editorACaches.length,
      editorBPersistentCaches: editorBCaches.length,
      installed,
      firstRead,
      secondRead,
      restartRead,
    };
    if (
      report.networkDownloads !== 1 ||
      report.canonicalPersistentObjects !== 1 ||
      report.editorAPersistentCaches !== 0 ||
      report.editorBPersistentCaches !== 0 ||
      !firstRead.matches ||
      !secondRead.matches ||
      !restartRead.matches
    ) {
      throw new Error(`canonical cache broker experiment failed\n${JSON.stringify(report, null, 2)}`);
    }
    process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  } finally {
    await context.close();
    await browser.close();
    await new Promise((resolve) => server.close(resolve));
  }
}

run().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
