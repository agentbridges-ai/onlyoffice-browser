#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadReleasePublication } from './verify-release-publication.mjs';

const REQUIRED_BROKER_ANCESTORS = [
  'https://piwork.getpi.work',
  'https://onlyoffice.getpi.work',
  'https://aries.getpi.work',
  'https://taurus.getpi.work',
  'https://gemini.getpi.work',
  'https://cancer.getpi.work',
  'https://leo.getpi.work',
  'https://virgo.getpi.work',
  'https://libra.getpi.work',
  'https://scorpio.getpi.work',
  'https://sagittarius.getpi.work',
  'https://capricorn.getpi.work',
  'https://aquarius.getpi.work',
  'https://pisces.getpi.work',
];

function fail(message) {
  throw new Error(message);
}

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function normalizeOrigin(value, label) {
  let url;
  try {
    url = new URL(value);
  } catch (error) {
    throw new Error(`${label} is invalid`, { cause: error });
  }
  if ((url.protocol !== 'https:' && url.protocol !== 'http:') || url.pathname !== '/') {
    fail(`${label} must be an HTTP(S) origin`);
  }
  return url.origin;
}

async function readJsonResponse(response, label) {
  let bytes;
  try {
    bytes = new Uint8Array(await response.arrayBuffer());
    return { bytes, value: JSON.parse(new TextDecoder().decode(bytes)) };
  } catch (error) {
    throw new Error(`${label} is not valid JSON`, { cause: error });
  }
}

function expectHeader(response, name, expected, label) {
  const actual = response.headers.get(name);
  if (typeof expected === 'string' ? actual !== expected : !expected.test(actual || '')) {
    fail(`${label} ${name} mismatch: ${actual || 'missing'}`);
  }
}

function localBytes(publication, object, start, end) {
  const file = path.join(publication.root, ...object.key.split('/'));
  const handle = fs.openSync(file, 'r');
  try {
    const output = Buffer.alloc(end - start + 1);
    const bytes = fs.readSync(handle, output, 0, output.byteLength, start);
    if (bytes !== output.byteLength) fail(`Local object ${object.key} is truncated`);
    return output;
  } finally {
    fs.closeSync(handle);
  }
}

function assertBrokerCsp(csp) {
  if (!csp) fail('resource-broker.html is missing Content-Security-Policy');
  for (const directive of [
    "default-src 'none'",
    "script-src 'self'",
    "connect-src 'self'",
    "worker-src 'self'",
    "base-uri 'none'",
    "form-action 'none'",
    'frame-ancestors',
    ...REQUIRED_BROKER_ANCESTORS,
  ]) {
    if (!csp.includes(directive)) fail(`resource-broker.html CSP is missing ${directive}`);
  }
  if (/(?:^|[\s;])\*(?:[\s;]|$)/.test(csp)) fail('resource-broker.html CSP must not contain a wildcard source');
}

async function verifyPointers(publication, origin, fetchImpl, cacheBust) {
  for (const [name, expected] of [
    ['stable-v5.json', publication.stableV5],
    ['stable.json', publication.stableV4],
  ]) {
    const response = await fetchImpl(`${origin}/channels/${name}?ci=${encodeURIComponent(cacheBust)}`, {
      cache: 'no-store',
      redirect: 'manual',
    });
    if (response.status !== 200) fail(`${name} returned ${response.status}`);
    expectHeader(response, 'cache-control', /^no-store(?:$|,)/, name);
    const pointer = await readJsonResponse(response, name);
    if (
      pointer.value?.version !== 1 ||
      pointer.value.releaseId !== publication.releaseId ||
      pointer.value.manifestUrl !== expected.manifestUrl ||
      pointer.value.manifestSha256 !== expected.manifestSha256
    ) {
      fail(`${name} does not match the local release pointer`);
    }
    const manifestResponse = await fetchImpl(
      `${origin}${pointer.value.manifestUrl}?ci=${encodeURIComponent(cacheBust)}`,
      { cache: 'no-store', redirect: 'manual' },
    );
    if (manifestResponse.status !== 200) fail(`${name} manifest returned ${manifestResponse.status}`);
    const manifestBytes = new Uint8Array(await manifestResponse.arrayBuffer());
    if (sha256(manifestBytes) !== pointer.value.manifestSha256) {
      fail(`${name} remote manifest digest mismatch`);
    }
  }
}

export async function verifyReleaseHttp(publication, options = {}) {
  const fetchImpl = options.fetchImpl || fetch;
  const canonicalOrigin = normalizeOrigin(
    options.canonicalOrigin || 'https://onlyoffice.getpi.work',
    'canonicalOrigin',
  );
  const editorOrigin = normalizeOrigin(
    options.editorOrigin || 'https://office-editor-github-actions-smoke.getpi.work',
    'editorOrigin',
  );
  const cacheBust = options.cacheBust || publication.releaseId;
  const hostUrl = `${editorOrigin}/r/${encodeURIComponent(publication.releaseId)}/office-host.html`;
  const hostResponse = await fetchImpl(hostUrl, { cache: 'no-store', redirect: 'manual' });
  if (hostResponse.status !== 200) fail(`Release Host returned ${hostResponse.status}`);
  expectHeader(hostResponse, 'content-type', /^text\/html(?:;|$)/i, 'Release Host');
  expectHeader(hostResponse, 'cache-control', /max-age=31536000.*immutable/, 'Release Host');
  expectHeader(hostResponse, 'x-onlyoffice-asset-version', publication.releaseId, 'Release Host');
  expectHeader(hostResponse, 'origin-agent-cluster', '?1', 'Release Host');

  const brokerUrl = `${canonicalOrigin}/r/${encodeURIComponent(publication.releaseId)}/resource-broker.html`;
  const brokerResponse = await fetchImpl(brokerUrl, { cache: 'no-store', redirect: 'manual' });
  if (brokerResponse.status !== 200) fail(`resource-broker.html returned ${brokerResponse.status}`);
  assertBrokerCsp(brokerResponse.headers.get('content-security-policy'));

  const object = publication.objects.find(
    (candidate) =>
      (candidate.kind === 'whole' || candidate.kind === 'fastcdc' || candidate.kind === 'package-segment') &&
      candidate.bytes >= 8,
  );
  if (!object) fail('The release has no non-empty content object for HTTP Range verification');
  const objectUrl = `${canonicalOrigin}/objects/${encodeURIComponent(publication.releaseId)}/sha256/${object.sha256}`;
  const head = await fetchImpl(objectUrl, { method: 'HEAD', cache: 'no-store', redirect: 'manual' });
  if (head.status !== 200) fail(`Content object HEAD returned ${head.status}`);
  expectHeader(head, 'content-length', String(object.bytes), 'Content object');
  expectHeader(head, 'accept-ranges', 'bytes', 'Content object');
  expectHeader(head, 'x-content-sha256', object.sha256, 'Content object');
  expectHeader(head, 'x-onlyoffice-asset-version', publication.releaseId, 'Content object');
  expectHeader(head, 'cache-control', /max-age=31536000.*immutable/, 'Content object');
  expectHeader(head, 'access-control-allow-origin', '*', 'Content object');
  expectHeader(head, 'cross-origin-resource-policy', 'cross-origin', 'Content object');

  const end = Math.min(7, object.bytes - 1);
  const range = await fetchImpl(objectUrl, {
    headers: { Range: `bytes=0-${end}` },
    cache: 'no-store',
    redirect: 'manual',
  });
  if (range.status !== 206) fail(`Content object Range returned ${range.status}`);
  expectHeader(range, 'content-length', String(end + 1), 'Content object Range');
  expectHeader(range, 'content-range', `bytes 0-${end}/${object.bytes}`, 'Content object Range');
  const actualRange = Buffer.from(await range.arrayBuffer());
  if (!actualRange.equals(localBytes(publication, object, 0, end))) {
    fail('Content object Range bytes do not match the local immutable object');
  }

  const unsatisfiable = await fetchImpl(objectUrl, {
    headers: { Range: `bytes=${object.bytes}-` },
    cache: 'no-store',
    redirect: 'manual',
  });
  if (unsatisfiable.status !== 416) fail(`Content object unsatisfiable Range returned ${unsatisfiable.status}`);
  expectHeader(unsatisfiable, 'content-range', `bytes */${object.bytes}`, 'Content object unsatisfiable Range');

  if (options.verifyPointers) {
    await verifyPointers(publication, canonicalOrigin, fetchImpl, cacheBust);
  }
  return {
    releaseId: publication.releaseId,
    object: object.key,
    pointers: Boolean(options.verifyPointers),
  };
}

function option(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const publication = loadReleasePublication(option('--release-root', '.onlyoffice-release'), {
    expectedPackageVersion: option('--expected-package-version'),
    fastCdcEvidenceMode: option('--fastcdc-evidence-mode', 'automatic'),
    fastCdcEvidencePath: option('--fastcdc-evidence'),
  });
  const result = await verifyReleaseHttp(publication, {
    canonicalOrigin: option('--canonical-origin', 'https://onlyoffice.getpi.work'),
    editorOrigin: option('--editor-origin', 'https://office-editor-github-actions-smoke.getpi.work'),
    cacheBust: option('--cache-bust', process.env.GITHUB_SHA || publication.releaseId),
    verifyPointers: process.argv.includes('--verify-pointers'),
  });
  console.log(
    `Verified production Host, Broker CSP, and content-object Range for ${result.releaseId}${
      result.pointers ? ' with stable pointers' : ''
    }`,
  );
}
