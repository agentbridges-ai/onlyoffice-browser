#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadReleasePublication, loadV5ManifestPublication } from './verify-release-publication.mjs';

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

function requestOrigin(input) {
  try {
    return new URL(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url).origin;
  } catch {
    return null;
  }
}

function isRuntimeManifestRequest(input) {
  try {
    const url = new URL(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url);
    return url.pathname.endsWith('/onlyoffice-runtime-assets.json');
  } catch {
    return false;
  }
}

function workerVersionFetch(fetchImpl, expectedWorkerVersionId, overrideWorkerVersionId, workerVersionOrigin) {
  if (!expectedWorkerVersionId) return fetchImpl;
  if (!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(expectedWorkerVersionId)) {
    fail('expectedWorkerVersionId is invalid');
  }
  if (overrideWorkerVersionId && overrideWorkerVersionId !== expectedWorkerVersionId) {
    fail('overrideWorkerVersionId must match expectedWorkerVersionId');
  }
  return async (input, init = {}) => {
    const headers = new Headers(init.headers);
    if (overrideWorkerVersionId) {
      headers.set('Cloudflare-Workers-Version-Overrides', `onlyoffice-browser-runtime="${overrideWorkerVersionId}"`);
    }
    const response = await fetchImpl(input, { ...init, headers });
    // The runtime manifest is the explicit Worker-version probe surface. The
    // version metadata binding is diagnostic and is not guaranteed to survive
    // Cloudflare's immutable asset cache or an editor-origin redirect for
    // every response. Asset identity is verified independently below.
    if ((!workerVersionOrigin || requestOrigin(input) === workerVersionOrigin) && isRuntimeManifestRequest(input)) {
      expectHeader(response, 'x-onlyoffice-worker-version', expectedWorkerVersionId, 'Worker version');
    }
    return response;
  };
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

async function verifyStableRoot(publication, origin, fetchImpl, cacheBust) {
  const manifestObject = publication.objects.find((object) => object.kind === 'manifest-v5');
  const runtimeAsset = publication.manifest?.assets?.find((asset) => asset.path === 'onlyoffice-runtime-assets.json');
  const indexAsset = publication.manifest?.assets?.find((asset) => asset.path === 'index.html');
  const serviceWorkerAsset = publication.manifest?.assets?.find((asset) => asset.path === 'sw.js');
  if (
    !manifestObject ||
    !runtimeAsset ||
    !/^[a-f0-9]{64}$/.test(runtimeAsset.sha256 || '') ||
    !indexAsset ||
    !/^[a-f0-9]{64}$/.test(indexAsset.sha256 || '') ||
    !serviceWorkerAsset ||
    !/^[a-f0-9]{64}$/.test(serviceWorkerAsset.sha256 || '')
  ) {
    fail('Release does not contain verifiable stable root shell and service-worker assets');
  }
  const pointerResponse = await fetchImpl(`${origin}/channels/stable-v5.json?ci=${encodeURIComponent(cacheBust)}`, {
    cache: 'no-store',
    redirect: 'manual',
  });
  if (pointerResponse.status !== 200) fail(`stable-v5.json returned ${pointerResponse.status}`);
  expectHeader(pointerResponse, 'cache-control', /^no-store(?:$|,)/, 'stable-v5.json');
  const pointer = await readJsonResponse(pointerResponse, 'stable-v5.json');
  if (
    pointer.value?.version !== 1 ||
    pointer.value.releaseId !== publication.releaseId ||
    pointer.value.manifestUrl !== `/releases/${publication.releaseId}/manifest.json` ||
    pointer.value.manifestSha256 !== manifestObject.sha256
  ) {
    fail('stable-v5.json does not select the expected immutable release');
  }
  const rootResponse = await fetchImpl(`${origin}/onlyoffice-runtime-assets.json?ci=${encodeURIComponent(cacheBust)}`, {
    cache: 'no-store',
    redirect: 'manual',
  });
  if (rootResponse.status !== 200) fail(`Stable root runtime manifest returned ${rootResponse.status}`);
  expectHeader(rootResponse, 'x-onlyoffice-asset-version', publication.releaseId, 'Stable root runtime manifest');
  expectHeader(
    rootResponse,
    'cache-control',
    /^public, max-age=0, must-revalidate(?:,|$)/,
    'Stable root runtime manifest',
  );
  const bytes = new Uint8Array(await rootResponse.arrayBuffer());
  if (bytes.byteLength !== runtimeAsset.bytes || sha256(bytes) !== runtimeAsset.sha256) {
    fail('Stable root runtime manifest bytes do not match the selected release');
  }

  const indexResponse = await fetchImpl(`${origin}/index.html?ci=${encodeURIComponent(cacheBust)}`, {
    cache: 'no-store',
    redirect: 'manual',
  });
  if (indexResponse.status !== 200) fail(`Stable root index.html returned ${indexResponse.status}`);
  expectHeader(indexResponse, 'content-type', /^text\/html(?:;|$)/i, 'Stable root index.html');
  expectHeader(indexResponse, 'x-onlyoffice-asset-version', publication.releaseId, 'Stable root index.html');
  expectHeader(indexResponse, 'cache-control', /^public, max-age=0, must-revalidate(?:,|$)/, 'Stable root index.html');
  const indexBytes = new Uint8Array(await indexResponse.arrayBuffer());
  if (indexBytes.byteLength !== indexAsset.bytes || sha256(indexBytes) !== indexAsset.sha256) {
    fail('Stable root index.html bytes do not match the selected release');
  }

  const serviceWorkerResponse = await fetchImpl(`${origin}/sw.js?ci=${encodeURIComponent(cacheBust)}`, {
    cache: 'no-store',
    redirect: 'manual',
  });
  if (serviceWorkerResponse.status !== 200) fail(`Stable root sw.js returned ${serviceWorkerResponse.status}`);
  expectHeader(serviceWorkerResponse, 'content-type', /^text\/javascript(?:;|$)/i, 'Stable root sw.js');
  expectHeader(serviceWorkerResponse, 'x-onlyoffice-asset-version', publication.releaseId, 'Stable root sw.js');
  expectHeader(
    serviceWorkerResponse,
    'cache-control',
    /^public, max-age=0, must-revalidate(?:,|$)/,
    'Stable root sw.js',
  );
  const serviceWorkerBytes = new Uint8Array(await serviceWorkerResponse.arrayBuffer());
  if (
    serviceWorkerBytes.byteLength !== serviceWorkerAsset.bytes ||
    sha256(serviceWorkerBytes) !== serviceWorkerAsset.sha256
  ) {
    fail('Stable root sw.js bytes do not match the selected release');
  }
}

export async function verifyReleaseHttp(publication, options = {}) {
  const expectedWorkerVersionId = options.expectedWorkerVersionId || options.workerVersionId;
  const canonicalOrigin = normalizeOrigin(
    options.canonicalOrigin || 'https://onlyoffice.getpi.work',
    'canonicalOrigin',
  );
  const workerVersionOrigin = normalizeOrigin(options.workerVersionOrigin || canonicalOrigin, 'workerVersionOrigin');
  const fetchImpl = workerVersionFetch(
    options.fetchImpl || fetch,
    expectedWorkerVersionId,
    options.workerVersionId,
    workerVersionOrigin,
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
  expectHeader(range, 'x-content-sha256', object.sha256, 'Content object Range');
  if (!options.remoteRangeVerification && !actualRange.equals(localBytes(publication, object, 0, end))) {
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
  if (options.verifyStableRoot) {
    await verifyStableRoot(publication, canonicalOrigin, fetchImpl, cacheBust);
  }
  return {
    releaseId: publication.releaseId,
    object: object.key,
    pointers: Boolean(options.verifyPointers),
    stableRoot: Boolean(options.verifyStableRoot),
  };
}

function option(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const commonOptions = {
    expectedPackageVersion: option('--expected-package-version'),
    fastCdcEvidenceMode: option('--fastcdc-evidence-mode', 'automatic'),
    fastCdcEvidencePath: option('--fastcdc-evidence'),
  };
  const v5Manifest = option('--v5-manifest');
  const publication = v5Manifest
    ? loadV5ManifestPublication(v5Manifest, {
        ...commonOptions,
        releaseId: option('--release-id'),
        expectedManifestSha256: option('--expected-manifest-sha256'),
      })
    : loadReleasePublication(option('--release-root', '.onlyoffice-release'), commonOptions);
  const result = await verifyReleaseHttp(publication, {
    canonicalOrigin: option('--canonical-origin', 'https://onlyoffice.getpi.work'),
    editorOrigin: option('--editor-origin', 'https://office-editor-github-actions-smoke.getpi.work'),
    cacheBust: option('--cache-bust', process.env.GITHUB_SHA || publication.releaseId),
    verifyPointers: process.argv.includes('--verify-pointers'),
    verifyStableRoot: process.argv.includes('--verify-stable-root'),
    workerVersionId: option('--worker-version-id'),
    expectedWorkerVersionId: option('--expected-worker-version-id'),
    workerVersionOrigin: option('--worker-version-origin'),
    remoteRangeVerification: process.argv.includes('--remote-range-verification'),
  });
  console.log(
    `Verified production Host, Broker CSP, and content-object Range for ${result.releaseId}${
      result.pointers ? ' with stable pointers' : ''
    }${result.stableRoot ? ' with the stable root alias' : ''}`,
  );
}
