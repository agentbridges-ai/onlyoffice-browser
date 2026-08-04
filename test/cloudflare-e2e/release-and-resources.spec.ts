import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium, expect, type Browser, type BrowserContext, type Frame, type Page, test } from '@playwright/test';
import { planCanonicalPackageTransfer, type CanonicalPackageTransport } from '../../src/lib/canonical-resource-store';
import { releaseManifestV5ToContentModel } from '../../src/lib/release-content-manifest';
import { planReleaseContent } from '../../src/lib/release-content-model';
import type { ReleaseManifestV5 } from '../../src/lib/release-resources';

const releaseId = process.env.ONLYOFFICE_CF_MATRIX_RELEASE_ID;
const nextReleaseId = process.env.ONLYOFFICE_CF_MATRIX_NEXT_RELEASE_ID;
const nextManifestSha256 = process.env.ONLYOFFICE_CF_MATRIX_NEXT_MANIFEST_SHA256;
const nextChangedPath = process.env.ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_PATH;
const nextChangedSha256 = process.env.ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_SHA256;
const nextChangedBytes = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_NEXT_CHANGED_BYTES || '0', 10);
const failedReleaseId = process.env.ONLYOFFICE_CF_MATRIX_FAILED_RELEASE_ID;
const failedManifestSha256 = process.env.ONLYOFFICE_CF_MATRIX_FAILED_MANIFEST_SHA256;
const failedChangedSha256 = process.env.ONLYOFFICE_CF_MATRIX_FAILED_CHANGED_SHA256;
const packageVersion = process.env.ONLYOFFICE_CF_MATRIX_PACKAGE_VERSION;
const matrixControlToken = process.env.ONLYOFFICE_CF_MATRIX_CONTROL_TOKEN;
const matrixMode = process.env.ONLYOFFICE_CF_MATRIX_MODE || 'full-v5';
const persistentBrowserChannel = process.env.ONLYOFFICE_CF_MATRIX_BROWSER_CHANNEL;
const stallSegment = process.env.ONLYOFFICE_CF_MATRIX_STALL_SEGMENT;
const port = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_PORT || '8787', 10);
const canonicalOrigin = `http://onlyoffice.localhost:${port}`;
const editorOrigin = `http://aries.localhost:${port}`;
const reusableEditorOrigin = `http://host-aries.office.localhost:${port}`;
const productionEditorOrigins = ['aries', 'taurus', 'gemini'].map(
  (slot) => `http://host-${slot}.office.localhost:${port}`,
);

type ReleaseManifest = {
  version: number;
  releaseId: string;
  packageVersion: string;
  contentProtocol?: {
    version: number;
    digest: string;
    cacheKeyFormat: string;
    storageSetSha256: string;
    fastcdcPolicyId: string;
  };
  package: {
    path: string;
    bytes: number;
    sha256: string;
    segments: Array<{ id: string; offset: number; bytes: number; sha256: string }>;
  };
  assets: Array<{
    path: string;
    mime: string;
    bytes: number;
    sha256: string;
    packageOffset: number;
    representations?: {
      whole?: { bytes: number; sha256: string };
      fastcdc?: { chunks: Array<{ offset: number; bytes: number; sha256: string }> };
    };
  }>;
};

type FailurePhase = 'normal' | 'offline-recovery' | 'online-update';
type BrowserFailure = { source: string; message: string; url?: string; phase?: FailurePhase };
type ContentCounter = {
  workerRequests: number;
  cacheHits: number;
  r2Heads: number;
  r2Gets: number;
  declaredBytes: number;
  actualBytes: number;
  r2Bytes: number;
  completed: number;
  aborted: number;
  failed: number;
  stalled: number;
  statuses: Record<string, number>;
};

const emptyContentCounter = (): ContentCounter => ({
  workerRequests: 0,
  cacheHits: 0,
  r2Heads: 0,
  r2Gets: 0,
  declaredBytes: 0,
  actualBytes: 0,
  r2Bytes: 0,
  completed: 0,
  aborted: 0,
  failed: 0,
  stalled: 0,
  statuses: {},
});

async function readBroadcastCounts(pages: Page[]): Promise<number[]> {
  return Promise.all(
    pages.map((page) =>
      page.evaluate(
        () =>
          (globalThis as typeof globalThis & { __ONLYOFFICE_MATRIX_BROADCAST_COUNT__: number })
            .__ONLYOFFICE_MATRIX_BROADCAST_COUNT__,
      ),
    ),
  );
}

async function waitForStableBroadcastCounts(pages: Page[], stableFor = 2_500, timeout = 15_000): Promise<number[]> {
  const startedAt = Date.now();
  let lastCounts = await readBroadcastCounts(pages);
  let stableSince = Date.now();
  while (Date.now() - startedAt < timeout) {
    await pages[0].waitForTimeout(250);
    const currentCounts = await readBroadcastCounts(pages);
    if (currentCounts.some((count, index) => count !== lastCounts[index])) {
      lastCounts = currentCounts;
      stableSince = Date.now();
    } else if (Date.now() - stableSince >= stableFor) {
      return currentCounts;
    }
  }
  throw new Error(`Resource broadcasts did not settle within ${timeout} ms`);
}

function collectFailures(page: Page, failures: BrowserFailure[], phase?: () => FailurePhase) {
  page.on('pageerror', (error) =>
    failures.push({ source: 'pageerror', message: error.message, ...(phase ? { phase: phase() } : {}) }),
  );
  page.on('console', (message) => {
    const text = message.text();
    if (
      message.type() === 'error' ||
      (message.type() === 'warning' &&
        (text.includes('[onlyoffice-browser]') ||
          /PressResponder|pressable child|unhandled (?:promise )?rejection|React has detected|Warning:.*React/i.test(
            text,
          )))
    ) {
      const url = message.location().url;
      failures.push({
        source: 'console',
        message: text,
        ...(url ? { url } : {}),
        ...(phase ? { phase: phase() } : {}),
      });
    }
  });
}

function isExpectedOfflineStablePointerFailure(failure: BrowserFailure): boolean {
  if (
    failure.phase !== 'offline-recovery' ||
    failure.source !== 'console' ||
    !failure.url ||
    !/^Failed to load resource: the server responded with a status of 503(?: \([^)]*\))?$/.test(failure.message)
  ) {
    return false;
  }
  const actual = new URL(failure.url);
  return actual.origin === canonicalOrigin && actual.pathname === '/channels/stable-v5.json' && !actual.search;
}

async function freshContext(browser: Browser): Promise<BrowserContext> {
  const context = await browser.newContext({
    locale: 'en-US',
    serviceWorkers: 'allow',
    acceptDownloads: true,
  });
  await configureMatrixContext(context);
  return context;
}

async function configureMatrixContext(context: BrowserContext, options: { forceDownloadSave?: boolean } = {}) {
  const forceDownloadSave = options.forceDownloadSave === true;
  await context.addInitScript(() => {
    try {
      localStorage.setItem('onlyoffice-browser.locale', 'en-US');
    } catch {
      // The init script also runs in the inaccessible initial about:blank.
    }
    const OriginalBroadcastChannel = globalThis.BroadcastChannel;
    let postCount = 0;
    Object.defineProperty(globalThis, '__ONLYOFFICE_MATRIX_BROADCAST_COUNT__', {
      configurable: false,
      get: () => postCount,
    });
    globalThis.BroadcastChannel = class extends OriginalBroadcastChannel {
      postMessage(message: unknown) {
        if (this.name === 'onlyoffice-resources-v3') postCount += 1;
        return super.postMessage(message);
      }
    };
  });
  if (forceDownloadSave) {
    await context.addInitScript(() => {
      Object.defineProperty(globalThis, 'showSaveFilePicker', {
        configurable: true,
        value: undefined,
      });
    });
  }
}

async function launchPersistentMatrixContext(
  profile: string,
  options: { forceDownloadSave?: boolean } = {},
): Promise<BrowserContext> {
  const context = await chromium.launchPersistentContext(profile, {
    ...(persistentBrowserChannel ? { channel: persistentBrowserChannel as 'chrome' } : {}),
    args: process.env.ONLYOFFICE_CF_MATRIX_USE_SYSTEM_PROXY === '1' ? [] : ['--no-proxy-server'],
    acceptDownloads: true,
    headless: true,
    locale: 'en-US',
    serviceWorkers: 'allow',
    viewport: { width: 1280, height: 900 },
  });
  await configureMatrixContext(context, options);
  return context;
}

async function browserFetch<T>(
  page: Page,
  url: string,
  init?: { method?: string; headers?: Record<string, string>; cache?: RequestCache; body?: string },
): Promise<{ status: number; headers: Record<string, string>; body: T }> {
  return page.evaluate(
    async ({ target, requestInit }) => {
      const response = await fetch(target, requestInit);
      const headers = Object.fromEntries(response.headers.entries());
      const contentType = response.headers.get('content-type') || '';
      const body = contentType.includes('json') ? await response.json() : await response.text();
      return { status: response.status, headers, body };
    },
    { target: url, requestInit: init },
  ) as Promise<{ status: number; headers: Record<string, string>; body: T }>;
}

async function switchStableV5Pointer(
  page: Page,
  pointer: { version: 1; releaseId: string; manifestUrl: string; manifestSha256: string },
): Promise<void> {
  expect(matrixControlToken, 'The local matrix control token must be injected by the runner').toBeTruthy();
  const result = await browserFetch<{ ok: boolean; releaseId: string }>(
    page,
    `${canonicalOrigin}/__matrix__/stable-v5`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-OnlyOffice-Matrix-Control-Token': matrixControlToken!,
      },
      body: JSON.stringify(pointer),
    },
  );
  expect(result).toMatchObject({
    status: 200,
    body: { ok: true, releaseId: pointer.releaseId },
  });
  const stored = await browserFetch<typeof pointer>(page, `${canonicalOrigin}/channels/stable-v5.json`, {
    cache: 'no-store',
  });
  expect(stored.status).toBe(200);
  expect(stored.body).toEqual(pointer);
}

function planColdRelease(
  manifest: ReleaseManifest,
  manifestSha256: string,
): {
  requiredObjects: Map<string, number>;
  transferSegments: Map<string, number>;
} {
  const releaseManifest = manifest as unknown as ReleaseManifestV5;
  const model = releaseManifestV5ToContentModel(releaseManifest, manifestSha256);
  const contentPlan = planReleaseContent(model, [], [], { preferCanonicalForColdInstall: true });
  const transport: CanonicalPackageTransport = {
    bytes: releaseManifest.package.bytes,
    headerBytes: releaseManifest.package.headerBytes,
    segments: releaseManifest.package.segments.map((segment) => ({
      offset: segment.offset,
      bytes: segment.bytes,
      sha256: segment.sha256,
    })),
    assets: releaseManifest.assets.map((asset) => ({
      path: asset.path,
      packageOffset: asset.packageOffset,
      bytes: asset.bytes,
      sha256: asset.sha256,
      mime: asset.mime,
    })),
  };
  return {
    requiredObjects: new Map(contentPlan.requiredObjects.map((object) => [object.sha256, object.bytes])),
    transferSegments: new Map(
      planCanonicalPackageTransfer(model, transport, contentPlan).map((segment) => [segment.sha256, segment.bytes]),
    ),
  };
}

async function readContentCounters(page: Page): Promise<Record<string, ContentCounter>> {
  const result = await browserFetch<{
    objects: Record<string, ContentCounter>;
  }>(page, `${canonicalOrigin}/__matrix__/content-counters`, { cache: 'no-store' });
  expect(result.status).toBe(200);
  return result.body.objects;
}

async function readRouteCounters(page: Page): Promise<Record<string, ContentCounter>> {
  const result = await browserFetch<Record<string, ContentCounter>>(
    page,
    `${canonicalOrigin}/__matrix__/route-counters`,
    { cache: 'no-store' },
  );
  expect(result.status).toBe(200);
  return result.body;
}

function counterDelta(
  before: Record<string, ContentCounter>,
  after: Record<string, ContentCounter>,
  key: string,
): ContentCounter {
  const initial = before[key] || emptyContentCounter();
  const current = after[key] || emptyContentCounter();
  const statuses = Object.fromEntries(
    [...new Set([...Object.keys(initial.statuses), ...Object.keys(current.statuses)])].map((status) => [
      status,
      (current.statuses[status] || 0) - (initial.statuses[status] || 0),
    ]),
  );
  return {
    workerRequests: current.workerRequests - initial.workerRequests,
    cacheHits: current.cacheHits - initial.cacheHits,
    r2Heads: current.r2Heads - initial.r2Heads,
    r2Gets: current.r2Gets - initial.r2Gets,
    declaredBytes: current.declaredBytes - initial.declaredBytes,
    actualBytes: current.actualBytes - initial.actualBytes,
    r2Bytes: current.r2Bytes - initial.r2Bytes,
    completed: current.completed - initial.completed,
    aborted: current.aborted - initial.aborted,
    failed: current.failed - initial.failed,
    stalled: current.stalled - initial.stalled,
    statuses,
  };
}

async function cacheInventory(
  cdp: Awaited<ReturnType<BrowserContext['newCDPSession']>>,
  securityOrigin: string,
): Promise<Array<{ cacheName: string; urls: string[] }>> {
  const { caches: cacheDescriptors } = await cdp.send('CacheStorage.requestCacheNames', { securityOrigin });
  const inventory: Array<{ cacheName: string; urls: string[] }> = [];
  for (const descriptor of cacheDescriptors) {
    const urls: string[] = [];
    let skipCount = 0;
    const pageSize = 250;
    while (true) {
      const entries = await cdp.send('CacheStorage.requestEntries', {
        cacheId: descriptor.cacheId,
        skipCount,
        pageSize,
      });
      urls.push(...entries.cacheDataEntries.map((entry) => entry.requestURL));
      skipCount += entries.cacheDataEntries.length;
      if (entries.cacheDataEntries.length < pageSize) break;
    }
    inventory.push({ cacheName: descriptor.cacheName, urls });
  }
  return inventory;
}

type BrowserStorageInventory = {
  caches: Array<{
    cacheName: string;
    entries: Array<{
      url: string;
      status: number;
      contentLength: number | null;
      contentSha256: string | null;
    }>;
  }>;
  databases: Array<{ name: string; version: number }>;
  usage: number | null;
  quota: number | null;
};

async function browserStorageInventory(target: Page | Frame): Promise<BrowserStorageInventory> {
  return target.evaluate(async () => {
    const cacheNames = await caches.keys();
    const cacheInventory = [];
    for (const cacheName of cacheNames) {
      const cache = await caches.open(cacheName);
      const requests = await cache.keys();
      const entries = [];
      for (const request of requests) {
        const response = await cache.match(request);
        const rawLength = response?.headers.get('content-length') || '';
        entries.push({
          url: request.url,
          status: response?.status || 0,
          contentLength: /^(?:0|[1-9]\d*)$/.test(rawLength) ? Number(rawLength) : null,
          contentSha256:
            response?.headers.get('x-onlyoffice-content-sha256') || response?.headers.get('x-content-sha256') || null,
        });
      }
      cacheInventory.push({ cacheName, entries });
    }
    const databaseFactory = indexedDB as IDBFactory & {
      databases?: () => Promise<Array<{ name?: string; version?: number }>>;
    };
    const databases = databaseFactory.databases
      ? (await databaseFactory.databases())
          .filter((database): database is { name: string; version?: number } => Boolean(database.name))
          .map((database) => ({ name: database.name, version: database.version || 0 }))
      : [];
    const estimate = await navigator.storage.estimate();
    return {
      caches: cacheInventory,
      databases,
      usage: typeof estimate.usage === 'number' ? estimate.usage : null,
      quota: typeof estimate.quota === 'number' ? estimate.quota : null,
    };
  });
}

function isPermittedEditorShellCacheUrl(value: string, targetReleaseIds: string | readonly string[]): boolean {
  const url = new URL(value);
  const releases = typeof targetReleaseIds === 'string' ? [targetReleaseIds] : targetReleaseIds;
  return releases.some((targetReleaseId) => {
    const encodedReleaseId = encodeURIComponent(targetReleaseId);
    const releasePrefix = `/r/${encodedReleaseId}/`;
    if (url.pathname === `/.onlyoffice-internal/editor-shell/${encodedReleaseId}`) return true;
    if (!url.pathname.startsWith(releasePrefix)) return false;
    const path = url.pathname.slice(releasePrefix.length);
    return (
      path === 'office-host.html' ||
      path === 'editor-shell-prime.html' ||
      /^assets\/[a-zA-Z0-9._+-]+\.(?:css|js)$/.test(path)
    );
  });
}

type StandaloneDocumentCase = {
  type: 'docx' | 'xlsx' | 'pptx';
  extension: 'docx' | 'xlsx' | 'pptx';
  origin: string;
};

const productionDocumentCases: StandaloneDocumentCase[] = [
  { type: 'docx', extension: 'docx', origin: productionEditorOrigins[0] },
  { type: 'xlsx', extension: 'xlsx', origin: productionEditorOrigins[1] },
  { type: 'pptx', extension: 'pptx', origin: productionEditorOrigins[2] },
];

function standaloneAppUrl(documentCase: StandaloneDocumentCase, targetReleaseId = releaseId): string {
  return `${canonicalOrigin}/?resourcePrefetch=manual&hostUrl=${encodeURIComponent(
    `${documentCase.origin}/r/${targetReleaseId}/office-host.html`,
  )}`;
}

function findVersionedHostFrame(page: Page, targetReleaseId: string): Frame | undefined {
  const expectedPath = `/r/${encodeURIComponent(targetReleaseId)}/office-host.html`;
  return page.frames().find((frame) => {
    try {
      return new URL(frame.url()).pathname === expectedPath;
    } catch {
      return false;
    }
  });
}

function countVersionedHostFrames(context: BrowserContext): number {
  return context
    .pages()
    .flatMap((page) => page.frames())
    .filter((frame) => {
      try {
        return /^\/r\/[^/]+\/office-host\.html$/.test(new URL(frame.url()).pathname);
      } catch {
        return false;
      }
    }).length;
}

async function readVersionedHostAssetSha256(
  frame: Frame,
  targetReleaseId: string,
  assetPath: string,
): Promise<{ bytes: number; sha256: string; status: number }> {
  return frame.evaluate(
    async ({ release, path }) => {
      const response = await fetch(`/r/${encodeURIComponent(release)}/${path}`);
      const bytes = await response.arrayBuffer();
      const digest = await crypto.subtle.digest('SHA-256', bytes);
      return {
        bytes: bytes.byteLength,
        sha256: Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join(''),
        status: response.status,
      };
    },
    { release: targetReleaseId, path: assetPath },
  );
}

async function waitForStandaloneEditor(page: Page, documentCase: StandaloneDocumentCase): Promise<void> {
  await expect(page.locator('#resource-button')).toContainText('Resources ready', { timeout: 60_000 });
  await page.waitForFunction(
    () => {
      const demo = window.__officeDemo as { editor: { getState(): { status: string } } | null } | undefined;
      return demo?.editor?.getState().status === 'ready';
    },
    null,
    { timeout: 3 * 60_000 },
  );
  await expect(page.locator('#document-status')).toHaveText(documentCase.extension.toUpperCase(), {
    timeout: 30_000,
  });
}

async function openNewStandaloneDocument(
  context: BrowserContext,
  documentCase: StandaloneDocumentCase,
  failures: BrowserFailure[],
  targetReleaseId = releaseId,
): Promise<Page> {
  const page = await context.newPage();
  collectFailures(page, failures);
  await page.goto(standaloneAppUrl(documentCase, targetReleaseId), { waitUntil: 'domcontentloaded' });
  await expect(page.locator('#resource-button')).toContainText('Resources ready', { timeout: 60_000 });
  await page.evaluate((type) => {
    const demo = window.__officeDemo as {
      openEmpty(emptyType: 'docx' | 'xlsx' | 'pptx'): void;
    };
    demo.openEmpty(type);
  }, documentCase.type);
  await waitForStandaloneEditor(page, documentCase);
  return page;
}

async function openExistingStandaloneDocument(
  context: BrowserContext,
  documentCase: StandaloneDocumentCase,
  fixturePath: string,
  failures: BrowserFailure[],
  targetReleaseId = releaseId,
): Promise<Page> {
  const page = await context.newPage();
  collectFailures(page, failures);
  await page.goto(standaloneAppUrl(documentCase, targetReleaseId), { waitUntil: 'domcontentloaded' });
  await expect(page.locator('#resource-button')).toContainText('Resources ready', { timeout: 60_000 });
  await page.locator('#file-input').setInputFiles(fixturePath);
  await waitForStandaloneEditor(page, documentCase);
  return page;
}

async function makeStandaloneSpreadsheetDirty(page: Page): Promise<void> {
  const editor = page.locator('#editor-slot');
  await expect(editor).toBeVisible();
  const box = await editor.boundingBox();
  if (!box) throw new Error('Standalone spreadsheet editor is not visible');

  for (let attempt = 0; attempt < 4; attempt += 1) {
    await page.mouse.click(
      box.x + Math.min(260 + attempt * 28, box.width - 80),
      box.y + Math.min(220 + attempt * 24, box.height - 80),
    );
    await page.keyboard.type(`cloudflare-restart-${Date.now()}-${attempt}`);
    await page.keyboard.press('Enter');
    try {
      await page.waitForFunction(
        () => {
          const demo = window.__officeDemo as { editor: { getState(): { dirty: boolean } } | null } | undefined;
          return demo?.editor?.getState().dirty === true;
        },
        null,
        { timeout: 8_000 },
      );
      return;
    } catch {
      // Spreadsheet focus can vary slightly with the rendered toolbar height.
    }
  }
  throw new Error('Timed out waiting for the restarted spreadsheet to become dirty');
}

async function saveStandaloneDocument(
  page: Page,
  documentCase: StandaloneDocumentCase,
  destinationDirectory: string,
  prefix: string,
): Promise<string> {
  const downloadPromise = page.waitForEvent('download', { timeout: 3 * 60_000 });
  await page.locator('#save-button').click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(new RegExp(`\\.${documentCase.extension}$`, 'i'));
  const temporaryPath = await download.path();
  if (!temporaryPath) throw new Error(`Saved ${documentCase.type} download has no local path`);
  const destination = path.join(destinationDirectory, `${prefix}.${documentCase.extension}`);
  fs.copyFileSync(temporaryPath, destination);
  expect(fs.statSync(destination).size).toBeGreaterThan(0);
  const descriptor = fs.openSync(destination, 'r');
  try {
    const magic = Buffer.alloc(4);
    expect(fs.readSync(descriptor, magic, 0, magic.length, 0)).toBe(magic.length);
    expect([...magic]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  } finally {
    fs.closeSync(descriptor);
  }
  return destination;
}

async function destroyStandaloneDocuments(pages: Page[]): Promise<void> {
  await Promise.all(
    pages.map((page) =>
      page.evaluate(async () => {
        const demo = window.__officeDemo as {
          editor: { destroy(): Promise<void> } | null;
        };
        await demo.editor?.destroy();
      }),
    ),
  );
  await Promise.all(pages.map((page) => page.close()));
}

test('Release v5 routes reproduce production MIME, cache, range, CAS, and immutable Office Pack behavior', async ({
  browser,
}) => {
  test.skip(matrixMode !== 'full-v5', 'The minimal broker fixture intentionally exercises the legacy v4 shape.');
  expect(releaseId).toBeTruthy();
  const context = await freshContext(browser);
  const page = await context.newPage();
  await page.goto(`${canonicalOrigin}/channels/stable-v5.json`);
  const unknownHostPage = await context.newPage();
  const unknownHostResponse = await unknownHostPage.goto(`http://foo.localhost:${port}/channels/stable-v5.json`);
  expect(unknownHostResponse?.status()).toBe(404);
  await unknownHostPage.close();

  const pointer = await browserFetch<{
    version: number;
    releaseId: string;
    manifestUrl: string;
    manifestSha256: string;
  }>(page, `${canonicalOrigin}/channels/stable-v5.json`);
  expect(pointer.status).toBe(200);
  expect(pointer.headers['cache-control']).toBe('no-store');
  expect(pointer.body).toMatchObject({
    version: 1,
    releaseId,
    manifestUrl: `/releases/${releaseId}/manifest.json`,
    manifestSha256: expect.stringMatching(/^[a-f0-9]{64}$/),
  });

  const manifestResult = await browserFetch<ReleaseManifest>(
    page,
    `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
  );
  expect(manifestResult.status).toBe(200);
  expect(manifestResult.body.version).toBe(5);
  expect(manifestResult.body.releaseId).toBe(releaseId);
  expect(manifestResult.body.packageVersion).toBe(packageVersion);
  expect(manifestResult.body.package.path).toBe('office-resources.oobpack');
  expect(manifestResult.body.contentProtocol).toMatchObject({
    version: 1,
    digest: 'sha256',
    cacheKeyFormat: 'canonical-sha256-v1',
    storageSetSha256: expect.stringMatching(/^[a-f0-9]{64}$/),
  });

  const packMagic = await browserFetch<string>(page, `${canonicalOrigin}/p/${releaseId}/office-resources.oobpack`, {
    headers: { Range: 'bytes=0-7' },
  });
  expect(packMagic.status).toBe(206);
  expect(packMagic.body).toBe('OOBPACK1');
  expect(packMagic.headers['content-type']).toBe('application/vnd.onlyoffice.browser-pack');
  const firstSegment = manifestResult.body.package.segments[0];
  const segment = await browserFetch<string>(page, `${canonicalOrigin}/segments/sha256/${firstSegment.sha256}`, {
    method: 'HEAD',
  });
  expect(segment.status).toBe(200);
  expect(segment.headers['content-length']).toBe(String(firstSegment.bytes));
  expect(segment.headers['content-type']).toBe('application/vnd.onlyoffice.browser-pack-segment');

  const contentAsset = manifestResult.body.assets.find(
    (asset) => asset.bytes > 1 && asset.representations?.whole?.sha256 === asset.sha256,
  );
  expect(contentAsset).toBeTruthy();
  const contentObject = await browserFetch<string>(
    page,
    `${canonicalOrigin}/objects/${releaseId}/sha256/${contentAsset!.sha256}`,
    { headers: { Range: 'bytes=0-0' }, cache: 'no-store' },
  );
  expect(contentObject.status).toBe(206);
  expect(contentObject.headers['content-length']).toBe('1');
  expect(contentObject.headers['content-range']).toBe(`bytes 0-0/${contentAsset!.bytes}`);
  expect(contentObject.headers['x-content-sha256']).toBe(contentAsset!.sha256);
  expect(contentObject.headers['cache-control']).toContain('immutable');

  // Exercise a complete multi-megabyte content-object stream in a real
  // Chromium page. The range probe above can pass even when a streamed 200
  // response advertises the wrong length or terminates early.
  const largeContentObject = manifestResult.body.assets.find(
    (asset) => asset.bytes > 8 * 1024 * 1024 && asset.representations?.whole?.sha256 === asset.sha256,
  );
  if (largeContentObject) {
    const largeRead = await page.evaluate(
      async ({ target, expectedBytes }) => {
        const response = await fetch(target, { cache: 'no-store', credentials: 'omit' });
        const bytes = new Uint8Array(await response.arrayBuffer());
        const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
          .map((value) => value.toString(16).padStart(2, '0'))
          .join('');
        return {
          status: response.status,
          bytes: bytes.byteLength,
          contentLength: response.headers.get('content-length'),
          digest,
          expectedBytes,
        };
      },
      {
        target: `${canonicalOrigin}/objects/${releaseId}/sha256/${largeContentObject.sha256}`,
        expectedBytes: largeContentObject.bytes,
      },
    );
    expect(largeRead).toEqual({
      status: 200,
      bytes: largeContentObject.bytes,
      contentLength: String(largeContentObject.bytes),
      digest: largeContentObject.sha256,
      expectedBytes: largeContentObject.bytes,
    });
  }

  const largePackageSegment = manifestResult.body.package.segments.find(
    (candidate) => candidate.bytes > 8 * 1024 * 1024,
  );
  if (largePackageSegment) {
    const packageRead = await page.evaluate(
      async ({ target, expectedBytes, expectedSha256 }) => {
        const response = await fetch(target, { cache: 'no-store', credentials: 'omit' });
        const bytes = new Uint8Array(await response.arrayBuffer());
        const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)))
          .map((value) => value.toString(16).padStart(2, '0'))
          .join('');
        return {
          status: response.status,
          bytes: bytes.byteLength,
          contentLength: response.headers.get('content-length'),
          digest,
          expectedBytes,
          expectedSha256,
        };
      },
      {
        target: `${canonicalOrigin}/objects/${releaseId}/sha256/${largePackageSegment.sha256}`,
        expectedBytes: largePackageSegment.bytes,
        expectedSha256: largePackageSegment.sha256,
      },
    );
    expect(packageRead).toEqual({
      status: 200,
      bytes: largePackageSegment.bytes,
      contentLength: String(largePackageSegment.bytes),
      digest: largePackageSegment.sha256,
      expectedBytes: largePackageSegment.bytes,
      expectedSha256: largePackageSegment.sha256,
    });
  }

  const host = await browserFetch<string>(page, `${editorOrigin}/r/${releaseId}/office-host.html`, { method: 'HEAD' });
  expect(host.status).toBe(200);
  expect(host.headers['content-type']).toBe('text/html; charset=utf-8');
  expect(host.headers['cache-control']).toContain('immutable');
  expect(host.headers['x-onlyoffice-asset-version']).toBe(releaseId);
  expect(host.headers['x-onlyoffice-matrix-host']).toBe('aries.localhost');
  expect(host.headers['x-onlyoffice-matrix-isolated']).toBe('1');
  expect(host.headers['content-security-policy']).toContain(
    `frame-ancestors https://piwork.getpi.work http://onlyoffice.getpi.work http://piwork.localhost:${port} ${canonicalOrigin}`,
  );
  expect(host.headers['content-security-policy']).not.toContain('*');

  const wasm = manifestResult.body.assets.find((asset) => asset.mime === 'application/wasm');
  expect(wasm).toBeTruthy();
  const range = await browserFetch<string>(page, `${canonicalOrigin}/r/${releaseId}/${wasm!.path}`, {
    headers: { Range: 'bytes=0-1023' },
  });
  expect(range.status).toBe(206);
  expect(range.headers['content-range']).toBe(`bytes 0-1023/${wasm!.bytes}`);
  expect(range.headers['content-length']).toBe('1024');

  const missing = await browserFetch<string>(page, `${editorOrigin}/r/${releaseId}/missing-runtime-object.js`);
  expect(missing.status).toBe(404);
  await context.close();
});

test('fault-injected segment stream aborts and retries through the local Cloudflare gateway', async ({ browser }) => {
  test.skip(matrixMode !== 'full-v5' || !stallSegment, 'The fault-injected local matrix is opt-in.');
  test.setTimeout(30_000);
  const page = await browser.newPage();
  try {
    await page.goto(canonicalOrigin);
    const manifest = await browserFetch<ReleaseManifest>(
      page,
      `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
    );
    const segment = manifest.body.package.segments.find(({ sha256 }) => sha256 === stallSegment);
    expect(segment, 'The injected stall must target a manifest package segment').toBeTruthy();
    const target = `${canonicalOrigin}/segments/sha256/${stallSegment}?matrix-stall-test=1`;
    const beforeResult = await browserFetch<{ segments: Record<string, ContentCounter> }>(
      page,
      `${canonicalOrigin}/__matrix__/content-counters`,
      { cache: 'no-store' },
    );
    const before = beforeResult.body.segments;
    const first = await page.evaluate(async (url) => {
      const controller = new AbortController();
      const started = performance.now();
      const response = await fetch(url, { cache: 'no-store', signal: controller.signal });
      const reader = response.body?.getReader();
      if (!reader) throw new Error('stalled segment response has no body');
      const firstRead = await reader.read();
      const pending = reader.read();
      window.setTimeout(() => controller.abort(), 100);
      let aborted = false;
      try {
        await pending;
      } catch {
        aborted = true;
      }
      return { firstBytes: firstRead.value?.byteLength || 0, aborted, elapsedMs: performance.now() - started };
    }, target);
    expect(first.firstBytes).toBeGreaterThan(0);
    expect(first.elapsedMs).toBeLessThan(2_000);
    expect(first.aborted).toBe(true);

    const second = await page.evaluate(async (url) => {
      const response = await fetch(url, { cache: 'no-store' });
      return { status: response.status, bytes: (await response.arrayBuffer()).byteLength };
    }, target);
    expect(second).toEqual({ status: 200, bytes: segment!.bytes });
    const afterResult = await browserFetch<{ segments: Record<string, ContentCounter> }>(
      page,
      `${canonicalOrigin}/__matrix__/content-counters`,
      { cache: 'no-store' },
    );
    const after = afterResult.body.segments;
    const delta = counterDelta(before, after, stallSegment!);
    expect(delta.r2Gets).toBe(2);
    expect(delta.stalled).toBe(1);
    expect(delta.completed).toBeGreaterThanOrEqual(1);
  } finally {
    await page.close();
  }
});

test('synthetic canonical broker fixture stores one copy and reuses it across three editor origins', async ({
  browser,
}) => {
  test.skip(matrixMode !== 'synthetic-broker', 'Covered by the production v5 matrix in full mode.');
  const context = await freshContext(browser);
  const failures: BrowserFailure[] = [];
  const bootstrap = await context.newPage();
  collectFailures(bootstrap, failures);
  await bootstrap.goto(`${canonicalOrigin}/releases/${releaseId}/manifest.json`);
  const manifestResult = await browserFetch<ReleaseManifest>(
    bootstrap,
    `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
  );
  const segment = manifestResult.body.package.segments[0];
  const segmentUrl = `${canonicalOrigin}/segments/sha256/${segment.sha256}`;
  let segmentRequests = 0;
  context.on('request', (request) => {
    if (request.url() === segmentUrl) segmentRequests += 1;
  });

  const aries = await context.newPage();
  const taurus = await context.newPage();
  const gemini = await context.newPage();
  collectFailures(aries, failures);
  collectFailures(taurus, failures);
  collectFailures(gemini, failures);
  await Promise.all([
    aries.goto(`${editorOrigin}/__matrix__/cache-broker-editor.html`),
    taurus.goto(`http://taurus.localhost:${port}/__matrix__/cache-broker-editor.html`),
    gemini.goto(`http://gemini.localhost:${port}/__matrix__/cache-broker-editor.html`),
  ]);
  const cdp = await context.newCDPSession(aries);
  await cdp.send('Network.enable');
  const brokerProxyResponses: Array<{ fromServiceWorker: boolean; status: number }> = [];
  cdp.on('Network.responseReceived', ({ response }) => {
    if (response.url.includes('/__matrix__/broker-segments/')) {
      brokerProxyResponses.push({
        fromServiceWorker: Boolean(response.fromServiceWorker),
        status: response.status,
      });
    }
  });
  const storageBeforeInstall = await cdp.send('Storage.getUsageAndQuota', { origin: canonicalOrigin });
  await Promise.all([
    aries.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__),
    taurus.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__),
    gemini.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__),
  ]);
  const installs = await Promise.all(
    [aries, taurus, gemini].map((page) =>
      page.evaluate((sha256) => {
        const broker = globalThis as typeof globalThis & {
          __BROKER_REQUEST__: (message: object) => Promise<{
            reused: boolean;
            bytes: number;
            sha256: string;
          }>;
        };
        return broker.__BROKER_REQUEST__({ type: 'INSTALL', sha256 });
      }, segment.sha256),
    ),
  );
  expect(installs.filter((result) => result.reused)).toHaveLength(2);
  expect(installs.filter((result) => !result.reused)).toHaveLength(1);
  expect(segmentRequests).toBe(1);

  const reads = await Promise.all(
    [aries, taurus, gemini].map((page) =>
      page.evaluate((sha256) => {
        const broker = globalThis as typeof globalThis & {
          __BROKER_VERIFY__: (sha256: string) => Promise<{
            bytes: number;
            bytesSent: number;
            headers: {
              status: number;
              headers: {
                acceptRanges: string;
                contentLength: number;
                contentRange: string | null;
                contentType: string;
              };
            };
            sha256: string;
            matches: boolean;
          }>;
        };
        return broker.__BROKER_VERIFY__(sha256);
      }, segment.sha256),
    ),
  );
  expect(reads.every((result) => result.matches)).toBe(true);
  expect(reads.every((result) => result.bytes === segment.bytes && result.bytesSent === segment.bytes)).toBe(true);
  expect(reads[0].headers).toEqual({
    type: 'HEADERS',
    id: expect.any(String),
    status: 200,
    headers: {
      acceptRanges: 'bytes',
      contentLength: segment.bytes,
      contentRange: null,
      contentType: 'application/vnd.onlyoffice.browser-pack-segment',
    },
  });

  const serviceWorkerReads = await Promise.all(
    [aries, taurus, gemini].map((page) =>
      page.evaluate((sha256) => {
        const broker = globalThis as typeof globalThis & {
          __BROKER_SW_VERIFY__: (sha256: string) => Promise<{
            broker: string | null;
            bytes: number;
            contentLength: number;
            contentRange: string | null;
            contentType: string | null;
            releaseId: string | null;
            sha256: string;
            status: number;
          }>;
        };
        return broker.__BROKER_SW_VERIFY__(sha256);
      }, segment.sha256),
    ),
  );
  expect(serviceWorkerReads).toEqual([
    {
      broker: '1',
      bytes: segment.bytes,
      contentLength: segment.bytes,
      contentRange: null,
      contentType: 'application/vnd.onlyoffice.browser-pack-segment',
      releaseId,
      sha256: segment.sha256,
      status: 200,
    },
    {
      broker: '1',
      bytes: segment.bytes,
      contentLength: segment.bytes,
      contentRange: null,
      contentType: 'application/vnd.onlyoffice.browser-pack-segment',
      releaseId,
      sha256: segment.sha256,
      status: 200,
    },
    {
      broker: '1',
      bytes: segment.bytes,
      contentLength: segment.bytes,
      contentRange: null,
      contentType: 'application/vnd.onlyoffice.browser-pack-segment',
      releaseId,
      sha256: segment.sha256,
      status: 200,
    },
  ]);
  expect(brokerProxyResponses).toContainEqual({ fromServiceWorker: true, status: 200 });

  await cdp.send('ServiceWorker.enable');
  await cdp.send('ServiceWorker.stopAllWorkers');
  const recoveredAfterWorkerTermination = await aries.evaluate((sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_SW_VERIFY__: (
        sha256: string,
      ) => Promise<{ broker: string | null; bytes: number; sha256: string; status: number }>;
    };
    return broker.__BROKER_SW_VERIFY__(sha256);
  }, segment.sha256);
  expect(recoveredAfterWorkerTermination).toMatchObject({
    broker: '1',
    bytes: segment.bytes,
    sha256: segment.sha256,
    status: 200,
  });
  expect(segmentRequests).toBe(1);
  expect(await aries.evaluate(() => caches.keys())).toEqual(['onlyoffice-editor-broker-shell-experiment-v1']);
  expect(await taurus.evaluate(() => caches.keys())).toEqual(['onlyoffice-editor-broker-shell-experiment-v1']);
  expect(await gemini.evaluate(() => caches.keys())).toEqual(['onlyoffice-editor-broker-shell-experiment-v1']);

  const rangeStart = 150_000;
  const rangeEnd = 750_123;
  const expectedRange = Buffer.alloc(rangeEnd - rangeStart + 1);
  for (let index = 0; index < expectedRange.length; index += 1) {
    expectedRange[index] = ((rangeStart + index) * 31 + 17) & 0xff;
  }
  const expectedRangeSha256 = crypto.createHash('sha256').update(expectedRange).digest('hex');
  const rangeRead = await aries.evaluate(
    ({ sha256, start, end }) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_VERIFY__: (
          sha256: string,
          range: { start: number; end: number },
        ) => Promise<{
          bytes: number;
          bytesSent: number;
          headers: {
            status: number;
            headers: {
              acceptRanges: string;
              contentLength: number;
              contentRange: string | null;
              contentType: string;
            };
          };
          maxChunkBytes: number;
          sha256: string;
        }>;
      };
      return broker.__BROKER_VERIFY__(sha256, { start, end });
    },
    { sha256: segment.sha256, start: rangeStart, end: rangeEnd },
  );
  expect(rangeRead).toMatchObject({
    bytes: expectedRange.length,
    bytesSent: expectedRange.length,
    sha256: expectedRangeSha256,
    headers: {
      status: 206,
      headers: {
        acceptRanges: 'bytes',
        contentLength: expectedRange.length,
        contentRange: `bytes ${rangeStart}-${rangeEnd}/${segment.bytes}`,
        contentType: 'application/vnd.onlyoffice.browser-pack-segment',
      },
    },
  });
  expect(rangeRead.maxChunkBytes).toBeGreaterThan(0);
  expect(rangeRead.maxChunkBytes).toBeLessThanOrEqual(256 * 1024);
  const serviceWorkerRange = await taurus.evaluate(
    ({ sha256, start, end }) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_SW_VERIFY__: (
          sha256: string,
          range: { start: number; end: number },
        ) => Promise<{
          broker: string | null;
          bytes: number;
          contentLength: number;
          contentRange: string | null;
          sha256: string;
          status: number;
        }>;
      };
      return broker.__BROKER_SW_VERIFY__(sha256, { start, end });
    },
    { sha256: segment.sha256, start: rangeStart, end: rangeEnd },
  );
  expect(serviceWorkerRange).toMatchObject({
    broker: '1',
    bytes: expectedRange.length,
    contentLength: expectedRange.length,
    contentRange: `bytes ${rangeStart}-${rangeEnd}/${segment.bytes}`,
    sha256: expectedRangeSha256,
    status: 206,
  });

  const unsatisfiedRange = await taurus.evaluate(
    ({ sha256, start }) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_READ__: (
          sha256: string,
          range: { start: number; end: number },
        ) => Promise<{
          bytes: ArrayBuffer;
          byteLength: number;
          bytesSent: number;
          headers: {
            status: number;
            headers: { contentLength: number; contentRange: string };
          };
        }>;
      };
      return broker.__BROKER_READ__(sha256, { start, end: start + 1 });
    },
    { sha256: segment.sha256, start: segment.bytes },
  );
  expect(unsatisfiedRange.byteLength).toBe(0);
  expect(unsatisfiedRange.bytesSent).toBe(0);
  expect(unsatisfiedRange.headers).toMatchObject({
    status: 416,
    headers: { contentLength: 0, contentRange: `bytes */${segment.bytes}` },
  });
  const cancelledRead = await gemini.evaluate((sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_READ__: (
        sha256: string,
        range: undefined,
        cancelAfterFirst: boolean,
      ) => Promise<{
        bytes: ArrayBuffer;
        bytesSent: number;
        cancelled: boolean;
        maxChunkBytes: number;
      }>;
    };
    return broker.__BROKER_READ__(sha256, undefined, true);
  }, segment.sha256);
  expect(cancelledRead.cancelled).toBe(true);
  expect(cancelledRead.bytesSent).toBeGreaterThan(0);
  expect(cancelledRead.bytesSent).toBeLessThanOrEqual(256 * 1024);
  expect(cancelledRead.maxChunkBytes).toBeLessThanOrEqual(256 * 1024);
  const serviceWorkerCancelledRead = await aries.evaluate((sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_SW_CANCEL__: (
        sha256: string,
      ) => Promise<{ broker: string | null; firstChunkBytes: number; status: number }>;
    };
    return broker.__BROKER_SW_CANCEL__(sha256);
  }, segment.sha256);
  expect(serviceWorkerCancelledRead.broker).toBe('1');
  expect(serviceWorkerCancelledRead.status).toBe(200);
  expect(serviceWorkerCancelledRead.firstChunkBytes).toBeGreaterThan(0);
  expect(serviceWorkerCancelledRead.firstChunkBytes).toBeLessThanOrEqual(256 * 1024);

  const replay = await aries.evaluate(() => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_REPLAY_CONNECT__: () => Promise<{ type: string; error?: string }>;
    };
    return broker.__BROKER_REPLAY_CONNECT__();
  });
  expect(replay).toEqual({ type: 'CONNECT_ERROR', error: 'capability' });

  const attacker = await context.newPage();
  collectFailures(attacker, failures);
  await attacker.goto(`http://office-editor-evil.localhost:${port}/__matrix__/cache-broker-editor.html`);
  await attacker.waitForTimeout(300);
  expect(
    await attacker.evaluate(
      () => (globalThis as typeof globalThis & { __BROKER_CONNECTED__: boolean }).__BROKER_CONNECTED__,
    ),
  ).toBe(false);

  const rejectedMessages = await Promise.all([
    aries.evaluate(async () => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_REQUEST__: (message: object) => Promise<unknown>;
      };
      try {
        await broker.__BROKER_REQUEST__({ type: 'INSTALL', sha256: '0'.repeat(64) });
        return '';
      } catch (error) {
        return error instanceof Error ? error.message : String(error);
      }
    }),
    aries.evaluate(
      async ({ sha256, wrongRelease }) => {
        const broker = globalThis as typeof globalThis & {
          __BROKER_REQUEST__: (message: object) => Promise<unknown>;
        };
        try {
          await broker.__BROKER_REQUEST__({
            type: 'INSTALL',
            sha256,
            releaseId: wrongRelease,
          });
          return '';
        } catch (error) {
          return error instanceof Error ? error.message : String(error);
        }
      },
      { sha256: segment.sha256, wrongRelease: `${releaseId}-other` },
    ),
  ]);
  expect(rejectedMessages).toEqual(['manifest', 'release']);

  const stats = await aries.evaluate(() =>
    (
      globalThis as typeof globalThis & {
        __BROKER_REQUEST__: (message: object) => Promise<{
          ok: true;
          activeReads: number;
          cacheNames: string[];
          keys: string[];
        }>;
      }
    ).__BROKER_REQUEST__({ type: 'STATS' }),
  );
  expect(stats.cacheNames).toContain('onlyoffice-cloudflare-cache-broker-experiment-v3');
  expect(stats.cacheNames).toContain('onlyoffice-cache-broker-manifests-experiment-v2');
  expect(stats.cacheNames).toContain('onlyoffice-cache-broker-shell-experiment-v2');
  expect(stats.keys).toEqual([segmentUrl]);
  expect(stats.activeReads).toBe(0);
  const storageAfterInstall = await cdp.send('Storage.getUsageAndQuota', { origin: canonicalOrigin });
  const cacheStorageUsage = (result: typeof storageAfterInstall) =>
    result.usageBreakdown.find((entry) => entry.storageType === 'cache_storage')?.usage || 0;
  const cacheStorageGrowth = cacheStorageUsage(storageAfterInstall) - cacheStorageUsage(storageBeforeInstall);
  expect(cacheStorageGrowth).toBeGreaterThanOrEqual(segment.bytes);
  expect(cacheStorageGrowth).toBeLessThanOrEqual(segment.bytes + 256 * 1024);
  const countersBeforeClear = await browserFetch<
    Record<string, { workerRequests: number; r2Heads: number; r2Gets: number; r2Bytes: number }>
  >(bootstrap, `${canonicalOrigin}/__matrix__/segment-counters`);
  expect(countersBeforeClear.body[segment.sha256]).toMatchObject({
    workerRequests: 1,
    r2Heads: 1,
    r2Gets: 1,
    r2Bytes: segment.bytes,
  });

  await cdp.send('Network.clearBrowserCache');
  const readsAfterHttpCacheClear = await Promise.all(
    [aries, taurus, gemini].map((page) =>
      page.evaluate((sha256) => {
        const broker = globalThis as typeof globalThis & {
          __BROKER_SW_VERIFY__: (sha256: string) => Promise<{ sha256: string; status: number }>;
        };
        return broker.__BROKER_SW_VERIFY__(sha256);
      }, segment.sha256),
    ),
  );
  expect(readsAfterHttpCacheClear.every((result) => result.status === 200 && result.sha256 === segment.sha256)).toBe(
    true,
  );
  expect(segmentRequests).toBe(1);
  const countersAfterClear = await browserFetch<
    Record<string, { workerRequests: number; r2Heads: number; r2Gets: number; r2Bytes: number }>
  >(bootstrap, `${canonicalOrigin}/__matrix__/segment-counters`);
  expect(countersAfterClear.body).toEqual(countersBeforeClear.body);

  await context.setOffline(true);
  const offlineRead = await taurus.evaluate((sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_SW_VERIFY__: (sha256: string) => Promise<{
        broker: string | null;
        bytes: number;
        sha256: string;
        status: number;
      }>;
    };
    return broker.__BROKER_SW_VERIFY__(sha256);
  }, segment.sha256);
  expect(offlineRead).toMatchObject({
    broker: '1',
    bytes: segment.bytes,
    sha256: segment.sha256,
    status: 200,
  });
  await context.setOffline(false);
  expect(failures).toEqual([]);

  /* Recovery note: an intermediate duplicate of the production restart matrix
     was inserted into this synthetic test by parallel patch replay.
  await context.close();

  context = await launchPersistentMatrixContext(profile, { forceDownloadSave: true });
  await Promise.all(context.pages().map((page) => page.close()));
  const contentRequestsAfterRestart: string[] = [];
  context.on('request', (request) => {
    const pathname = new URL(request.url()).pathname;
    if (
      pathname.startsWith('/objects/') ||
      pathname.startsWith('/segments/sha256/') ||
      pathname.startsWith('/blobs/sha256/')
    ) {
      contentRequestsAfterRestart.push(request.url());
    }
  });
  const restartFailures: BrowserFailure[] = [];
  await context.setOffline(true);

  const newDocumentPages = await Promise.all(
    productionDocumentCases.map((documentCase) =>
      openNewStandaloneDocument(context, documentCase, restartFailures),
    ),
  );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    for (let index = 0; index < newDocumentPages.length; index += 1) {
      await saveStandaloneDocument(
        newDocumentPages[index],
        productionDocumentCases[index],
        savedDocuments,
        `new-${productionDocumentCases[index].type}`,
      );
    }
    await destroyStandaloneDocuments(newDocumentPages);
  expect(context.pages()).toHaveLength(0);

  const existingFixturePaths = productionDocumentCases.map((documentCase) =>
    path.resolve(`public/fixtures/office/local.${documentCase.extension}`),
  );
  for (const fixturePath of existingFixturePaths) {
    expect(fs.existsSync(fixturePath), `Missing restart fixture ${fixturePath}`).toBe(true);
  }
  const existingDocumentPages = await Promise.all(
    productionDocumentCases.map((documentCase, index) =>
      openExistingStandaloneDocument(
        context,
        documentCase,
        existingFixturePaths[index],
        restartFailures,
      ),
    ),
  );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    await makeStandaloneSpreadsheetDirty(existingDocumentPages[1]);
  const editedSpreadsheetState = await existingDocumentPages[1].evaluate(() => {
    const demo = window.__officeDemo as {
      editor: { getState(): { dirty: boolean } } | null;
    };
    return demo.editor?.getState();
  });
    expect(editedSpreadsheetState?.dirty).toBe(true);
    const savedExistingPaths: string[] = [];
    for (let index = 0; index < existingDocumentPages.length; index += 1) {
      savedExistingPaths.push(
        await saveStandaloneDocument(
          existingDocumentPages[index],
          productionDocumentCases[index],
        savedDocuments,
        `existing-${productionDocumentCases[index].type}`,
        ),
      );
    }
    await destroyStandaloneDocuments(existingDocumentPages);
  expect(context.pages()).toHaveLength(0);

  const reopenedDocumentPages = await Promise.all(
    productionDocumentCases.map((documentCase, index) =>
      openExistingStandaloneDocument(
        context,
        documentCase,
        savedExistingPaths[index],
        restartFailures,
      ),
    ),
  );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    for (let index = 0; index < reopenedDocumentPages.length; index += 1) {
      await saveStandaloneDocument(
        reopenedDocumentPages[index],
        productionDocumentCases[index],
        savedDocuments,
        `reopened-${productionDocumentCases[index].type}`,
      );
    }
    const usageAfterInstall = await cdp.send('Storage.getUsageAndQuota', { origin: canonicalOrigin });
    const storageUsage = (
      result: typeof usageAfterInstall,
      storageType: string,
    ): number => result.usageBreakdown.find((entry) => entry.storageType === storageType)?.usage || 0;
    const expectedUniqueBytes = [...expectedObjects.values()].reduce((sum, bytes) => sum + bytes, 0);
    const canonicalCacheGrowth =
      storageUsage(usageAfterInstall, 'cache_storage') - storageUsage(usageBeforeInstall, 'cache_storage');
    expect(canonicalCacheGrowth, 'canonical Cache Storage should contain one complete content set').toBeGreaterThanOrEqual(
      expectedUniqueBytes,
    );
    expect(canonicalCacheGrowth, 'canonical Cache Storage must not resemble one package per editor origin').toBeLessThan(
      expectedUniqueBytes + Math.max(64 * 1024 * 1024, Math.ceil(expectedUniqueBytes * 0.1)),
    );
    expect(storageUsage(usageAfterInstall, 'indexeddb'), 'IndexedDB must remain a metadata journal').toBeLessThan(
      32 * 1024 * 1024,
    );
    await owner.evaluate(async () => {
      const demo = window.__officeDemo as { closeAll(): Promise<void> };
      await demo.closeAll();
    });
    expect(findVersionedHostFrame(owner, releaseId!)).toBeUndefined();

  const restartedCdp = await context.newCDPSession(reopenedDocumentPages[0]);
  const restartedCanonicalCaches = await cacheInventory(restartedCdp, canonicalOrigin);
  const restartedContentCache = restartedCanonicalCaches.find(
    (cache) => cache.cacheName === 'onlyoffice-content-v1',
  );
  expect(restartedContentCache).toBeTruthy();
  const restartedContentUrls = new Set(restartedContentCache!.urls);
  for (const sha256 of expectedObjects.keys()) {
    expect(
      restartedContentUrls.has(`${canonicalOrigin}/__onlyoffice_content__/sha256/${sha256}`),
      `Restarted canonical cache should retain ${sha256}`,
    ).toBe(true);
  }
  for (const editor of productionEditorOrigins) {
    const editorCaches = await cacheInventory(restartedCdp, editor);
    expect(editorCaches.some((cache) => cache.cacheName === 'onlyoffice-content-v1')).toBe(false);
    expect(
      editorCaches
        .flatMap((cache) => cache.urls)
        .some((url) => /\/(?:objects\/|segments\/sha256\/|blobs\/sha256\/)/.test(new URL(url).pathname)),
      `${editor} must not duplicate canonical content after restart`,
    ).toBe(false);
    }
    const restartedEditorStorage = await Promise.all(
      reopenedDocumentPages.map(async (page, index) => {
        const frame = findVersionedHostFrame(page, releaseId!);
        expect(frame, `${productionEditorOrigins[index]} should restore its versioned frame`).toBeTruthy();
        return browserStorageInventory(frame!);
      }),
    );
    for (let index = 0; index < restartedEditorStorage.length; index += 1) {
      const storage = restartedEditorStorage[index];
      expect(storage.databases.map((database) => database.name)).not.toContain('onlyoffice-content-journal-v2');
      expect(
        storage.caches
          .flatMap((cache) => cache.entries)
          .every((entry) => isPermittedEditorShellCacheUrl(entry.url, releaseId!)),
        `${productionEditorOrigins[index]} must restore only its bounded shell cache`,
      ).toBe(true);
      expect(storage.usage ?? Number.POSITIVE_INFINITY).toBeLessThan(24 * 1024 * 1024);
    }
    expect(contentRequestsAfterRestart).toEqual([]);

  await context.setOffline(false);
  expect(await readContentCounters(reopenedDocumentPages[0])).toEqual(countersBeforeRestart);
  expect(contentRequestsAfterRestart).toEqual([]);
  expect(restartFailures).toEqual([]);
  await destroyStandaloneDocuments(reopenedDocumentPages);
  completed = true;
  } finally {
    await context.close().catch(() => undefined);
    if (completed) {
      fs.rmSync(profile, { force: true, recursive: true });
      fs.rmSync(savedDocuments, { force: true, recursive: true });
    } else {
      console.error(`Full-v5 Chromium profile retained for diagnosis: ${profile}`);
      console.error(`Full-v5 saved documents retained for diagnosis: ${savedDocuments}`);
    }
  }
});

  */
  await context.close();
});

test('canonical broker editor shell fault matrix keeps three origins explicit and recoverable', async ({ browser }) => {
  test.skip(
    matrixMode !== 'synthetic-broker',
    'The local Cloudflare/R2 fixture is used for this focused fault matrix.',
  );
  test.setTimeout(90_000);

  const representativeContext = await freshContext(browser);
  try {
    const representativeCases = [
      { host: 'aries', fault: 'cacheStorageOpenError', expectedMode: 'network' },
      { host: 'taurus', fault: 'cacheStoragePutError', expectedMode: 'network' },
      { host: 'gemini', fault: 'none', expectedMode: 'cache' },
    ] as const;
    const representativePages = await Promise.all(
      representativeCases.map(async ({ host, fault }) => {
        const page = await representativeContext.newPage();
        await page.goto(`http://${host}.localhost:${port}/__matrix__/editor-shell-fault/index.html?fault=${fault}`);
        return page;
      }),
    );
    const representativeResults = await Promise.all(
      representativePages.map((page) =>
        page.evaluate(
          () =>
            (
              globalThis as typeof globalThis & {
                __ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT_PROMISE__: Promise<unknown>;
              }
            ).__ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT_PROMISE__,
        ),
      ),
    );
    for (let index = 0; index < representativeResults.length; index += 1) {
      const result = representativeResults[index] as {
        ready: boolean;
        storageMode: string;
        stages: string[];
        stage?: string;
      };
      expect(result.ready, `${representativeCases[index].fault}: ${JSON.stringify(result)}`).toBe(true);
      expect(result.storageMode, representativeCases[index].fault).toBe(representativeCases[index].expectedMode);
      expect(result.stages).toEqual(['service-worker', 'shell-cache', 'shell-route', 'broker-probe']);
      expect(result.stage).not.toBe('unknown');
    }
    const cacheInventory = await Promise.all(representativePages.map((page) => page.evaluate(() => caches.keys())));
    expect(cacheInventory[0]).not.toContain('onlyoffice-editor-shell-fault-v1');
    expect(cacheInventory[1]).toContain('onlyoffice-editor-shell-fault-v1');
    expect(cacheInventory[2]).toContain('onlyoffice-editor-shell-fault-v1');
    await Promise.all(representativePages.map((page) => page.close()));
  } finally {
    await representativeContext.close();
  }

  const faultCases = [
    { fault: 'serviceWorkerNoController', stage: 'service-worker', code: 'timeout' },
    { fault: 'primeDelayMs', stage: 'shell-cache', code: 'timeout' },
    { fault: 'shellNetwork404', stage: 'shell-route', code: 'network' },
    { fault: 'shellReleaseMismatch', stage: 'shell-route', code: 'integrity' },
    { fault: 'brokerProbeTimeout', stage: 'broker-probe', code: 'timeout' },
    { fault: 'offlineAfterLastObject', stage: 'broker-probe', code: 'offline' },
    { fault: 'abortAtLastMegabyte', stage: 'broker-probe', code: 'aborted' },
  ] as const;
  for (const [index, expected] of faultCases.entries()) {
    const context = await freshContext(browser);
    try {
      const page = await context.newPage();
      const host = ['aries', 'taurus', 'gemini'][index % 3];
      await page.goto(
        `http://${host}.localhost:${port}/__matrix__/editor-shell-fault/index.html?fault=${expected.fault}`,
      );
      const result = (await page.evaluate(
        () =>
          (
            globalThis as typeof globalThis & {
              __ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT_PROMISE__: Promise<unknown>;
            }
          ).__ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT_PROMISE__,
      )) as { ready: boolean; stage: string; code: string; stages: string[] };
      expect(result.ready, expected.fault).toBe(false);
      expect(result.stage, expected.fault).toBe(expected.stage);
      expect(result.code, expected.fault).toBe(expected.code);
      expect(result.stage, expected.fault).not.toBe('unknown');
      expect(result.stages[0], expected.fault).toBe('service-worker');
      await page.close();
    } finally {
      await context.close();
    }
  }
});

test('synthetic canonical broker fixture survives a full Chromium restart and serves offline', async () => {
  test.skip(matrixMode !== 'synthetic-broker', 'Covered by the production v5 matrix in full mode.');
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-broker-restart-'));
  let context: BrowserContext | null = null;
  try {
    context = await launchPersistentMatrixContext(profile);
    let page = context.pages()[0] || (await context.newPage());
    const failures: BrowserFailure[] = [];
    collectFailures(page, failures);
    await page.goto(`${editorOrigin}/__matrix__/cache-broker-editor.html`);
    await page.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__);
    const manifestResult = await browserFetch<ReleaseManifest>(
      page,
      `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
    );
    const segment = manifestResult.body.package.segments[0];
    const segmentUrl = `${canonicalOrigin}/segments/sha256/${segment.sha256}`;
    await page.evaluate((sha256) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_REQUEST__: (message: object) => Promise<unknown>;
      };
      return broker.__BROKER_REQUEST__({ type: 'INSTALL', sha256 });
    }, segment.sha256);
    const onlineRead = await page.evaluate((sha256) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_SW_VERIFY__: (
          sha256: string,
        ) => Promise<{ broker: string | null; bytes: number; sha256: string; status: number }>;
      };
      return broker.__BROKER_SW_VERIFY__(sha256);
    }, segment.sha256);
    expect(onlineRead).toMatchObject({
      broker: '1',
      bytes: segment.bytes,
      sha256: segment.sha256,
      status: 200,
    });
    const countersBeforeRestart = await browserFetch<
      Record<string, { workerRequests: number; r2Heads: number; r2Gets: number; r2Bytes: number }>
    >(page, `${canonicalOrigin}/__matrix__/segment-counters`);
    expect(countersBeforeRestart.body[segment.sha256]).toBeTruthy();
    expect(failures).toEqual([]);
    const persistedEditorShell = await page.evaluate(async () => {
      const cache = await caches.open('onlyoffice-editor-broker-shell-experiment-v1');
      const registrations = await navigator.serviceWorker.getRegistrations();
      return {
        controlled: Boolean(navigator.serviceWorker.controller),
        keys: (await cache.keys()).map((request) => new URL(request.url).pathname),
        registrations: registrations.map((registration) => registration.scope),
      };
    });
    expect(persistedEditorShell).toMatchObject({
      controlled: true,
      keys: ['/__matrix__/cache-broker-editor.html'],
    });
    expect(persistedEditorShell.registrations).toContain(`${editorOrigin}/__matrix__/`);
    await context.close();
    context = null;
    await new Promise((resolve) => setTimeout(resolve, 750));

    context = await launchPersistentMatrixContext(profile);
    let segmentRequestsAfterRestart = 0;
    context.on('request', (request) => {
      if (request.url() === segmentUrl) segmentRequestsAfterRestart += 1;
    });
    page = context.pages()[0] || (await context.newPage());
    const restartFailures: BrowserFailure[] = [];
    collectFailures(page, restartFailures);
    await page.goto(`${editorOrigin}/__matrix__/cache-broker-editor.html`);
    await context.setOffline(true);
    await page.reload();
    const restartReady = await page.evaluate(async () => {
      const broker = globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> };
      return Promise.race([
        broker.__BROKER_READY__.then(() => ({ ready: true as const })),
        new Promise<{ ready: false; diagnostics: unknown }>((resolve) =>
          setTimeout(async () => {
            const registrations = await navigator.serviceWorker.getRegistrations();
            const cacheNames = await caches.keys();
            resolve({
              ready: false,
              diagnostics: {
                cacheNames,
                controller: navigator.serviceWorker.controller?.scriptURL || null,
                frames: Array.from(document.querySelectorAll('iframe'), (frame) => frame.src),
                registrations: registrations.map((registration) => ({
                  active: registration.active?.scriptURL || null,
                  scope: registration.scope,
                })),
              },
            });
          }, 10_000),
        ),
      ]);
    });
    expect(restartReady, 'offline Broker should reconnect after the persisted profile restarts').toEqual({
      ready: true,
    });
    const offlineRead = await page.evaluate((sha256) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_SW_VERIFY__: (
          sha256: string,
        ) => Promise<{ broker: string | null; bytes: number; sha256: string; status: number }>;
      };
      return broker.__BROKER_SW_VERIFY__(sha256);
    }, segment.sha256);
    expect(offlineRead).toMatchObject({
      broker: '1',
      bytes: segment.bytes,
      sha256: segment.sha256,
      status: 200,
    });
    expect(segmentRequestsAfterRestart).toBe(0);
    expect(restartFailures).toEqual([]);

    await context.setOffline(false);
    const countersAfterRestart = await browserFetch<
      Record<string, { workerRequests: number; r2Heads: number; r2Gets: number; r2Bytes: number }>
    >(page, `${canonicalOrigin}/__matrix__/segment-counters`);
    expect(countersAfterRestart.body).toEqual(countersBeforeRestart.body);
  } finally {
    await context?.close();
    fs.rmSync(profile, { force: true, recursive: true });
  }
});

test('synthetic canonical broker fixture pins release A while release B reuses only unchanged content', async ({
  browser,
}) => {
  test.skip(matrixMode !== 'synthetic-broker', 'The full matrix uses the production v5 installer and CAS.');
  expect(releaseId).toBeTruthy();
  expect(nextReleaseId).toBeTruthy();
  const context = await freshContext(browser);
  const failures: BrowserFailure[] = [];
  const bootstrap = await context.newPage();
  await bootstrap.goto(`${canonicalOrigin}/releases/${releaseId}/manifest.json`);
  const releaseA = await browserFetch<ReleaseManifest>(
    bootstrap,
    `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
  );
  const releaseB = await browserFetch<ReleaseManifest>(
    bootstrap,
    `${canonicalOrigin}/releases/${nextReleaseId}/manifest.json`,
  );
  const unchanged = releaseA.body.package.segments[0];
  const changed = releaseB.body.package.segments.find((segment) => segment.sha256 !== unchanged.sha256);
  expect(changed).toBeTruthy();
  const unchangedUrl = `${canonicalOrigin}/segments/sha256/${unchanged.sha256}`;
  const changedUrl = `${canonicalOrigin}/segments/sha256/${changed!.sha256}`;
  const networkRequests = new Map([
    [unchanged.sha256, 0],
    [changed!.sha256, 0],
  ]);
  context.on('request', (request) => {
    if (request.url() === unchangedUrl) {
      networkRequests.set(unchanged.sha256, (networkRequests.get(unchanged.sha256) || 0) + 1);
    }
    if (request.url() === changedUrl) {
      networkRequests.set(changed!.sha256, (networkRequests.get(changed!.sha256) || 0) + 1);
    }
  });

  const editorA = await context.newPage();
  const editorB = await context.newPage();
  collectFailures(editorA, failures);
  collectFailures(editorB, failures);
  await Promise.all([
    editorA.goto(`${editorOrigin}/__matrix__/cache-broker-editor.html?releaseId=${releaseId}`),
    editorB.goto(`http://taurus.localhost:${port}/__matrix__/cache-broker-editor.html?releaseId=${nextReleaseId}`),
  ]);
  await Promise.all([
    editorA.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__),
    editorB.evaluate(() => (globalThis as typeof globalThis & { __BROKER_READY__: Promise<void> }).__BROKER_READY__),
  ]);

  const install = (page: Page, sha256: string) =>
    page.evaluate((digest) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_REQUEST__: (message: object) => Promise<{ bytes: number; reused: boolean; sha256: string }>;
      };
      return broker.__BROKER_REQUEST__({ type: 'INSTALL', sha256: digest });
    }, sha256);
  expect(await install(editorA, unchanged.sha256)).toMatchObject({ reused: false });
  expect(await install(editorB, unchanged.sha256)).toMatchObject({ reused: true });
  expect(networkRequests.get(unchanged.sha256)).toBe(1);

  const missingBeforeInstall = await editorB.evaluate(async (sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_READ__: (sha256: string) => Promise<unknown>;
    };
    try {
      await broker.__BROKER_READ__(sha256);
      return '';
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  }, changed!.sha256);
  expect(missingBeforeInstall).toBe('missing');
  expect(await install(editorB, changed!.sha256)).toMatchObject({ reused: false });
  expect(networkRequests.get(changed!.sha256)).toBe(1);

  const verify = (page: Page, sha256: string) =>
    page.evaluate((digest) => {
      const broker = globalThis as typeof globalThis & {
        __BROKER_SW_VERIFY__: (
          sha256: string,
        ) => Promise<{ broker: string | null; releaseId: string | null; sha256: string; status: number }>;
      };
      return broker.__BROKER_SW_VERIFY__(digest);
    }, sha256);
  expect(await verify(editorA, unchanged.sha256)).toMatchObject({
    broker: '1',
    releaseId,
    sha256: unchanged.sha256,
    status: 200,
  });
  expect(await verify(editorB, changed!.sha256)).toMatchObject({
    broker: '1',
    releaseId: nextReleaseId,
    sha256: changed!.sha256,
    status: 200,
  });

  const crossReleaseFailure = await editorA.evaluate(async (sha256) => {
    const broker = globalThis as typeof globalThis & {
      __BROKER_REQUEST__: (message: object) => Promise<unknown>;
    };
    try {
      await broker.__BROKER_REQUEST__({ type: 'INSTALL', sha256 });
      return '';
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  }, changed!.sha256);
  expect(crossReleaseFailure).toBe('manifest');
  expect(await verify(editorA, unchanged.sha256)).toMatchObject({
    releaseId,
    sha256: unchanged.sha256,
    status: 200,
  });
  expect(networkRequests).toEqual(
    new Map([
      [unchanged.sha256, 1],
      [changed!.sha256, 1],
    ]),
  );
  expect(failures).toEqual([]);
  await context.close();
});

test('canonical broker service worker serves its capability-query shell after an offline restart', async () => {
  test.setTimeout(90_000);
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-broker-shell-restart-'));
  let context: BrowserContext | null = null;
  try {
    context = await launchPersistentMatrixContext(profile);
    let page = context.pages()[0] || (await context.newPage());
    await page.goto(`${canonicalOrigin}/resource-broker.html`);
    await page.evaluate(async () => {
      const registration = await navigator.serviceWorker.register('/sw.js', {
        scope: '/',
        updateViaCache: 'none',
      });
      await navigator.serviceWorker.ready;
      if (registration.active?.state === 'activated') return;
      await new Promise<void>((resolve, reject) => {
        const worker = registration.installing || registration.waiting || registration.active;
        if (!worker) {
          reject(new Error('canonical service worker was not installed'));
          return;
        }
        const timeout = window.setTimeout(
          () => reject(new Error('canonical service worker activation timed out')),
          30_000,
        );
        worker.addEventListener('statechange', () => {
          if (worker.state !== 'activated') return;
          window.clearTimeout(timeout);
          resolve();
        });
      });
    });
    await context.close();
    context = null;

    context = await launchPersistentMatrixContext(profile);
    await Promise.all(context.pages().map((candidate) => candidate.close()));
    await context.setOffline(true);
    page = await context.newPage();
    const brokerUrl = new URL('/resource-broker.html', canonicalOrigin);
    brokerUrl.searchParams.set('releaseId', 'offline-shell-probe');
    brokerUrl.searchParams.set('sessionId', 'offline-shell-probe');
    brokerUrl.searchParams.set('parentOrigin', 'https://onlyoffice.getpi.work');
    brokerUrl.searchParams.set('localMatrix', '1');
    const response = await page.goto(brokerUrl.href);
    expect(response?.status()).toBe(200);
    await expect(page).toHaveTitle('OnlyOffice Resource Broker');
    expect((await page.locator('body').textContent())?.trim()).toBe('');
  } finally {
    await context?.close();
    fs.rmSync(profile, { recursive: true, force: true });
  }
});

test('retained full-v5 profile can reopen its installed editor shell', async () => {
  const retainedProfile = process.env.ONLYOFFICE_CF_MATRIX_RETAINED_PROFILE;
  const retainedRelease = process.env.ONLYOFFICE_CF_MATRIX_RETAINED_RELEASE_ID;
  test.skip(!retainedProfile || !retainedRelease, 'This focused recovery check requires a retained failed profile.');
  test.setTimeout(90_000);
  const context = await launchPersistentMatrixContext(path.resolve(retainedProfile!));
  try {
    await Promise.all(context.pages().map((page) => page.close()));
    await context.setOffline(true);
    const page = await context.newPage();
    const failures: BrowserFailure[] = [];
    collectFailures(page, failures);
    await page.goto(standaloneAppUrl(productionDocumentCases[0], retainedRelease!));
    await page.locator('#resource-button').waitFor({ state: 'visible' });
    await page.waitForFunction(() => {
      const state = document.querySelector<HTMLButtonElement>('#resource-button')?.dataset.state;
      return state === 'ready' || state === 'error' || state === 'needed';
    });
    const recoveryState = await page.evaluate(async () => {
      const readTransactions = (database: IDBDatabase) =>
        new Promise<unknown[]>((resolve, reject) => {
          const rows: unknown[] = [];
          const request = database
            .transaction('releaseTransactions', 'readonly')
            .objectStore('releaseTransactions')
            .openCursor();
          request.onsuccess = () => {
            const cursor = request.result;
            if (!cursor) {
              resolve(rows);
              return;
            }
            const value = cursor.value as {
              releaseId?: unknown;
              manifestSha256?: unknown;
              state?: unknown;
              previousActiveReleaseId?: unknown;
              activationRollbackState?: unknown;
              failureCode?: unknown;
              requiredObjects?: unknown[];
              plannedMappings?: unknown[];
              committedMappings?: unknown[];
            };
            rows.push({
              releaseId: value.releaseId,
              manifestSha256: value.manifestSha256,
              state: value.state,
              previousActiveReleaseId: value.previousActiveReleaseId,
              activationRollbackState: value.activationRollbackState,
              failureCode: value.failureCode,
              requiredObjectCount: value.requiredObjects?.length ?? null,
              plannedMappingCount: value.plannedMappings?.length ?? null,
              committedMappingCount: value.committedMappings?.length ?? null,
            });
            cursor.continue();
          };
          request.onerror = () => reject(request.error);
        });
      const readMetadata = (database: IDBDatabase) =>
        new Promise<unknown[]>((resolve, reject) => {
          const rows: unknown[] = [];
          const request = database.transaction('metadata', 'readonly').objectStore('metadata').openCursor();
          request.onsuccess = () => {
            const cursor = request.result;
            if (!cursor) {
              resolve(rows);
              return;
            }
            rows.push({ key: cursor.key, value: cursor.value });
            cursor.continue();
          };
          request.onerror = () => reject(request.error);
        });
      const database = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('onlyoffice-content-journal-v2');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      try {
        return {
          snapshot: (window.__officeDemo as { resourceSnapshot?: unknown } | undefined)?.resourceSnapshot ?? null,
          resourceButton: {
            state: document.querySelector<HTMLButtonElement>('#resource-button')?.dataset.state ?? null,
            text: document.querySelector<HTMLButtonElement>('#resource-button')?.textContent ?? null,
          },
          transactions: await readTransactions(database),
          metadata: await readMetadata(database),
        };
      } finally {
        database.close();
      }
    });
    expect(recoveryState).toMatchObject({
      snapshot: {
        installedRelease: retainedRelease,
        readiness: 'ready',
        error: null,
      },
    });
    const query = new URLSearchParams({
      scenario: 'new-document',
      type: 'docx',
      hostUrl: `${productionEditorOrigins[0]}/r/${retainedRelease}/office-host.html`,
    });
    await page.goto(`${canonicalOrigin}/save-e2e.html?${query}`);
    try {
      await page.waitForFunction(
        () => {
          const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
          return status?.ready === true || Boolean(status?.error);
        },
        null,
        { timeout: 45_000 },
      );
    } catch (error) {
      const serviceWorkers = await Promise.all(
        context.serviceWorkers().map(async (worker) => ({
          url: worker.url(),
          metrics: await worker.evaluate(() => {
            const scope = globalThis as typeof globalThis & {
              __ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__?: () => unknown;
            };
            return scope.__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__?.() ?? null;
          }),
        })),
      );
      throw new Error(
        `${error instanceof Error ? error.message : String(error)}\nBroker diagnostics: ${JSON.stringify({
          frames: page.frames().map((frame) => frame.url()),
          serviceWorkers,
          status: await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus() ?? null),
        })}`,
      );
    }
    expect(await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__!.getStatus())).toMatchObject({
      error: '',
      ready: true,
    });
    expect(failures).toEqual([]);
  } finally {
    await context.close();
  }
});

test('production v5 installs once, survives restart with three offline editors, and updates incrementally', async () => {
  test.skip(matrixMode !== 'full-v5', 'The production installer requires a real v5 release.');
  test.setTimeout(15 * 60_000);
  expect(releaseId).toBeTruthy();
  const retainedProfile = process.env.ONLYOFFICE_CF_MATRIX_RETAINED_PROFILE;
  const profile = retainedProfile
    ? path.resolve(retainedProfile)
    : fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-full-v5-profile-'));
  if (retainedProfile && !fs.existsSync(profile)) throw new Error(`Retained Chrome profile does not exist: ${profile}`);
  const savedDocuments = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-full-v5-saved-'));
  let completed = false;
  let context = await launchPersistentMatrixContext(profile);
  try {
    const failures: BrowserFailure[] = [];
    let failurePhase: FailurePhase = 'normal';
    const collectMatrixFailures = (page: Page) => collectFailures(page, failures, () => failurePhase);
    const bootstrap = await context.newPage();
    collectMatrixFailures(bootstrap);
    await bootstrap.goto(canonicalOrigin);
    const release = await browserFetch<ReleaseManifest>(
      bootstrap,
      `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
    );
    expect(release.body.version).toBe(5);
    const stablePointer = await browserFetch<{
      version: number;
      releaseId: string;
      manifestUrl: string;
      manifestSha256: string;
    }>(bootstrap, `${canonicalOrigin}/channels/stable-v5.json`, { cache: 'no-store' });
    const coldReleasePlan = planColdRelease(release.body, stablePointer.body.manifestSha256);
    const expectedObjects = coldReleasePlan.requiredObjects;
    const expectedColdPackageSegments = coldReleasePlan.transferSegments;
    expect(expectedColdPackageSegments.size).toBeGreaterThan(0);
    const countersBeforeInstall = await readContentCounters(bootstrap);
    const appUrl = `${canonicalOrigin}/?resourcePrefetch=manual&hostUrl=${encodeURIComponent(
      `${productionEditorOrigins[0]}/r/${releaseId}/office-host.html`,
    )}`;
    const owner = await context.newPage();
    const follower = await context.newPage();
    collectMatrixFailures(owner);
    collectMatrixFailures(follower);
    await Promise.all([owner.goto(appUrl), follower.goto(appUrl)]);
    await Promise.all([
      expect(owner.locator('#resource-button')).toBeVisible(),
      expect(follower.locator('#resource-button')).toBeVisible(),
    ]);
    await expect(owner.locator('#version-label')).toContainText(`Version ${packageVersion}`);
    const cdp = await context.newCDPSession(owner);
    await cdp.send('Network.enable');
    const usageBeforeInstall = await cdp.send('Storage.getUsageAndQuota', { origin: canonicalOrigin });

    const dialog = owner.locator('#resource-dialog');
    await owner.getByText('New', { exact: true }).click();
    await owner.getByRole('button', { name: 'Word document' }).click();
    await expect(dialog).toBeVisible();
    await expect(owner.locator('iframe')).toHaveCount(0);
    const dialogBox = await dialog.boundingBox();
    const viewport = owner.viewportSize();
    expect(dialogBox).not.toBeNull();
    expect(viewport).not.toBeNull();
    expect(Math.abs(dialogBox!.x + dialogBox!.width / 2 - viewport!.width / 2)).toBeLessThanOrEqual(1);
    expect(Math.abs(dialogBox!.y + dialogBox!.height / 2 - viewport!.height / 2)).toBeLessThanOrEqual(1);
    await follower.locator('#resource-button').click();
    const followerDialog = follower.locator('#resource-dialog');
    await expect(followerDialog).toBeVisible();
    await Promise.all([
      owner.getByRole('button', { name: 'Install complete package' }).dispatchEvent('click'),
      follower.getByRole('button', { name: 'Install complete package' }).dispatchEvent('click'),
    ]);
    const pauseButtons = [owner, follower].map((page) => page.getByRole('button', { name: 'Pause' }));
    await expect
      .poll(
        async () => {
          const primerFailure = failures.find(
            ({ message }) =>
              message.includes('Failed to prime the editor origin shell') ||
              message.includes('Editor shell prime did not complete'),
          );
          if (primerFailure) throw new Error(primerFailure.message);
          for (const page of [owner, follower]) {
            const resourceFailure = page.getByText(/editor-origins\//).first();
            if (await resourceFailure.isVisible()) {
              throw new Error(`resource installer failed before Pause: ${await resourceFailure.innerText()}`);
            }
          }
          return (await Promise.all(pauseButtons.map((button) => button.isVisible()))).filter(Boolean).length;
        },
        {
          timeout: 90_000,
          message: 'at least the tab owning the shared install must expose Pause while network bytes are in flight',
        },
      )
      .toBeGreaterThan(0);
    const pauseVisibility = await Promise.all(pauseButtons.map((button) => button.isVisible()));
    const visiblePause = pauseButtons[pauseVisibility.findIndex(Boolean)];
    expect(visiblePause).toBeTruthy();
    await expect(visiblePause).toBeEnabled();
    await expect(owner.getByText('Stage 2/4', { exact: true })).toBeVisible();
    await expect(owner.getByText(/Package segment \d+\/\d+/, { exact: true })).toBeVisible();
    const readyLabel = owner.getByText('Resources ready', { exact: true }).first();
    await expect
      .poll(
        async () => {
          const productFailure = failures.find(({ message }) => message.includes('[onlyoffice-browser]'));
          if (productFailure) throw new Error(productFailure.message);
          return readyLabel.isVisible();
        },
        { timeout: 6 * 60_000, message: 'canonical installation must activate and expose Resources ready' },
      )
      .toBe(true);
    await expect(follower.locator('#resource-button')).toContainText('Resources ready', {
      timeout: 30_000,
    });
    await expect(owner.getByRole('progressbar')).toHaveCount(0);
    await expect(owner.locator('iframe')).toHaveCount(1, { timeout: 3 * 60_000 });
    await expect(owner.locator('#document-status')).toHaveText(/^(DOCX|●)$/, { timeout: 3 * 60_000 });
    await followerDialog.getByRole('button', { name: 'Close dialog' }).click();
    await expect(follower.getByRole('progressbar')).toHaveCount(0);

    const countersAfterInstallAndFirstEditor = await readContentCounters(bootstrap);
    for (const [sha256, bytes] of expectedColdPackageSegments) {
      const delta = counterDelta(countersBeforeInstall, countersAfterInstallAndFirstEditor, `${releaseId}:${sha256}`);
      expect(delta.workerRequests, `cold package segment ${sha256} should be requested exactly once`).toBe(1);
      expect(delta.r2Gets, `cold package segment ${sha256} should be read from local R2 exactly once`).toBe(1);
      expect(delta.r2Bytes, `cold package segment ${sha256} should transfer exactly once`).toBe(bytes);
    }
    for (const [sha256] of expectedObjects) {
      if (expectedColdPackageSegments.has(sha256)) continue;
      const delta = counterDelta(countersBeforeInstall, countersAfterInstallAndFirstEditor, `${releaseId}:${sha256}`);
      expect(delta.workerRequests, `final canonical object ${sha256} must be materialized without a second GET`).toBe(
        0,
      );
      expect(delta.r2Gets, `final canonical object ${sha256} must not be read separately from R2`).toBe(0);
      expect(delta.r2Bytes, `final canonical object ${sha256} must not duplicate cold-install bytes`).toBe(0);
    }

    const sentinelAsset =
      release.body.assets.find((asset) => asset.path === 'office-host.html') ||
      release.body.assets.find((asset) => asset.bytes > 0 && asset.bytes < 1024 * 1024);
    expect(sentinelAsset, 'The release needs a bounded immutable HTTP-cache sentinel').toBeTruthy();
    const sentinelNonce = crypto.randomBytes(8).toString('hex');
    const sentinelUrl = `${canonicalOrigin}/r/${encodeURIComponent(releaseId!)}/${sentinelAsset!.path}?matrix-http-cache-sentinel=${sentinelNonce}`;
    const sentinelRouteKey = `GET ${new URL(sentinelUrl).pathname}${new URL(sentinelUrl).search}`;
    const sentinelPage = await context.newPage();
    const fetchSentinel = () =>
      sentinelPage.evaluate(async (target) => {
        const response = await fetch(target, { cache: 'default', credentials: 'omit' });
        const bytes = (await response.arrayBuffer()).byteLength;
        return {
          status: response.status,
          bytes,
        };
      }, sentinelUrl);
    const sentinelBefore = await readRouteCounters(bootstrap);
    expect(await fetchSentinel()).toEqual({
      status: 200,
      bytes: sentinelAsset!.bytes,
    });
    const sentinelAfterFirst = await readRouteCounters(bootstrap);
    expect(counterDelta(sentinelBefore, sentinelAfterFirst, sentinelRouteKey)).toMatchObject({
      workerRequests: 1,
      r2Gets: 1,
      declaredBytes: sentinelAsset!.bytes,
      actualBytes: sentinelAsset!.bytes,
      completed: 1,
      failed: 0,
    });
    expect(await fetchSentinel()).toMatchObject({ status: 200, bytes: sentinelAsset!.bytes });
    const sentinelAfterSecond = await readRouteCounters(bootstrap);
    expect(
      counterDelta(sentinelAfterFirst, sentinelAfterSecond, sentinelRouteKey),
      'the second immutable fetch may use browser HTTP cache or the local Worker cache, but must never read R2 again',
    ).toMatchObject({ r2Gets: 0, r2Bytes: 0, actualBytes: 0, failed: 0 });
    expect(counterDelta(sentinelAfterFirst, sentinelAfterSecond, sentinelRouteKey).workerRequests).toBeLessThanOrEqual(
      1,
    );

    await cdp.send('Network.clearBrowserCache');
    expect(await fetchSentinel()).toMatchObject({ status: 200, bytes: sentinelAsset!.bytes });
    const sentinelAfterClear = await readRouteCounters(bootstrap);
    expect(counterDelta(sentinelAfterSecond, sentinelAfterClear, sentinelRouteKey)).toMatchObject({
      r2Gets: 0,
      r2Bytes: 0,
      actualBytes: 0,
      failed: 0,
    });
    await cdp.send('Network.clearBrowserCache');
    await sentinelPage.close();
    const closeInitialEditor = owner.evaluate(async () => {
      const demo = window.__officeDemo as { closeAll(): Promise<void> };
      await demo.closeAll();
    });
    const dirtyDialog = owner.locator('#dirty-dialog');
    await dirtyDialog.waitFor({ state: 'visible', timeout: 5_000 }).catch(() => undefined);
    if (await dirtyDialog.isVisible()) {
      await dirtyDialog.locator('button[value="discard"]').click();
    }
    await closeInitialEditor;
    await expect(owner.locator('iframe')).toHaveCount(0, { timeout: 30_000 });
    const countersBeforeThreeEditors = await readContentCounters(bootstrap);
    const documentPages = await Promise.all(
      productionDocumentCases.map(async ({ type, origin }) => {
        const page = await context.newPage();
        collectMatrixFailures(page);
        const query = new URLSearchParams({
          scenario: 'new-document',
          type,
          hostUrl: `${origin}/r/${releaseId}/office-host.html`,
        });
        await page.goto(`${canonicalOrigin}/save-e2e.html?${query}`);
        await page.waitForFunction(
          () => {
            const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
            return status?.ready === true || Boolean(status?.error);
          },
          null,
          { timeout: 3 * 60_000 },
        );
        const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__!.getStatus());
        expect(status.error, `${type} editor should start from canonical resources`).toBe('');
        expect(status.ready).toBe(true);
        return page;
      }),
    );
    const countersAfterThreeEditors = await readContentCounters(bootstrap);
    expect(countersAfterThreeEditors).toEqual(countersBeforeThreeEditors);

    const canonicalBrowserStorage = await browserStorageInventory(owner);
    expect(canonicalBrowserStorage.databases.map((database) => database.name)).toContain(
      'onlyoffice-content-journal-v2',
    );
    const canonicalBrowserContentCache = canonicalBrowserStorage.caches.find(
      (cache) => cache.cacheName === 'onlyoffice-content-v1',
    );
    expect(canonicalBrowserContentCache).toBeTruthy();
    expect(canonicalBrowserContentCache!.entries).toHaveLength(expectedObjects.size);
    for (const entry of canonicalBrowserContentCache!.entries) {
      const digest = new URL(entry.url).pathname.split('/').at(-1)!;
      expect(expectedObjects.has(digest), `canonical cache contains unexpected digest ${digest}`).toBe(true);
      expect(entry.status).toBe(200);
      expect(entry.contentLength).toBe(expectedObjects.get(digest));
      expect(entry.contentSha256).toBe(digest);
    }
    for (const cache of canonicalBrowserStorage.caches.filter(
      (candidate) => candidate.cacheName !== 'onlyoffice-content-v1',
    )) {
      expect(
        cache.entries.some((entry) => {
          const pathname = new URL(entry.url).pathname;
          return (
            pathname.startsWith('/r/') ||
            pathname.startsWith('/p/') ||
            pathname.includes('.oobpack') ||
            /^\/(?:objects|segments|blobs)\//.test(pathname) ||
            /^\/(?:sdkjs|web-apps|fonts|wasm)\//.test(pathname) ||
            pathname.startsWith('/__onlyoffice_content__/')
          );
        }),
        `${cache.cacheName} must not duplicate Office content through a general Workbox cache`,
      ).toBe(false);
    }

    const editorBrowserStorage = await Promise.all(
      documentPages.map(async (page, index) => {
        const frame = findVersionedHostFrame(page, releaseId!);
        expect(frame, `${productionDocumentCases[index].origin} should expose its installed editor frame`).toBeTruthy();
        return browserStorageInventory(frame!);
      }),
    );
    for (let index = 0; index < editorBrowserStorage.length; index += 1) {
      const storage = editorBrowserStorage[index];
      const editor = productionEditorOrigins[index];
      expect(storage.databases.map((database) => database.name)).not.toContain('onlyoffice-content-journal-v2');
      for (const cache of storage.caches) {
        for (const entry of cache.entries) {
          expect(
            isPermittedEditorShellCacheUrl(entry.url, releaseId!),
            `${editor} cache ${cache.cacheName} must contain only the bounded versioned editor shell: ${entry.url}`,
          ).toBe(true);
        }
      }
      expect(
        storage.usage ?? Number.POSITIVE_INFINITY,
        `${editor} storage must stay below the 16 MiB shell bound`,
      ).toBeLessThan(24 * 1024 * 1024);
    }
    expect(
      editorBrowserStorage.reduce((sum, storage) => sum + (storage.usage || 0), 0),
      'three editor origins must not contain copies of the canonical Office package',
    ).toBeLessThan(64 * 1024 * 1024);

    const canonicalCaches = await cacheInventory(cdp, canonicalOrigin);
    const canonicalContentCache = canonicalCaches.find((cache) => cache.cacheName === 'onlyoffice-content-v1');
    expect(canonicalContentCache).toBeTruthy();
    expect(canonicalContentCache!.urls).toHaveLength(expectedObjects.size);
    expect(
      canonicalContentCache!.urls.every((url) => new URL(url).pathname.startsWith('/__onlyoffice_content__/sha256/')),
    ).toBe(true);
    for (const cache of canonicalCaches.filter((candidate) => candidate.cacheName !== 'onlyoffice-content-v1')) {
      expect(
        cache.urls.some((url) => /\/(?:objects\/|segments\/sha256\/|blobs\/sha256\/)/.test(new URL(url).pathname)),
        `${cache.cacheName} must not duplicate canonical Office content`,
      ).toBe(false);
    }

    const usageAfterInstall = await cdp.send('Storage.getUsageAndQuota', { origin: canonicalOrigin });
    const storageUsage = (result: typeof usageAfterInstall, storageType: string): number =>
      result.usageBreakdown.find((entry) => entry.storageType === storageType)?.usage || 0;
    const expectedUniqueBytes = [...expectedObjects.values()].reduce((sum, bytes) => sum + bytes, 0);
    const canonicalCacheGrowth =
      storageUsage(usageAfterInstall, 'cache_storage') - storageUsage(usageBeforeInstall, 'cache_storage');
    expect(
      canonicalCacheGrowth,
      'canonical Cache Storage should contain one complete content set',
    ).toBeGreaterThanOrEqual(expectedUniqueBytes);
    expect(
      canonicalCacheGrowth,
      'canonical Cache Storage must not resemble one package per editor origin',
    ).toBeLessThan(expectedUniqueBytes + Math.max(64 * 1024 * 1024, Math.ceil(expectedUniqueBytes * 0.1)));
    expect(storageUsage(usageAfterInstall, 'indexeddb'), 'IndexedDB must remain a metadata journal').toBeLessThan(
      32 * 1024 * 1024,
    );

    const offlineRecoveryFailureStart = failures.length;
    failurePhase = 'offline-recovery';
    await context.setOffline(true);
    const saved = await Promise.all(
      documentPages.map((page) => page.evaluate(async () => window.__ONLYOFFICE_SAVE_E2E__!.save())),
    );
    for (let index = 0; index < saved.length; index += 1) {
      expect(saved[index].fileName).toMatch(new RegExp(`\\.${productionDocumentCases[index].extension}$`, 'i'));
      expect(saved[index].size).toBeGreaterThan(0);
      expect(saved[index].firstBytes.slice(0, 4)).toEqual([0x50, 0x4b, 0x03, 0x04]);
    }
    await context.setOffline(false);
    await bootstrap.waitForFunction(() => navigator.onLine, undefined, { timeout: 10_000 });
    await owner.waitForTimeout(1_500);
    const offlineRecoveryFailures = failures.splice(offlineRecoveryFailureStart);
    expect(offlineRecoveryFailures.filter((failure) => !isExpectedOfflineStablePointerFailure(failure))).toEqual([]);
    failurePhase = 'normal';
    expect(await readContentCounters(bootstrap)).toEqual(countersBeforeThreeEditors);
    await Promise.all(
      documentPages.map((page) =>
        page.evaluate(async () => {
          await window.__ONLYOFFICE_SAVE_E2E__!.destroy();
        }),
      ),
    );
    await Promise.all(documentPages.map((page) => page.close()));
    await owner.evaluate(() => {
      const demo = window.__officeDemo as {
        openEmpty(emptyType: 'docx'): void;
      };
      demo.openEmpty('docx');
    });
    await waitForStandaloneEditor(owner, productionDocumentCases[0]);

    expect(nextReleaseId).toBeTruthy();
    expect(nextManifestSha256).toMatch(/^[a-f0-9]{64}$/);
    expect(nextChangedPath).toBeTruthy();
    expect(nextChangedSha256).toMatch(/^[a-f0-9]{64}$/);
    expect(nextChangedBytes).toBeGreaterThan(0);
    const activeAHostFrame = findVersionedHostFrame(owner, releaseId!);
    expect(activeAHostFrame, 'Release A editor should remain mounted during the update').toBeTruthy();
    const activeAChangedAsset = release.body.assets.find((asset) => asset.path === nextChangedPath);
    expect(activeAChangedAsset, 'Derived release B must change a release A asset').toBeTruthy();
    const initialSecondarySnapshot = await owner.evaluate(() => {
      const manager = (
        window.__officeDemo as {
          resourceManager: {
            getSnapshot(): { installedRelease: string | null; readiness: string };
          } | null;
        }
      ).resourceManager;
      if (!manager) throw new Error('Standalone resource manager is not initialized');
      (
        globalThis as typeof globalThis & {
          __ONLYOFFICE_MATRIX_RESOURCE_MANAGER__: unknown;
        }
      ).__ONLYOFFICE_MATRIX_RESOURCE_MANAGER__ = manager;
      return manager.getSnapshot();
    });
    expect(initialSecondarySnapshot).toMatchObject({ installedRelease: releaseId, readiness: 'ready' });

    expect(failedReleaseId).toBeTruthy();
    expect(failedManifestSha256).toMatch(/^[a-f0-9]{64}$/);
    expect(failedChangedSha256).toMatch(/^[a-f0-9]{64}$/);
    failurePhase = 'online-update';
    const countersBeforeFailedUpdate = await readContentCounters(bootstrap);
    const failuresBeforeFailedUpdate = failures.length;
    await switchStableV5Pointer(bootstrap, {
      version: 1,
      releaseId: failedReleaseId!,
      manifestUrl: `/releases/${failedReleaseId}/manifest.json`,
      manifestSha256: failedManifestSha256!,
    });
    const failedUpdate = await owner.evaluate(async () => {
      const manager = (
        globalThis as typeof globalThis & {
          __ONLYOFFICE_MATRIX_RESOURCE_MANAGER__: {
            checkForUpdates(): Promise<void>;
            plan(request: { scope: 'all' }): Promise<{ releaseId: string }>;
            apply(plan: unknown): Promise<void>;
            getSnapshot(): {
              installedRelease: string | null;
              readiness: string;
              error: { code: string; path?: string } | null;
            };
          };
        }
      ).__ONLYOFFICE_MATRIX_RESOURCE_MANAGER__;
      await manager.checkForUpdates();
      const plan = await manager.plan({ scope: 'all' });
      let error = '';
      try {
        await manager.apply(plan);
      } catch (caught) {
        error = caught instanceof Error ? caught.message : String(caught);
      }
      return { plan, error, snapshot: manager.getSnapshot() };
    });
    expect(failedUpdate.plan).toMatchObject({ releaseId: failedReleaseId });
    expect(failedUpdate.error).toBeTruthy();
    expect(failedUpdate.snapshot).toMatchObject({
      installedRelease: releaseId,
      readiness: 'error',
      error: { code: 'network' },
    });
    const countersAfterFailedUpdate = await readContentCounters(bootstrap);
    const failedObjectDelta = counterDelta(
      countersBeforeFailedUpdate,
      countersAfterFailedUpdate,
      `${failedReleaseId}:${failedChangedSha256}`,
    );
    expect(failedObjectDelta.workerRequests).toBeGreaterThanOrEqual(1);
    expect(failedObjectDelta.r2Gets).toBe(0);
    expect(failedObjectDelta.actualBytes).toBe(0);
    expect(failedObjectDelta.statuses['404']).toBeGreaterThanOrEqual(1);
    const expectedFailedObjectUrl = `${canonicalOrigin}/objects/${failedReleaseId}/sha256/${failedChangedSha256}`;
    const isExpectedFailedObjectUrl = (candidate: string | undefined): boolean => {
      if (!candidate) return false;
      const actual = new URL(candidate);
      const expected = new URL(expectedFailedObjectUrl);
      return (
        actual.origin === expected.origin &&
        actual.pathname === expected.pathname &&
        [...actual.searchParams.keys()].every((key) => key === '__onlyoffice_transfer_retry') &&
        [...actual.searchParams.values()].every((value) => /^(?:0|[1-9]\d*)$/.test(value))
      );
    };
    const failedUpdateConsoleErrors = failures.slice(failuresBeforeFailedUpdate);
    // A handled fetch of the intentionally missing object does not have to emit
    // a browser console error. Chromium versions differ here: some surface the
    // 404 as a failed-resource console entry, while others keep it entirely in
    // the application error state. In either case the counter and structured
    // snapshot assertions above are the authoritative failure contract. If the
    // browser does emit a console entry, it must be exactly this expected 404.
    if (failedUpdateConsoleErrors.length > 0) {
      expect(
        failedUpdateConsoleErrors.every((failure) =>
          (() => {
            if (
              failure.source !== 'console' ||
              !failure.url ||
              !failure.message.startsWith('Failed to load resource:')
            ) {
              return false;
            }
            return isExpectedFailedObjectUrl(failure.url);
          })(),
        ),
        `unexpected browser failures during intentional failed update: ${JSON.stringify(failedUpdateConsoleErrors)}`,
      ).toBe(true);
    }
    failures.splice(failuresBeforeFailedUpdate);
    expect(
      Object.entries(countersAfterFailedUpdate)
        .filter(([key]) => key.startsWith(`${failedReleaseId}:`) && key !== `${failedReleaseId}:${failedChangedSha256}`)
        .every(([, counter]) => counter.workerRequests === 0),
      'a failed B install must not fetch unrelated objects',
    ).toBe(true);
    expect(await readVersionedHostAssetSha256(activeAHostFrame!, releaseId!, nextChangedPath!)).toEqual({
      bytes: activeAChangedAsset!.bytes,
      sha256: activeAChangedAsset!.sha256,
      status: 200,
    });
    expect(
      await owner.evaluate(() => {
        const demo = window.__officeDemo as {
          editor: { getState(): { status: string } } | null;
        };
        return demo.editor?.getState().status;
      }),
    ).toBe('ready');

    const countersBeforeIncrementalUpdate = countersAfterFailedUpdate;
    await switchStableV5Pointer(bootstrap, {
      version: 1,
      releaseId: nextReleaseId!,
      manifestUrl: `/releases/${nextReleaseId}/manifest.json`,
      manifestSha256: nextManifestSha256!,
    });
    const incremental = await owner.evaluate(async () => {
      const manager = (
        globalThis as typeof globalThis & {
          __ONLYOFFICE_MATRIX_RESOURCE_MANAGER__: {
            checkForUpdates(): Promise<void>;
            plan(request: { scope: 'all' }): Promise<{
              downloadBytes: number;
              reusedBytes: number;
              releaseId: string;
            }>;
            apply(plan: unknown): Promise<void>;
            getSnapshot(): { installedRelease: string | null; readiness: string };
          };
        }
      ).__ONLYOFFICE_MATRIX_RESOURCE_MANAGER__;
      await manager.checkForUpdates();
      const plan = await manager.plan({ scope: 'all' });
      await manager.apply(plan);
      return { plan, snapshot: manager.getSnapshot() };
    });
    expect(incremental.plan).toMatchObject({
      releaseId: nextReleaseId,
      downloadBytes: nextChangedBytes,
    });
    expect(incremental.plan.reusedBytes).toBeGreaterThan(0);
    expect(incremental.snapshot).toMatchObject({ installedRelease: nextReleaseId, readiness: 'ready' });
    const countersAfterIncrementalUpdate = await readContentCounters(bootstrap);
    expect(
      counterDelta(
        countersBeforeIncrementalUpdate,
        countersAfterIncrementalUpdate,
        `${nextReleaseId}:${nextChangedSha256}`,
      ),
    ).toEqual({
      workerRequests: 1,
      cacheHits: 0,
      r2Heads: 1,
      r2Gets: 1,
      declaredBytes: nextChangedBytes,
      actualBytes: nextChangedBytes,
      r2Bytes: nextChangedBytes,
      completed: 1,
      aborted: 0,
      failed: 0,
      stalled: 0,
      statuses: { '200': 1 },
    });
    expect(
      Object.entries(countersAfterIncrementalUpdate)
        .filter(([key]) => key.startsWith(`${nextReleaseId}:`) && key !== `${nextReleaseId}:${nextChangedSha256}`)
        .every(([, counter]) => counter.workerRequests === 0),
    ).toBe(true);

    const pinnedARead = await readVersionedHostAssetSha256(activeAHostFrame!, releaseId!, nextChangedPath!);
    expect(pinnedARead).toEqual({
      bytes: activeAChangedAsset!.bytes,
      sha256: activeAChangedAsset!.sha256,
      status: 200,
    });
    const activeAState = await owner.evaluate(() => {
      const demo = window.__officeDemo as {
        editor: { getState(): { status: string } } | null;
      };
      return demo.editor?.getState();
    });
    expect(activeAState?.status).toBe('ready');
    expect(findVersionedHostFrame(owner, releaseId!)).toBe(activeAHostFrame);

    const releaseBPage = await openNewStandaloneDocument(context, productionDocumentCases[1], failures, nextReleaseId!);
    expect(findVersionedHostFrame(releaseBPage, nextReleaseId!)).toBeTruthy();
    expect(countVersionedHostFrames(context)).toBe(2);
    expect(await readContentCounters(bootstrap)).toEqual(countersAfterIncrementalUpdate);

    await switchStableV5Pointer(
      bootstrap,
      stablePointer.body as {
        version: 1;
        releaseId: string;
        manifestUrl: string;
        manifestSha256: string;
      },
    );
    const rollback = await owner.evaluate(async () => {
      const manager = (
        globalThis as typeof globalThis & {
          __ONLYOFFICE_MATRIX_RESOURCE_MANAGER__: {
            checkForUpdates(): Promise<void>;
            plan(request: { scope: 'all' }): Promise<{
              downloadBytes: number;
              reusedBytes: number;
              releaseId: string;
            }>;
            apply(plan: unknown): Promise<void>;
            getSnapshot(): { installedRelease: string | null; readiness: string };
          };
        }
      ).__ONLYOFFICE_MATRIX_RESOURCE_MANAGER__;
      await manager.checkForUpdates();
      const plan = await manager.plan({ scope: 'all' });
      await manager.apply(plan);
      return { plan, snapshot: manager.getSnapshot() };
    });
    expect(rollback.plan).toMatchObject({ releaseId, downloadBytes: 0 });
    expect(rollback.snapshot).toMatchObject({ installedRelease: releaseId, readiness: 'ready' });
    expect(await readContentCounters(bootstrap)).toEqual(countersAfterIncrementalUpdate);
    const retainedBFrame = findVersionedHostFrame(releaseBPage, nextReleaseId!);
    expect(retainedBFrame, 'the running B editor must survive a stable-pointer rollback to A').toBeTruthy();
    expect(await readVersionedHostAssetSha256(retainedBFrame!, nextReleaseId!, nextChangedPath!)).toEqual({
      bytes: nextChangedBytes,
      sha256: nextChangedSha256,
      status: 200,
    });
    expect(
      await releaseBPage.evaluate(() => {
        const demo = window.__officeDemo as {
          editor: { getState(): { status: string } } | null;
        };
        return demo.editor?.getState().status;
      }),
    ).toBe('ready');

    const rollbackAPage = await openNewStandaloneDocument(context, productionDocumentCases[2], failures, releaseId!);
    expect(findVersionedHostFrame(rollbackAPage, releaseId!)).toBeTruthy();
    expect(countVersionedHostFrames(context)).toBe(3);
    expect(await readContentCounters(bootstrap)).toEqual(countersAfterIncrementalUpdate);
    await destroyStandaloneDocuments([releaseBPage, rollbackAPage]);

    const before = await waitForStableBroadcastCounts([owner, follower]);
    await owner.waitForTimeout(2_000);
    const after = await readBroadcastCounts([owner, follower]);
    expect(after).toEqual(before);
    expect(failures.filter((failure) => !isExpectedFailedObjectUrl(failure.url))).toEqual([]);

    const countersBeforeRestart = await readContentCounters(bootstrap);
    await context.close();

    context = await launchPersistentMatrixContext(profile, { forceDownloadSave: true });
    await Promise.all(context.pages().map((page) => page.close()));
    const contentRequestsAfterRestart: string[] = [];
    context.on('request', (request) => {
      const pathname = new URL(request.url()).pathname;
      if (
        pathname.startsWith('/objects/') ||
        pathname.startsWith('/segments/sha256/') ||
        pathname.startsWith('/blobs/sha256/')
      ) {
        contentRequestsAfterRestart.push(request.url());
      }
    });
    const restartFailures: BrowserFailure[] = [];
    await context.setOffline(true);

    const newDocumentPages = await Promise.all(
      productionDocumentCases.map((documentCase) => openNewStandaloneDocument(context, documentCase, restartFailures)),
    );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    for (let index = 0; index < newDocumentPages.length; index += 1) {
      await saveStandaloneDocument(
        newDocumentPages[index],
        productionDocumentCases[index],
        savedDocuments,
        `new-${productionDocumentCases[index].type}`,
      );
    }
    await destroyStandaloneDocuments(newDocumentPages);
    expect(context.pages()).toHaveLength(0);

    const existingFixturePaths = productionDocumentCases.map((documentCase) =>
      path.resolve(`public/fixtures/office/local.${documentCase.extension}`),
    );
    for (const fixturePath of existingFixturePaths) {
      expect(fs.existsSync(fixturePath), `Missing restart fixture ${fixturePath}`).toBe(true);
    }
    const existingDocumentPages = await Promise.all(
      productionDocumentCases.map((documentCase, index) =>
        openExistingStandaloneDocument(context, documentCase, existingFixturePaths[index], restartFailures),
      ),
    );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    await makeStandaloneSpreadsheetDirty(existingDocumentPages[1]);
    const editedSpreadsheetState = await existingDocumentPages[1].evaluate(() => {
      const demo = window.__officeDemo as {
        editor: { getState(): { dirty: boolean } } | null;
      };
      return demo.editor?.getState();
    });
    expect(editedSpreadsheetState?.dirty).toBe(true);
    const savedExistingPaths: string[] = [];
    for (let index = 0; index < existingDocumentPages.length; index += 1) {
      savedExistingPaths.push(
        await saveStandaloneDocument(
          existingDocumentPages[index],
          productionDocumentCases[index],
          savedDocuments,
          `existing-${productionDocumentCases[index].type}`,
        ),
      );
    }
    await destroyStandaloneDocuments(existingDocumentPages);
    expect(context.pages()).toHaveLength(0);

    const reopenedDocumentPages = await Promise.all(
      productionDocumentCases.map((documentCase, index) =>
        openExistingStandaloneDocument(context, documentCase, savedExistingPaths[index], restartFailures),
      ),
    );
    expect(context.pages()).toHaveLength(3);
    expect(countVersionedHostFrames(context)).toBe(3);
    for (let index = 0; index < reopenedDocumentPages.length; index += 1) {
      await saveStandaloneDocument(
        reopenedDocumentPages[index],
        productionDocumentCases[index],
        savedDocuments,
        `reopened-${productionDocumentCases[index].type}`,
      );
    }

    const restartedCdp = await context.newCDPSession(reopenedDocumentPages[0]);
    const restartedCanonicalCaches = await cacheInventory(restartedCdp, canonicalOrigin);
    const restartedContentCache = restartedCanonicalCaches.find((cache) => cache.cacheName === 'onlyoffice-content-v1');
    expect(restartedContentCache).toBeTruthy();
    const restartedContentUrls = new Set(restartedContentCache!.urls);
    for (const sha256 of expectedObjects.keys()) {
      expect(
        restartedContentUrls.has(`${canonicalOrigin}/__onlyoffice_content__/sha256/${sha256}`),
        `Restarted canonical cache should retain ${sha256}`,
      ).toBe(true);
    }
    const restartedEditorStorage = await Promise.all(
      reopenedDocumentPages.map(async (page, index) => {
        const frame = findVersionedHostFrame(page, releaseId!);
        expect(frame, `${productionEditorOrigins[index]} should restore its versioned frame`).toBeTruthy();
        return browserStorageInventory(frame!);
      }),
    );
    for (let index = 0; index < restartedEditorStorage.length; index += 1) {
      const storage = restartedEditorStorage[index];
      expect(storage.databases.map((database) => database.name)).not.toContain('onlyoffice-content-journal-v2');
      expect(
        storage.caches
          .flatMap((cache) => cache.entries)
          .every((entry) => isPermittedEditorShellCacheUrl(entry.url, [releaseId!, nextReleaseId!])),
        `${productionEditorOrigins[index]} must restore only its bounded shell cache`,
      ).toBe(true);
      expect(storage.usage ?? Number.POSITIVE_INFINITY).toBeLessThan(24 * 1024 * 1024);
    }
    expect(contentRequestsAfterRestart).toEqual([]);

    await context.setOffline(false);
    expect(await readContentCounters(reopenedDocumentPages[0])).toEqual(countersBeforeRestart);
    expect(contentRequestsAfterRestart).toEqual([]);
    expect(restartFailures).toEqual([]);
    await destroyStandaloneDocuments(reopenedDocumentPages);
    completed = true;
  } finally {
    await context.close().catch(() => undefined);
    if (completed && !retainedProfile) {
      fs.rmSync(profile, { force: true, recursive: true });
      fs.rmSync(savedDocuments, { force: true, recursive: true });
    } else {
      console.error(`Full-v5 Chromium profile retained for diagnosis: ${profile}`);
      console.error(`Full-v5 saved documents retained for diagnosis: ${savedDocuments}`);
    }
  }
});

test('a fixed editor origin can be reused immediately for a different document type', async ({ browser }) => {
  test.skip(
    matrixMode === 'full-v5',
    'The installed-context matrix above covers the production origins; slot release/reuse remains a lightweight unit contract.',
  );
  const context = await freshContext(browser);
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await cdp.send('Network.enable');
  const failures: BrowserFailure[] = [];
  const hostResponses: Array<{ fromDiskCache: boolean; fromServiceWorker: boolean; url: string }> = [];
  cdp.on('Network.responseReceived', ({ response, type }) => {
    if (type === 'Document' && response.url.includes('/office-host.html')) {
      hostResponses.push({
        fromDiskCache: Boolean(response.fromDiskCache),
        fromServiceWorker: Boolean(response.fromServiceWorker),
        url: response.url,
      });
    }
  });
  collectFailures(page, failures);
  const hostUrl = `${reusableEditorOrigin}/r/${releaseId}/office-host.html`;

  const openNewDocument = async (type: 'docx' | 'xlsx', editorPath: string) => {
    const query = new URLSearchParams({
      scenario: 'new-document',
      type,
      hostUrl,
    });
    await page.goto(`${canonicalOrigin}/save-e2e.html?${query}`);
    await page.waitForFunction(
      () => {
        const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
        return status?.ready === true || Boolean(status?.error);
      },
      null,
      { timeout: 3 * 60_000 },
    );
    const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__!.getStatus());
    expect(status.error).toBe('');
    expect(status.ready).toBe(true);
    await expect.poll(() => page.frames().some((frame) => frame.url().includes(editorPath))).toBe(true);
  };

  await openNewDocument('docx', '/documenteditor/main/index.html');
  await page.evaluate(async () => window.__ONLYOFFICE_SAVE_E2E__!.destroy());
  await openNewDocument('xlsx', '/spreadsheeteditor/main/index.html');

  expect(hostResponses).toHaveLength(2);
  expect(new URL(hostResponses[0].url).search).toBe('');
  expect(new URL(hostResponses[1].url).search).toBe('');
  expect(hostResponses[1].fromDiskCache || hostResponses[1].fromServiceWorker).toBe(true);
  expect(failures).toEqual([]);
  await context.close();
});

for (const [type, editorPath, extension] of [
  ['docx', '/documenteditor/main/index.html', 'docx'],
  ['xlsx', '/spreadsheeteditor/main/index.html', 'xlsx'],
  ['pptx', '/presentationeditor/main/index.html', 'pptx'],
] as const) {
  test(`new ${type} opens and converts for save through the immutable Host`, async ({ browser }) => {
    test.skip(
      matrixMode === 'full-v5',
      'Word, spreadsheet, and presentation are exercised concurrently after the one real v5 installation above.',
    );
    const context = await freshContext(browser);
    const page = await context.newPage();
    const failures: BrowserFailure[] = [];
    collectFailures(page, failures);
    const hostUrl = `${editorOrigin}/r/${releaseId}/office-host.html`;
    const query = new URLSearchParams({
      scenario: 'new-document',
      type,
      hostUrl,
    });
    await page.goto(`${canonicalOrigin}/save-e2e.html?${query}`);
    await page.waitForFunction(
      () => {
        const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
        return status?.ready === true || Boolean(status?.error);
      },
      null,
      { timeout: 3 * 60_000 },
    );
    const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__!.getStatus());
    expect(status.error).toBe('');
    expect(status.ready).toBe(true);
    await expect.poll(() => page.frames().some((frame) => frame.url().includes(editorPath))).toBe(true);

    const saved = await page.evaluate(async () => window.__ONLYOFFICE_SAVE_E2E__!.save());
    expect(saved.fileName).toMatch(new RegExp(`\\.${extension}$`, 'i'));
    expect(saved.size).toBeGreaterThan(0);
    expect(saved.firstBytes.slice(0, 4)).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(failures).toEqual([]);
    await context.close();
  });
}
