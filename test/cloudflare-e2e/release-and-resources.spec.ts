import { expect, type Browser, type BrowserContext, type Page, test } from '@playwright/test';

const releaseId = process.env.ONLYOFFICE_CF_MATRIX_RELEASE_ID;
const packageVersion = process.env.ONLYOFFICE_CF_MATRIX_PACKAGE_VERSION;
const port = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_PORT || '8787', 10);
const canonicalOrigin = `http://onlyoffice.localhost:${port}`;
const editorOrigin = `http://aries.localhost:${port}`;

type ReleaseManifest = {
  version: number;
  releaseId: string;
  packageVersion: string;
  package: {
    path: string;
    bytes: number;
    sha256: string;
    segments: Array<{ id: string; offset: number; bytes: number; sha256: string }>;
  };
  assets: Array<{ path: string; mime: string; bytes: number; packageOffset: number }>;
};

type BrowserFailure = { source: string; message: string };

async function waitForStableProgress(page: Page, stableFor = 1_500, timeout = 15_000): Promise<string | null> {
  const progress = page.getByRole('progressbar');
  const startedAt = Date.now();
  let lastValue = await progress.getAttribute('aria-valuenow');
  let stableSince = Date.now();
  while (Date.now() - startedAt < timeout) {
    await page.waitForTimeout(250);
    const currentValue = await progress.getAttribute('aria-valuenow');
    if (currentValue !== lastValue) {
      lastValue = currentValue;
      stableSince = Date.now();
    } else if (Date.now() - stableSince >= stableFor) {
      return currentValue;
    }
  }
  throw new Error(`Paused resource progress did not settle within ${timeout} ms`);
}

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

function collectFailures(page: Page, failures: BrowserFailure[]) {
  page.on('pageerror', (error) => failures.push({ source: 'pageerror', message: error.message }));
  page.on('console', (message) => {
    if (message.type() === 'error') failures.push({ source: 'console', message: message.text() });
  });
}

async function freshContext(browser: Browser): Promise<BrowserContext> {
  const context = await browser.newContext({
    locale: 'en-US',
    serviceWorkers: 'allow',
    acceptDownloads: true,
  });
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
  return context;
}

async function browserFetch<T>(
  page: Page,
  url: string,
  init?: { method?: string; headers?: Record<string, string>; cache?: RequestCache },
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

test('Release v4 routes reproduce production MIME, cache, range, and immutable Office Pack behavior', async ({
  browser,
}) => {
  expect(releaseId).toBeTruthy();
  const context = await freshContext(browser);
  const page = await context.newPage();
  await page.goto(`${canonicalOrigin}/`);

  const pointer = await browserFetch<{ version: number; releaseId: string }>(
    page,
    `${canonicalOrigin}/channels/stable.json`,
  );
  expect(pointer.status).toBe(200);
  expect(pointer.headers['cache-control']).toBe('no-store');
  expect(pointer.body).toEqual({ version: 1, releaseId });

  const manifestResult = await browserFetch<ReleaseManifest>(
    page,
    `${canonicalOrigin}/releases/${releaseId}/manifest.json`,
  );
  expect(manifestResult.status).toBe(200);
  expect(manifestResult.body.version).toBe(4);
  expect(manifestResult.body.releaseId).toBe(releaseId);
  expect(manifestResult.body.packageVersion).toBe(packageVersion);
  expect(manifestResult.body.package.path).toBe('office-resources.oobpack');

  const packMagic = await browserFetch<string>(page, `${canonicalOrigin}/p/${releaseId}/office-resources.oobpack`, {
    headers: { Range: 'bytes=0-7' },
  });
  expect(packMagic.status).toBe(206);
  expect(packMagic.body).toBe('OOBPACK1');
  expect(packMagic.headers['content-type']).toBe('application/vnd.onlyoffice.browser-pack');

  const host = await browserFetch<string>(page, `${editorOrigin}/r/${releaseId}/office-host.html`, { method: 'HEAD' });
  expect(host.status).toBe(200);
  expect(host.headers['content-type']).toBe('text/html; charset=utf-8');
  expect(host.headers['cache-control']).toContain('immutable');
  expect(host.headers['x-onlyoffice-asset-version']).toBe(releaseId);

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

test('two tabs pause, resume, finish resources without broadcast ping-pong', async ({ browser }) => {
  const context = await freshContext(browser);
  await context.route('**/p/**', async (route) => {
    await new Promise((resolve) => setTimeout(resolve, 40));
    await route.continue();
  });
  const failures: BrowserFailure[] = [];
  const appUrl = `${canonicalOrigin}/?resourcePrefetch=manual&hostUrl=${encodeURIComponent(
    `${editorOrigin}/office-host.html`,
  )}`;
  const owner = await context.newPage();
  const follower = await context.newPage();
  collectFailures(owner, failures);
  collectFailures(follower, failures);
  await Promise.all([owner.goto(appUrl), follower.goto(appUrl)]);
  await Promise.all([
    expect(owner.locator('#resource-button')).toBeVisible(),
    expect(follower.locator('#resource-button')).toBeVisible(),
  ]);

  await owner.locator('#resource-button').click();
  const dialog = owner.locator('#resource-dialog');
  const dialogBox = await dialog.boundingBox();
  const viewport = owner.viewportSize();
  expect(dialogBox).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(Math.abs(dialogBox!.x + dialogBox!.width / 2 - viewport!.width / 2)).toBeLessThanOrEqual(1);
  expect(Math.abs(dialogBox!.y + dialogBox!.height / 2 - viewport!.height / 2)).toBeLessThanOrEqual(1);
  await owner.getByRole('button', { name: 'Install complete package' }).click();
  const pause = owner.getByRole('button', { name: 'Pause' });
  await expect(pause).toBeVisible({ timeout: 90_000 });
  await expect(owner.getByText('Stage 2/4', { exact: true })).toBeVisible();
  await expect(owner.getByText(/Package segment \d+\/\d+/, { exact: true })).toBeVisible();
  await pause.click();
  await expect(owner.getByText('Paused', { exact: true }).first()).toBeVisible();
  const pausedValue = await waitForStableProgress(owner);
  await owner.waitForTimeout(750);
  expect(await owner.getByRole('progressbar').getAttribute('aria-valuenow')).toBe(pausedValue);

  await owner.getByRole('button', { name: 'Resume' }).click();
  await expect(owner.getByText('Resources ready', { exact: true }).first()).toBeVisible({
    timeout: 6 * 60_000,
  });
  await expect(follower.locator('#resource-button')).toContainText('Resources ready', {
    timeout: 30_000,
  });
  await expect(owner.getByRole('progressbar')).toHaveCount(0);
  await follower.locator('#resource-button').click();
  await expect(follower.getByRole('progressbar')).toHaveCount(0);

  await context.setOffline(true);
  const offlineMagic = await owner.evaluate(async (url) => {
    const response = await caches.match(url);
    if (!response) return { status: 0, body: '' };
    const bytes = new Uint8Array(await response.arrayBuffer());
    return {
      status: response.status,
      body: new TextDecoder().decode(bytes.subarray(0, 8)),
    };
  }, `${canonicalOrigin}/p/${releaseId}/office-resources.oobpack?segment=segment-001`);
  expect(offlineMagic.status).toBe(200);
  expect(offlineMagic.body).toBe('OOBPACK1');
  await context.setOffline(false);

  const before = await waitForStableBroadcastCounts([owner, follower]);
  await owner.waitForTimeout(2_000);
  const after = await readBroadcastCounts([owner, follower]);
  expect(after).toEqual(before);
  expect(failures).toEqual([]);
  await context.close();
});

for (const [type, editorPath, extension] of [
  ['docx', '/documenteditor/main/index.html', 'docx'],
  ['xlsx', '/spreadsheeteditor/main/index.html', 'xlsx'],
  ['pptx', '/presentationeditor/main/index.html', 'pptx'],
] as const) {
  test(`new ${type} opens and converts for save through the immutable Host`, async ({ browser }) => {
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
