import { expect, type Page, test } from '@playwright/test';

type BurstStatus = {
  mounted: number;
  ready: number;
  errors: string[];
  inFlight: number;
  maxInFlight: number;
  retries: number;
  uniqueOrigins: number;
  outerFrames: number;
  done: boolean;
};

test.describe.configure({ mode: 'serial', timeout: 11 * 60_000 });
test.use({ trace: 'off', video: 'off' });

const burstCount = Math.max(1, Number.parseInt(process.env.OFFICE_BURST_COUNT || '3', 10));

function collectPageFailures(page: Page): string[] {
  const failures: string[] = [];
  page.on('pageerror', (error) => failures.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') failures.push(message.text());
  });
  return failures;
}

async function getStatus(page: Page): Promise<BurstStatus> {
  return page.evaluate(() => window.__ONLYOFFICE_BURST_E2E__!.getStatus());
}

test('loads and cleans up a mixed real-file burst with isolated hosts', async ({ page }) => {
  test.setTimeout(11 * 60_000);
  const failures = collectPageFailures(page);
  await page.goto('/burst-e2e.html');
  await page.evaluate(
    (count) =>
      window.__ONLYOFFICE_BURST_E2E__!.start({
        count,
        intervalMs: 10,
        activationBudget: 1,
      }),
    burstCount,
  );

  await expect.poll(() => getStatus(page).then((value) => value.mounted), { timeout: 30_000 }).toBe(burstCount);
  const mounted = await getStatus(page);
  expect(mounted.outerFrames).toBe(burstCount);
  expect(mounted.uniqueOrigins).toBe(burstCount);

  await page.waitForFunction(() => window.__ONLYOFFICE_BURST_E2E__?.getStatus().done === true, null, {
    timeout: 10 * 60_000,
  });
  const ready = await getStatus(page);
  expect(ready.errors).toEqual([]);
  expect(ready.ready).toBe(burstCount);
  expect(ready.maxInFlight).toBe(1);
  expect(ready.retries).toBeLessThanOrEqual(burstCount);

  const hostFrames = page.frames().filter((frame) => frame.url().includes('/office-host.html'));
  expect(hostFrames).toHaveLength(burstCount);
  for (const hostFrame of hostFrames) {
    await expect(hostFrame.locator('iframe[name="frameEditor"]')).toHaveCount(1, { timeout: 30_000 });
  }

  await page.evaluate(() => window.__ONLYOFFICE_BURST_E2E__!.closeAll());
  await expect.poll(() => getStatus(page).then((value) => value.outerFrames), { timeout: 30_000 }).toBe(0);
  expect(failures).toEqual([]);
});

declare global {
  interface Window {
    __ONLYOFFICE_BURST_E2E__?: {
      start(options?: { count?: number; intervalMs?: number; activationBudget?: number }): void;
      wait(): Promise<void>;
      closeAll(): Promise<void>;
      getStatus(): BurstStatus;
    };
  }
}
