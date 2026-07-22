import { expect, type Page, test } from '@playwright/test';

const FIXTURE_URL = '/fixtures/regressions/example-document-title-ole.doc';
const FIXTURE_NAME = 'Example Document Title.doc';
const FIXTURE_SHA256 = 'd85e44ae5368ccbbe57ded8533ced05a250c30cfa15da10f19fdaf63f080238c';

type SaveE2EStatus = {
  type: string;
  ready: boolean;
  initialHash: string;
  initialSize: number;
  error: string;
  state: { status: string } | null;
};

test('legacy DOC with embedded OpenDocument chart reaches READY without a conversion error', async ({ page }) => {
  const failures = collectPageFailures(page);
  const params = new URLSearchParams({
    scenario: 'local-file',
    type: 'doc',
    fixtureUrl: FIXTURE_URL,
    fixtureName: FIXTURE_NAME,
  });

  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 90_000 },
  );

  const status = await getStatus(page);
  expect(status.error).toBe('');
  expect(status.type).toBe('doc');
  expect(status.ready).toBe(true);
  expect(status.state?.status).toBe('ready');
  expect(status.initialSize).toBeGreaterThan(0);
  expect(status.initialHash).toBe(FIXTURE_SHA256);
  expect(failures).toEqual([]);
});

function collectPageFailures(page: Page): string[] {
  const failures: string[] = [];
  page.on('pageerror', (error) => failures.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') failures.push(message.text());
  });
  page.on('dialog', async (dialog) => {
    failures.push(`Unexpected dialog: ${dialog.message()}`);
    await dialog.dismiss().catch(() => undefined);
  });
  return failures;
}

async function getStatus(page: Page): Promise<SaveE2EStatus> {
  return page.evaluate(() => {
    const api = (
      window as Window & {
        __ONLYOFFICE_SAVE_E2E__?: { getStatus: () => SaveE2EStatus };
      }
    ).__ONLYOFFICE_SAVE_E2E__;
    if (!api) throw new Error('Save E2E controller is not installed');
    return api.getStatus();
  });
}
