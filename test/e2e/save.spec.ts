import { expect, type ElementHandle, type Page, test } from '@playwright/test';

type SaveButtonState = {
  disabled: boolean;
  frameUrl: string;
  label: string;
  title: string;
  className: string;
};

type SaveE2EStatus = {
  ready: boolean;
  dirty: boolean;
  writeCount: number;
  lastHash: string;
  lastSize: number;
  initialHash: string;
  initialSize: number;
  error: string;
};

type SaveE2EResult = {
  fileName: string;
  size: number;
  hash: string;
  firstBytes: number[];
};

test.describe.configure({ mode: 'serial', timeout: 120_000 });

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
    const api = (window as Window & { __ONLYOFFICE_SAVE_E2E__?: { getStatus: () => SaveE2EStatus } }).__ONLYOFFICE_SAVE_E2E__;
    if (!api) throw new Error('Save E2E controller is not installed');
    return api.getStatus();
  });
}

async function waitForReady(page: Page): Promise<void> {
  await page.waitForFunction(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus().ready === true, null, {
    timeout: 90_000,
  });
  const status = await getStatus(page);
  expect(status.error).toBe('');
}

async function findSaveButton(page: Page): Promise<ElementHandle<HTMLElement> | null> {
  for (const frame of page.frames()) {
    const handle = await frame
      .evaluateHandle(() => {
        const isVisible = (element: HTMLElement) => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
        };
        const isSaveCandidate = (element: HTMLElement) => {
          const label = [
            element.getAttribute('aria-label'),
            element.getAttribute('title'),
            element.getAttribute('data-hint'),
            element.getAttribute('data-tooltip'),
            element.textContent,
          ]
            .filter(Boolean)
            .join(' ')
            .trim();
          if (!label) return false;
          return /(^|\s)(save|保存)(\s|$)/i.test(label);
        };
        const candidates = Array.from(document.querySelectorAll<HTMLElement>('button, [role="button"], a'))
          .filter(isVisible)
          .filter(isSaveCandidate)
          .sort((left, right) => {
            const leftRect = left.getBoundingClientRect();
            const rightRect = right.getBoundingClientRect();
            return leftRect.top - rightRect.top || leftRect.left - rightRect.left;
          });
        return candidates[0] || null;
      })
      .catch(() => null);
    const element = handle?.asElement() as ElementHandle<HTMLElement> | null | undefined;
    if (element) return element;
  }
  return null;
}

async function getSaveButtonState(page: Page): Promise<SaveButtonState | null> {
  const button = await findSaveButton(page);
  if (!button) return null;
  const frame = await button.ownerFrame();
  return button.evaluate(
    (element, frameUrl) => {
      const className = typeof element.className === 'string' ? element.className : '';
      const disabled =
        (element instanceof HTMLButtonElement && element.disabled) ||
        element.getAttribute('aria-disabled') === 'true' ||
        element.hasAttribute('disabled') ||
        /\b(disabled|disable|inactive)\b/i.test(className) ||
        Boolean(element.closest('.disabled, .disable, [aria-disabled="true"], [disabled]'));
      return {
        disabled,
        frameUrl,
        label: element.textContent?.trim() || '',
        title: element.getAttribute('title') || element.getAttribute('aria-label') || element.getAttribute('data-hint') || '',
        className,
      };
    },
    frame?.url() || '',
  );
}

async function clickNativeSave(page: Page): Promise<void> {
  const button = await findSaveButton(page);
  expect(button, 'OnlyOffice built-in save button').not.toBeNull();
  await button!.click();
}

async function makeSpreadsheetDirty(page: Page): Promise<void> {
  const editor = page.locator('#save-e2e-editor');
  await expect(editor).toBeVisible();
  const box = await editor.boundingBox();
  if (!box) throw new Error('Editor box is not visible');

  for (let attempt = 0; attempt < 4; attempt += 1) {
    await page.mouse.click(box.x + Math.min(260 + attempt * 28, box.width - 80), box.y + Math.min(220 + attempt * 24, box.height - 80));
    await page.keyboard.type(`save-e2e-${Date.now()}-${attempt}`);
    await page.keyboard.press('Enter');
    try {
      await page.waitForFunction(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus().dirty === true, null, {
        timeout: 8_000,
      });
      return;
    } catch {
      // Try another grid coordinate; spreadsheet focus can vary by toolbar height.
    }
  }

  throw new Error('Timed out waiting for spreadsheet dirty state after editing');
}

test('local xlsx saves through callback without downloads or autosave', async ({ page }) => {
  const failures = collectPageFailures(page);
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));

  await page.goto('/save-e2e.html?scenario=local-file&type=xlsx');
  await waitForReady(page);

  await expect(page.locator('#save-e2e-root > button, #save-e2e-root [data-action="save"]')).toHaveCount(0);
  await expect
    .poll(() => getSaveButtonState(page).then((state) => state?.disabled), { timeout: 20_000 })
    .toBe(true);

  const initial = await getStatus(page);
  expect(initial.writeCount).toBe(0);
  expect(initial.initialHash).toBeTruthy();
  expect(initial.lastHash).toBe(initial.initialHash);

  await makeSpreadsheetDirty(page);
  await expect
    .poll(() => getSaveButtonState(page).then((state) => state?.disabled), { timeout: 20_000 })
    .toBe(false);

  await page.waitForTimeout(2_000);
  expect((await getStatus(page)).writeCount).toBe(0);

  await clickNativeSave(page);
  await expect.poll(() => getStatus(page).then((status) => status.writeCount), { timeout: 45_000 }).toBe(1);

  const saved = await getStatus(page);
  expect(saved.lastSize).toBeGreaterThan(0);
  expect(saved.lastHash).not.toBe(initial.initialHash);
  expect(downloads).toEqual([]);
  expect(failures).toEqual([]);
});

test('new xlsx saves through browser download without callback writes', async ({ page }) => {
  const failures = collectPageFailures(page);
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));

  await page.goto('/save-e2e.html?scenario=new-document&type=xlsx');
  await waitForReady(page);

  await expect(page.locator('#save-e2e-root > button, #save-e2e-root [data-action="save"]')).toHaveCount(0);
  await makeSpreadsheetDirty(page);
  await expect
    .poll(() => getSaveButtonState(page).then((state) => state?.disabled), { timeout: 20_000 })
    .toBe(false);

  const downloadPromise = page.waitForEvent('download', { timeout: 45_000 });
  await clickNativeSave(page);
  const download = await downloadPromise;

  expect(download.suggestedFilename()).toMatch(/\.xlsx$/i);
  expect((await getStatus(page)).writeCount).toBe(0);
  expect(downloads).toHaveLength(1);
  expect(failures).toEqual([]);
});

const legacySaveOutputs = {
  doc: 'docx',
  ppt: 'pptx',
  xls: 'xlsx',
} as const;

for (const [type, outputType] of Object.entries(legacySaveOutputs) as Array<
  [keyof typeof legacySaveOutputs, (typeof legacySaveOutputs)[keyof typeof legacySaveOutputs]]
>) {
  test(`local ${type} saves through callback as ${outputType} without downloads`, async ({ page }) => {
    const failures = collectPageFailures(page);
    const downloads: string[] = [];
    page.on('download', (download) => downloads.push(download.suggestedFilename()));

    await page.goto(`/save-e2e.html?scenario=local-file&type=${type}`);
    await waitForReady(page);

    const result = await page.evaluate(
      async (targetExt): Promise<SaveE2EResult> => {
        const api = (
          window as Window & {
            __ONLYOFFICE_SAVE_E2E__?: {
              save: (targetExt?: string) => Promise<SaveE2EResult>;
            };
          }
        ).__ONLYOFFICE_SAVE_E2E__;
        if (!api) throw new Error('Save E2E controller is not installed');
        return api.save(targetExt);
      },
      type.toUpperCase(),
    );

    expect(result.fileName).toMatch(new RegExp(`\\.${outputType}$`, 'i'));
    expect(result.size).toBeGreaterThan(0);
    expect(result.hash).toBeTruthy();
    expect(result.firstBytes.slice(0, 4)).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect((await getStatus(page)).writeCount).toBe(1);
    expect(downloads).toEqual([]);
    expect(failures).toEqual([]);
  });
}
