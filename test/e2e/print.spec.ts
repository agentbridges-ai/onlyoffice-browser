import { expect, type ElementHandle, type Page, test } from '@playwright/test';

type SaveE2EStatus = {
  ready: boolean;
  error: string;
};

type PrintTargetType = 'xlsx' | 'xls' | 'docx' | 'doc' | 'pptx' | 'ppt';

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
  expect((await getStatus(page)).error).toBe('');
}

async function findPrintButton(page: Page): Promise<ElementHandle<HTMLElement> | null> {
  for (const frame of page.frames()) {
    const handle = await frame
      .evaluateHandle(() => {
        const isVisible = (element: HTMLElement) => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
        };
        const isPrintCandidate = (element: HTMLElement) => {
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
          return /(^|\s)(print|打印)(\s|$)/i.test(label) || /Print file/i.test(label);
        };
        const candidates = Array.from(document.querySelectorAll<HTMLElement>('button, [role="button"], a'))
          .filter(isVisible)
          .filter(isPrintCandidate)
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

async function clickNativePrint(page: Page): Promise<void> {
  const button = await findPrintButton(page);
  expect(button, 'OnlyOffice built-in print button').not.toBeNull();
  await button!.click();
}

async function findPrintSettingsConfirmButton(page: Page): Promise<ElementHandle<HTMLElement> | null> {
  for (const frame of page.frames()) {
    const handle = await frame
      .evaluateHandle(() => {
        const isVisible = (element: HTMLElement) => {
          const style = window.getComputedStyle(element);
          const rect = element.getBoundingClientRect();
          return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
        };
        const candidates = Array.from(document.querySelectorAll<HTMLElement>('button, [role="button"]'))
          .filter(isVisible)
          .filter((element) => (element.textContent || '').trim().toLowerCase() === 'print')
          .sort((left, right) => {
            const leftRect = left.getBoundingClientRect();
            const rightRect = right.getBoundingClientRect();
            return rightRect.top - leftRect.top || rightRect.left - leftRect.left;
          });
        return candidates[0] || null;
      })
      .catch(() => null);
    const element = handle?.asElement() as ElementHandle<HTMLElement> | null | undefined;
    if (element) return element;
  }
  return null;
}

async function getPrintFrameInfo(page: Page): Promise<{ src: string; frameUrl: string } | null> {
  for (const frame of page.frames()) {
    const info = await frame
      .evaluate(() => {
        const iframe = document.querySelector<HTMLIFrameElement>('#id-print-frame');
        const src = iframe?.src || '';
        return src ? { src, frameUrl: window.location.href } : null;
      })
      .catch(() => null);
    if (info) return info;
  }
  return null;
}

async function runNativePrintScenario(page: Page, type: PrintTargetType): Promise<void> {
  const failures = collectPageFailures(page);

  await page.goto(`/save-e2e.html?scenario=local-file&type=${type}`);
  await waitForReady(page);

  await expect
    .poll(
      () => {
        const frame = page.frames().find((candidate) => candidate.url().includes('/office-host.html'));
        return frame?.url() || '';
      },
      { timeout: 30_000 },
    )
    .not.toBe('');
  const frame = page.frames().find((candidate) => candidate.url().includes('/office-host.html'));
  expect(frame, 'office host frame').toBeTruthy();

  await clickNativePrint(page);
  if (!(await getPrintFrameInfo(page))) {
    const nextPrintButton = (await findPrintSettingsConfirmButton(page)) || (await findPrintButton(page));
    if (nextPrintButton) {
      await nextPrintButton.click();
    }
  }

  await expect.poll(() => getPrintFrameInfo(page).then((info) => info?.src || ''), { timeout: 45_000 }).toMatch(/^blob:/);
  const printFrameInfo = await getPrintFrameInfo(page);
  expect(printFrameInfo).not.toBeNull();

  const pdfInfo = await frame!.evaluate(async (url) => {
    const response = await fetch(url);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const header = Array.from(bytes.slice(0, 4), (byte) => String.fromCharCode(byte)).join('');
    return {
      ok: response.ok,
      contentType: response.headers.get('content-type') || '',
      byteLength: bytes.byteLength,
      header,
    };
  }, printFrameInfo!.src);

  expect(pdfInfo.ok).toBe(true);
  expect(pdfInfo.contentType).toContain('application/pdf');
  expect(pdfInfo.byteLength).toBeGreaterThan(0);
  expect(pdfInfo.header).toBe('%PDF');
  expect(failures).toEqual([]);
}

for (const type of ['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt'] as const) {
  test(`local ${type} print uses the built-in print flow and creates a PDF print frame`, async ({ page }) => {
    await runNativePrintScenario(page, type);
  });
}
