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

async function getPrintFrameInfo(
  page: Page,
): Promise<{ id: string; src: string; frameUrl: string; href: string; canPrint: boolean; accessError: string } | null> {
  for (const frame of page.frames()) {
    const info = await frame
      .evaluate(() => {
        const iframe = document.querySelector<HTMLIFrameElement>('#id-print-frame');
        const src = iframe?.src || '';
        if (!src) return null;

        let href = '';
        let canPrint = false;
        let accessError = '';
        try {
          href = iframe?.contentWindow?.location.href || '';
          canPrint = typeof iframe?.contentWindow?.print === 'function';
        } catch (error) {
          accessError = error instanceof Error ? `${error.name}: ${error.message}` : String(error);
        }

        return { id: iframe?.id || '', src, frameUrl: window.location.href, href, canPrint, accessError };
      })
      .catch(() => null);
    if (info) return info;
  }
  return null;
}

async function runNativePrintScenario(page: Page, type: PrintTargetType): Promise<void> {
  const failures = collectPageFailures(page);
  const popups: string[] = [];
  page.on('popup', async (popup) => {
    popups.push(popup.url());
    await popup.close().catch(() => undefined);
  });

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

  await expect
    .poll(() => getPrintFrameInfo(page).then((info) => info?.src || ''), { timeout: 45_000 })
    .toContain('/__onlyoffice-browser-print__/');
  const printFrameInfo = await getPrintFrameInfo(page);
  expect(printFrameInfo).not.toBeNull();
  expect(printFrameInfo!.src).not.toMatch(/^blob:/);
  expect(printFrameInfo!.src).toMatch(/\.pdf$/);

  const pdfInfo = await frame!.evaluate(async (url) => {
    const headResponse = await fetch(url, { method: 'HEAD' });
    const rangeResponse = await fetch(url, { headers: { Range: 'bytes=0-3' } });
    const rangeBytes = new Uint8Array(await rangeResponse.arrayBuffer());
    const response = await fetch(url);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const header = Array.from(bytes.slice(0, 4), (byte) => String.fromCharCode(byte)).join('');
    const rangeHeader = Array.from(rangeBytes, (byte) => String.fromCharCode(byte)).join('');
    const decoder = new TextDecoder('latin1');
    const source = decoder.decode(bytes);
    const explicitZeroPageTree = (source.match(/\d+\s+\d+\s+obj[\s\S]*?endobj/g) || []).some(
      (objectSource) =>
        /\/Type\s*\/Pages\b/.test(objectSource) &&
        /\/Count\s+0\b/.test(objectSource) &&
        /\/Kids\s*\[\s*\]/.test(objectSource),
    );
    return {
      ok: response.ok,
      acceptRanges: response.headers.get('accept-ranges') || '',
      contentDisposition: response.headers.get('content-disposition') || '',
      contentType: response.headers.get('content-type') || '',
      byteLength: bytes.byteLength,
      header,
      explicitZeroPageTree,
      headOk: headResponse.ok,
      headAcceptRanges: headResponse.headers.get('accept-ranges') || '',
      headContentLength: headResponse.headers.get('content-length') || '',
      rangeOk: rangeResponse.ok,
      rangeStatus: rangeResponse.status,
      rangeContentRange: rangeResponse.headers.get('content-range') || '',
      rangeHeader,
    };
  }, printFrameInfo!.src);

  expect(pdfInfo.ok).toBe(true);
  expect(pdfInfo.acceptRanges).toBe('bytes');
  expect(pdfInfo.contentDisposition).toContain('inline;');
  expect(pdfInfo.contentType).toContain('application/pdf');
  expect(pdfInfo.byteLength).toBeGreaterThan(0);
  expect(pdfInfo.header).toBe('%PDF');
  expect(pdfInfo.explicitZeroPageTree).toBe(false);
  expect(pdfInfo.headOk).toBe(true);
  expect(pdfInfo.headAcceptRanges).toBe('bytes');
  expect(Number(pdfInfo.headContentLength)).toBeGreaterThan(0);
  expect(pdfInfo.rangeOk).toBe(true);
  expect(pdfInfo.rangeStatus).toBe(206);
  expect(pdfInfo.rangeContentRange).toMatch(/^bytes 0-3\/\d+$/);
  expect(pdfInfo.rangeHeader).toBe('%PDF');
  expect(printFrameInfo!.id).toBe('id-print-frame');
  expect(popups).toEqual([]);
  expect(failures).toEqual([]);
}

for (const type of ['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt'] as const) {
  test(`local ${type} print uses the built-in print flow and creates a PDF print frame`, async ({ page }) => {
    await runNativePrintScenario(page, type);
  });
}
