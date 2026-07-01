import { expect, test } from '@playwright/test';

async function waitForSaveE2EReady(page: import('@playwright/test').Page): Promise<void> {
  await page.waitForFunction(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus().ready === true, null, {
    timeout: 90_000,
  });
  const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus());
  expect(status?.error).toBe('');
}

test('legacy xls built-in Download As PDF triggers a browser download', async ({ page }) => {
  await page.goto('/save-e2e.html?scenario=local-file&type=xls');
  await waitForSaveE2EReady(page);

  const editorFrame = page.frames().find((frame) => frame.url().includes('/spreadsheeteditor/'));
  expect(editorFrame, 'spreadsheet editor frame').toBeTruthy();

  await editorFrame!.locator('text=File').first().click({ timeout: 10_000 });
  await editorFrame!.locator('.svg-format-pdf').first().locator('xpath=..').click({ timeout: 10_000, force: true });
  await editorFrame!.getByText(/Save & Download|保存并下载|保存 & 下载/i).first().waitFor({
    state: 'visible',
    timeout: 10_000,
  });

  const downloadPromise = page.waitForEvent('download', { timeout: 45_000 });
  await editorFrame!.getByText(/Save & Download|保存并下载|保存 & 下载/i).first().click({
    timeout: 10_000,
    force: true,
  });
  const download = await downloadPromise;

  expect(download.suggestedFilename()).toBe('legacy.pdf');
  expect(await download.failure()).toBeNull();

  const stream = await download.createReadStream();
  expect(stream).toBeTruthy();

  let size = 0;
  let header = Buffer.alloc(0);
  for await (const chunk of stream!) {
    if (header.length < 8) {
      header = Buffer.concat([header, chunk]).subarray(0, 8);
    }
    size += chunk.length;
  }

  expect(size).toBeGreaterThan(0);
  expect(Array.from(header.slice(0, 4))).toEqual([0x25, 0x50, 0x44, 0x46]);
});
