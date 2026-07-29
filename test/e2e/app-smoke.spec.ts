import { expect, test } from '@playwright/test';

test('demo host loads without page errors', async ({ page }) => {
  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));

  await page.goto('/');

  await expect(page.locator('#app')).toBeVisible();
  await expect(page.locator('#resource-button')).toBeVisible();
  await expect(page.locator('#open-button')).toBeVisible();
  await expect(page.locator('.new-menu > summary')).toBeVisible();
  await expect(page.locator('[data-new="docx"]')).toBeAttached();
  await expect(page.locator('[data-new="xlsx"]')).toBeAttached();
  await expect(page.locator('[data-new="pptx"]')).toBeAttached();
  await expect(page.locator('#empty-open-button')).toBeVisible();
  await expect(page.locator('#editor-panel')).toBeAttached();
  await expect(page.locator('#dirty-dialog [value="cancel"]')).toHaveCount(1);
  await expect(page.locator('#dirty-dialog [value="discard"]')).toHaveCount(1);
  await expect(page.locator('#dirty-dialog [value="save"]')).toHaveCount(1);
  expect(pageErrors).toEqual([]);
});

test('OnlyOffice root discovery assets are reachable', async ({ request }) => {
  const serviceWorker = await request.get('/document_editor_service_worker.js');
  expect(serviceWorker.ok()).toBe(true);

  const plugins = await request.get('/plugins.json');
  expect(plugins.ok()).toBe(true);

  const themes = await request.get('/themes.json');
  expect(themes.ok()).toBe(true);

  const host = await request.get('/office-host.html');
  expect(host.ok()).toBe(true);
});
