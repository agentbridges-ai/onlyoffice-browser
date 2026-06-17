import { expect, test } from '@playwright/test';

test('demo host loads without page errors', async ({ page }) => {
  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));

  await page.goto('/');

  await expect(page.locator('#app')).toBeVisible();
  await expect(page.locator('#new-word-button')).toBeVisible();
  await expect(page.locator('#upload-button')).toBeVisible();
  await expect(page.locator('input[name="open-mode"][value="edit"]')).toBeChecked();
  await expect(page.locator('input[name="open-mode"][value="readonly"]')).toBeAttached();
  await expect(page.locator('input[name="open-mode"][value="preview"]')).toBeAttached();
  await expect(page.locator('#editor-grid')).toBeAttached();
  expect(pageErrors).toEqual([]);
});

test('OnlyOffice root discovery assets are reachable', async ({ request }) => {
  const serviceWorker = await request.get('/document_editor_service_worker.js');
  expect(serviceWorker.ok()).toBe(true);

  const plugins = await request.get('/plugins.json');
  expect(plugins.ok()).toBe(true);

  const themes = await request.get('/themes.json');
  expect(themes.ok()).toBe(true);
});
