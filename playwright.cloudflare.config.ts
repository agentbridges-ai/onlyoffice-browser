import { defineConfig, devices } from '@playwright/test';

const port = Number.parseInt(process.env.ONLYOFFICE_CF_MATRIX_PORT || '8787', 10);
const canonicalOrigin = `http://onlyoffice.localhost:${port}`;

export default defineConfig({
  testDir: 'test/cloudflare-e2e',
  timeout: 8 * 60_000,
  expect: {
    timeout: 15_000,
  },
  fullyParallel: false,
  workers: 1,
  reporter: [['list'], ['html', { outputFolder: 'playwright-report-cloudflare', open: 'never' }]],
  use: {
    baseURL: canonicalOrigin,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    launchOptions: {
      args: ['--no-proxy-server'],
    },
  },
  projects: [
    {
      name: 'chromium-cloudflare',
      use: { ...devices['Desktop Chrome'], viewport: { width: 1280, height: 900 } },
    },
  ],
});
