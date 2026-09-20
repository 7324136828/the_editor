import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  testMatch: ['browser.spec.ts', 'layout.spec.ts'],
  fullyParallel: false,
  workers: 1,
  timeout: 45000,
  expect: { timeout: 8000 },
  reporter: [['list'], ['html', { open: 'never', outputFolder: '.verification/playwright-report' }]],
  outputDir: '.verification/playwright-results',
  use: {
    baseURL: 'http://127.0.0.1:5185',
    channel: 'msedge',
    headless: true,
    viewport: { width: 1600, height: 1000 },
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command: 'node tests/browser-server.mjs',
    url: 'http://127.0.0.1:5185/api/health',
    reuseExistingServer: false,
    timeout: 30000,
  },
});
