import { expect, test } from '@playwright/test';

const backendUrl = 'http://127.0.0.1:8001';

test('backend health endpoint responds with ok', async ({ request }) => {
  const healthResponse = await request.get(`${backendUrl}/api/health`);
  expect(healthResponse.ok()).toBeTruthy();
  const data = await healthResponse.json();
  expect(data.status).toBe('ok');
  expect(data.storage).toBe('local-disk');
});

test('loads Code Office Studio with full workbench and navigation', async ({ page }) => {
  page.on('console', msg => console.log('[BROWSER CONSOLE]', msg.text()));
  page.on('pageerror', err => console.log('[BROWSER PAGE ERROR]', err.message));
  await page.goto('/');

  // Title bar and branding
  await expect(page.locator('.vscode-app-logo')).toContainText('Code Office');

  // View navigation switcher
  await expect(page.getByRole('button', { name: 'Studio', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'AI Copilot', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'History', exact: true })).toBeVisible();

  // Switch to AI Copilot
  await page.getByRole('button', { name: 'AI Copilot', exact: true }).click();
  await expect(page.getByText('Office Copilot').first()).toBeVisible();

  // Send a chat message
  const chatInput = page.getByPlaceholder(/Ask/i);
  await expect(chatInput).toBeVisible();
  await chatInput.fill('Summarize document');
  await chatInput.press('Enter');

  // Verify response appears
  await expect(page.getByText('You', { exact: true })).toBeVisible();
  await expect(page.getByText('Copilot').nth(1)).toBeVisible();

  // Switch to History view
  await page.getByRole('button', { name: 'History', exact: true }).click();
  await expect(page.getByText(/Historical Conversions/i)).toBeVisible();

  // Switch back to Studio
  await page.getByRole('button', { name: 'Studio', exact: true }).click();
  await expect(page.locator('.vscode-main-workspace')).toBeVisible();
});

