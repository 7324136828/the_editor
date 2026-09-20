import { test, expect, type Locator, type Page } from '@playwright/test';

const errors = new WeakMap<Page, string[]>();
const primary = (page: Page) => page.getByRole('separator', { name: 'Resize primary sidebar', exact: true });
const secondary = (page: Page) => page.getByRole('separator', { name: 'Resize secondary sidebar', exact: true });
const bottom = (page: Page) => page.getByRole('separator', { name: 'Resize bottom panel', exact: true });

test.beforeEach(async ({ page }) => {
  errors.set(page, []);
  page.on('pageerror', error => errors.get(page)!.push(error.message));
  await page.goto('/');
  await expect(page.getByTitle('Save to workspace (Ctrl+S)')).toBeEnabled();
});

test.afterEach(async ({ page }) => { expect(errors.get(page)).toEqual([]); });

async function size(locator: Locator, axis: 'width' | 'height') {
  const bounds = await locator.boundingBox();
  expect(bounds).not.toBeNull();
  return bounds![axis];
}

async function drag(page: Page, handle: Locator, dx: number, dy: number) {
  await expect(handle).toBeVisible();
  const bounds = (await handle.boundingBox())!;
  const x = bounds.x + bounds.width / 2;
  const y = bounds.y + bounds.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx, y + dy, { steps: 12 });
  await page.mouse.up();
}

async function expectSize(locator: Locator, axis: 'width' | 'height', expected: number) {
  await expect.poll(async () => Math.abs((await size(locator, axis)) - expected)).toBeLessThanOrEqual(1);
}

async function customizeViews(page: Page) {
  await page.getByRole('button', { name: 'View', exact: true }).click();
  await page.getByRole('menuitem', { name: /Customize views/ }).click();
  return page.getByRole('dialog', { name: 'Customize views', exact: true });
}

test('all three panes resize with the pointer and remember dimensions after closing and reloading', async ({ page }) => {
  await page.keyboard.press('Control+Shift+f');
  const leftPane = page.locator('.vscode-sidebar');
  const rightPane = page.locator('.vscode-right-view');
  const bottomPane = page.locator('.vscode-bottom-panel');
  const initial = { left: await size(leftPane, 'width'), right: await size(rightPane, 'width'), bottom: await size(bottomPane, 'height') };
  await drag(page, primary(page), 85, 0);
  await drag(page, secondary(page), -65, 0);
  await drag(page, bottom(page), 0, -70);
  const changed = { left: await size(leftPane, 'width'), right: await size(rightPane, 'width'), bottom: await size(bottomPane, 'height') };
  expect(changed.left).toBeGreaterThan(initial.left + 70);
  expect(changed.right).toBeGreaterThan(initial.right + 50);
  expect(changed.bottom).toBeGreaterThan(initial.bottom + 55);

  const titlebar = page.locator('.vscode-titlebar');
  for (const [toggle, handle] of [
    ['Toggle primary sidebar', primary(page)],
    ['Toggle inspector', secondary(page)],
    ['Toggle bottom panel', bottom(page)],
  ] as const) {
    await titlebar.getByRole('button', { name: toggle, exact: true }).click();
    await expect(handle).toBeHidden();
    await titlebar.getByRole('button', { name: toggle, exact: true }).click();
    await expect(handle).toBeVisible();
  }
  await expectSize(leftPane, 'width', changed.left);
  await expectSize(rightPane, 'width', changed.right);
  await expectSize(bottomPane, 'height', changed.bottom);
  await page.reload();
  await page.keyboard.press('Control+Shift+f');
  await expect(primary(page)).toBeVisible();
  await expectSize(leftPane, 'width', changed.left);
  await expectSize(rightPane, 'width', changed.right);
  await expectSize(bottomPane, 'height', changed.bottom);
});

test('resize separators support the keyboard and double click restores default sizes', async ({ page }) => {
  await page.keyboard.press('Control+Shift+f');
  for (const [handle, pane, axis, growKey] of [
    [primary(page), page.locator('.vscode-sidebar'), 'width', 'ArrowRight'],
    [secondary(page), page.locator('.vscode-right-view'), 'width', 'ArrowLeft'],
    [bottom(page), page.locator('.vscode-bottom-panel'), 'height', 'ArrowUp'],
  ] as const) {
    const initial = await size(pane, axis);
    await expect(handle).toHaveAttribute('tabindex', '0');
    await handle.focus();
    await handle.press(growKey);
    await handle.press(growKey);
    expect(await size(pane, axis)).toBeGreaterThan(initial);
    await handle.dblclick();
    await expectSize(pane, axis, initial);
  }
});

test('wide pane preferences are clamped to keep the editor accessible in a smaller window', async ({ page }) => {
  await drag(page, primary(page), 380, 0);
  await drag(page, secondary(page), -380, 0);
  await page.setViewportSize({ width: 900, height: 650 });
  const center = (await page.locator('.workspace-center').boundingBox())!;
  const left = (await page.locator('.vscode-sidebar').boundingBox())!;
  const right = (await page.locator('.vscode-right-view').boundingBox())!;
  expect(center.width).toBeGreaterThanOrEqual(280);
  expect(left.x + left.width).toBeLessThanOrEqual(center.x + 1);
  expect(center.x + center.width).toBeLessThanOrEqual(right.x + 1);
  expect(right.x + right.width).toBeLessThanOrEqual(901);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(900);
  await expect(page.getByTitle('Save to workspace (Ctrl+S)')).toBeInViewport();
  await page.screenshot({ path: '.verification/layout-compact.png', fullPage: true });
});

test('resizing the bottom panel in a small window preserves the preferred sidebar widths', async ({ page }) => {
  await drag(page, primary(page), 180, 0);
  await drag(page, secondary(page), -140, 0);
  const leftPane = page.locator('.vscode-sidebar');
  const rightPane = page.locator('.vscode-right-view');
  const preferred = { left: await size(leftPane, 'width'), right: await size(rightPane, 'width') };
  await page.setViewportSize({ width: 900, height: 650 });
  await page.keyboard.press('Control+Shift+f');
  await drag(page, bottom(page), 0, -45);
  await page.setViewportSize({ width: 1600, height: 1000 });
  await expectSize(leftPane, 'width', preferred.left);
  await expectSize(rightPane, 'width', preferred.right);
  await page.reload();
  await expectSize(leftPane, 'width', preferred.left);
  await expectSize(rightPane, 'width', preferred.right);
});

test('compact screens show one sidebar at a time and keep the layout toggles reachable', async ({ page }) => {
  await page.setViewportSize({ width: 680, height: 700 });
  const titlebar = page.locator('.vscode-titlebar');
  const togglePrimary = titlebar.getByRole('button', { name: 'Toggle primary sidebar', exact: true });
  const toggleSecondary = titlebar.getByRole('button', { name: 'Toggle inspector', exact: true });
  await expect(togglePrimary).toBeInViewport();
  await expect(toggleSecondary).toBeInViewport();
  await expect.poll(async () => Number(await primary(page).isVisible()) + Number(await secondary(page).isVisible())).toBe(1);
  if (!(await secondary(page).isVisible())) await toggleSecondary.click();
  await expect(secondary(page)).toBeVisible();
  await expect(primary(page)).toBeHidden();
  await togglePrimary.click();
  await expect(primary(page)).toBeVisible();
  await expect(secondary(page)).toBeHidden();
  await togglePrimary.click();
  await expect(primary(page)).toBeHidden();
  await expect(secondary(page)).toBeHidden();
  await expect(page.getByTitle('Save to workspace (Ctrl+S)')).toBeInViewport();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(680);
  await page.screenshot({ path: '.verification/layout-mobile.png', fullPage: true });
});

test('document chrome omits developer panels and the activity search icon while find remains accessible', async ({ page }) => {
  const activityBar = page.locator('.vscode-activitybar');
  await expect(activityBar.locator('svg.lucide-search')).toHaveCount(0);
  await expect(activityBar.getByRole('button', { name: /debug|search/i })).toHaveCount(0);
  await page.keyboard.press('Control+Shift+f');
  await expect(page.getByPlaceholder('Search across all open documents...')).toBeVisible();
  await expect(page.locator('.vscode-bottom-panel')).not.toContainText(/PROBLEMS|OUTPUT|TERMINAL|DEBUG/);
  await page.keyboard.press('Control+Shift+p');
  const commandSearch = page.getByPlaceholder('Type a command or search files...');
  for (const command of ['problems', 'output', 'terminal', 'debug']) {
    await commandSearch.fill(command);
    await expect(page.getByText('No matching commands found.', { exact: true })).toBeVisible();
  }
  await page.keyboard.press('Escape');
  await page.locator('.vscode-titlebar').getByRole('button', { name: 'Toggle bottom panel', exact: true }).click();
  await page.getByRole('button', { name: 'Edit', exact: true }).click();
  await page.getByRole('menuitem', { name: /Find \/ Replace in files/ }).click();
  await expect(page.getByPlaceholder('Search across all open documents...')).toBeVisible();
});

test('views can be hidden, reordered and added with persistent notes and independent application arrangements', async ({ page }) => {
  const originalTitle = await page.locator('.vscode-tab.active > span').textContent();
  let dialog = await customizeViews(page);
  await dialog.getByRole('checkbox', { name: /^Document outline/ }).uncheck();
  await dialog.getByRole('button', { name: 'Move Workspace files up', exact: true }).click();
  await dialog.getByRole('checkbox', { name: /^Inspector/ }).uncheck();
  await dialog.getByLabel('View name', { exact: true }).fill('Review notes');
  await dialog.getByRole('combobox', { name: 'Content', exact: true }).selectOption('notes');
  await dialog.getByRole('combobox', { name: 'Location', exact: true }).selectOption('right');
  await dialog.getByRole('button', { name: 'Add view', exact: true }).click();
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(page.locator('.vscode-sidebar .vscode-accordion-header').first()).toContainText(/WORKSPACE FILES/i);
  await expect(page.locator('.vscode-sidebar').getByRole('button', { name: /DOCUMENT OUTLINE/i })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: 'Inspector', exact: true })).toHaveCount(0);
  await page.getByRole('tab', { name: 'Review notes', exact: true }).click();
  await page.getByLabel('Review notes notes', { exact: true }).fill('Check headings before publishing.');
  await drag(page, primary(page), 60, 0);
  await drag(page, secondary(page), -115, 0);
  await page.keyboard.press('Control+Shift+f');
  await drag(page, bottom(page), 0, -40);
  await page.screenshot({ path: '.verification/layout-custom-views.png', fullPage: true });

  dialog = await customizeViews(page);
  await dialog.getByRole('checkbox', { name: /^Review notes/ }).uncheck();
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Review notes', exact: true })).toHaveCount(0);
  await page.reload();
  dialog = await customizeViews(page);
  await expect(dialog.getByRole('checkbox', { name: /^Document outline/ })).not.toBeChecked();
  await expect(dialog.getByRole('checkbox', { name: /^Review notes/ })).not.toBeChecked();
  await dialog.getByRole('checkbox', { name: /^Review notes/ }).check();
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await page.getByRole('tab', { name: 'Review notes', exact: true }).click();
  await expect(page.getByLabel('Review notes notes', { exact: true })).toHaveValue('Check headings before publishing.');

  await page.getByTitle('New file (Ctrl+N)').click();
  const createDialog = page.getByRole('dialog', { name: 'Create a file', exact: true });
  await createDialog.getByLabel('File type').selectOption('excel');
  await createDialog.getByLabel('Filename').fill(`View-profile-${Date.now()}.xlsx`);
  await createDialog.getByRole('button', { name: 'Create file', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Workbook explorer', exact: true })).toBeVisible();
  await expect(page.getByRole('tab', { name: 'Review notes', exact: true })).toHaveCount(0);
  await expect(page.getByRole('tab', { name: 'Inspector', exact: true })).toBeVisible();
  await page.locator('.vscode-tab').filter({ hasText: originalTitle! }).click();
  await expect(page.getByRole('tab', { name: 'Review notes', exact: true })).toBeVisible();
  dialog = await customizeViews(page);
  await dialog.getByRole('button', { name: 'Restore default views', exact: true }).click();
  await expect(dialog.getByRole('checkbox', { name: /^Document outline/ })).toBeChecked();
  await expect(dialog.getByRole('checkbox', { name: /^Review notes/ })).not.toBeChecked();
  await dialog.getByRole('checkbox', { name: /^Review notes/ }).check();
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await page.getByRole('tab', { name: 'Review notes', exact: true }).click();
  await expect(page.getByLabel('Review notes notes', { exact: true })).toHaveValue('Check headings before publishing.');

  dialog = await customizeViews(page);
  await dialog.getByRole('button', { name: 'Delete Review notes', exact: true }).click();
  await dialog.getByRole('button', { name: 'Cancel deletion', exact: true }).click();
  await expect(dialog.getByRole('checkbox', { name: /^Review notes/ })).toBeChecked();
  await dialog.getByRole('button', { name: 'Delete Review notes', exact: true }).click();
  await dialog.getByRole('button', { name: 'Delete view and notes', exact: true }).click();
  await expect(dialog.getByRole('checkbox', { name: /^Review notes/ })).toHaveCount(0);
  await dialog.getByRole('button', { name: 'Done', exact: true }).click();
  await page.reload();
  await expect(page.getByRole('tab', { name: 'Review notes', exact: true })).toHaveCount(0);
});
