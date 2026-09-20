import { test, expect, type Page } from '@playwright/test';

const run = Date.now().toString(36);
const filename = (stem: string, extension: string) => `Browser-${stem}-${run}.${extension}`;
const errors = new WeakMap<Page, string[]>();

test.beforeEach(async ({ page }) => {
  errors.set(page, []);
  page.on('pageerror', error => errors.get(page)!.push(error.message));
  await page.goto('/');
  await expect(page.getByTitle('Save to workspace (Ctrl+S)')).toBeEnabled();
});
test.afterEach(async ({ page }) => { expect(errors.get(page)).toEqual([]); });

async function createFile(page: Page, type: 'word' | 'excel' | 'powerpoint' | 'code', name: string) {
  await page.getByTitle('New file (Ctrl+N)').click();
  const dialog = page.getByRole('dialog', { name: 'Create a file' });
  await dialog.getByLabel('File type').selectOption(type);
  await dialog.getByLabel('Filename').fill(name);
  await dialog.getByRole('button', { name: 'Create file', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator('.vscode-tab.active')).toContainText(name);
}

async function save(page: Page) {
  const response = page.waitForResponse(response => response.request().method() === 'PUT' && response.url().includes('/api/documents/'));
  await page.keyboard.press('Control+s');
  const saved = await response;
  expect(saved.ok()).toBe(true);
  await expect(page.locator('.save-state')).toContainText('Saved');
  return saved.json();
}

async function workbookCell(page: Page, coord: string, value: string) {
  await page.locator(`[data-cell="${coord}"]`).dblclick();
  await page.getByLabel(`Edit ${coord}`, { exact: true }).fill(value);
  await page.getByLabel(`Edit ${coord}`, { exact: true }).press('Enter');
}

test('Word edits, formatting, table and comments save to disk and survive a clean reload', async ({ page }) => {
  await page.screenshot({ path: '.verification/browser-desktop.png', fullPage: true });
  const name = filename('document', 'docx');
  await createFile(page, 'word', name);
  await expect(page.getByRole('button', { name: 'Document outline', exact: true })).toHaveAttribute('data-document-type', 'word');
  await page.getByLabel('Heading 1', { exact: true }).fill('A usable document editor');
  const paragraph = page.getByLabel('Paragraph 2', { exact: true });
  await paragraph.fill('Typing stays in order');
  await paragraph.press('End');
  await paragraph.pressSequentially(' and keeps the caret.');
  await expect(paragraph).toHaveValue('Typing stays in order and keeps the caret.');
  await paragraph.press('Control+b');
  await expect(page.getByRole('button', { name: 'Bold paragraph', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByRole('button', { name: 'Document outline', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Add table', exact: true }).click();
  await page.getByLabel('Table 3, row 2, column 2', { exact: true }).fill('Editable table content');
  await page.getByRole('button', { name: 'Comments', exact: true }).click();
  await page.getByLabel('New comment', { exact: true }).fill('Review this table.');
  await page.getByRole('button', { name: 'Add comment', exact: true }).click();
  const record = await save(page);
  expect(record.document.data.paragraphs[1].text).toBe('Typing stays in order and keeps the caret.');
  expect(record.document.data.paragraphs[1].runs[0].bold).toBe(true);
  expect(record.document.data.paragraphs[2].tableData[1][1]).toBe('Editable table content');
  expect(record.document.data.comments[0].comment).toBe('Review this table.');
  await page.evaluate(() => localStorage.removeItem('office-studio.workspace.v2'));
  await page.reload();
  await page.getByTitle(`Open ${name}`, { exact: true }).click();
  await expect(page.getByLabel('Heading 1', { exact: true })).toHaveValue('A usable document editor');
  await expect(page.getByLabel('Table 3, row 2, column 2', { exact: true })).toHaveValue('Editable table content');
  await expect(page.locator('.save-state')).toContainText('Saved');
});

test('SmartArt plugin inserts an editable Word diagram that saves and reloads', async ({ page }) => {
  const name = filename('smartart', 'docx');
  await createFile(page, 'word', name);
  await page.getByRole('button', { name: 'Extensions & MCP (Ctrl+Shift+X)', exact: true }).click();
  const plugin = page.locator('.studio-extension-card').filter({ hasText: 'SmartArt for Word' });
  await expect(plugin).toContainText('Insert editable process, cycle, hierarchy, and pyramid diagrams');
  await plugin.getByRole('button', { name: 'Insert process SmartArt', exact: true }).click();
  await expect(page.getByLabel('SmartArt 3 title', { exact: true })).toHaveValue('Process');
  await page.getByLabel('SmartArt 3 title', { exact: true }).fill('Release workflow');
  await page.getByLabel('SmartArt layout', { exact: true }).selectOption('cycle');
  await page.getByLabel('SmartArt accent color', { exact: true }).fill('#2563eb');
  await page.getByRole('button', { name: 'Add item', exact: true }).click();
  await page.getByLabel('SmartArt 3, item 5', { exact: true }).fill('Measure outcomes');
  await page.getByRole('button', { name: 'Move SmartArt item 5 up', exact: true }).click();
  await expect(page.getByLabel('SmartArt 3, item 4', { exact: true })).toHaveValue('Measure outcomes');
  await page.screenshot({ path: '.verification/browser-smartart.png', fullPage: true });
  const record = await save(page);
  const diagram = record.document.data.paragraphs[2];
  expect(diagram.type).toBe('smartart');
  expect(diagram.smartArt.title).toBe('Release workflow');
  expect(diagram.smartArt.layout).toBe('cycle');
  expect(diagram.smartArt.accentColor).toBe('#2563eb');
  expect(diagram.smartArt.items[3].text).toBe('Measure outcomes');
  await page.evaluate(() => localStorage.removeItem('office-studio.workspace.v2'));
  await page.reload();
  await page.getByTitle(`Open ${name}`, { exact: true }).click();
  await expect(page.getByLabel('SmartArt 3 title', { exact: true })).toHaveValue('Release workflow');
  await expect(page.getByLabel('SmartArt 3, item 4', { exact: true })).toHaveValue('Measure outcomes');
});

test('multiple documents of the same type retain separate identities and dirty close can be cancelled', async ({ page }) => {
  const firstName = filename('first', 'docx');
  const secondName = filename('second', 'docx');
  await createFile(page, 'word', firstName);
  await page.getByLabel('Paragraph 2', { exact: true }).fill('First document body');
  const first = await save(page);
  await createFile(page, 'word', secondName);
  await page.getByLabel('Paragraph 2', { exact: true }).fill('Second document body');
  await page.keyboard.press('Control+w');
  const dialog = page.getByRole('dialog', { name: 'Save changes before closing?' });
  await expect(dialog).toContainText(secondName);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(page.getByLabel('Paragraph 2', { exact: true })).toHaveValue('Second document body');
  const second = await save(page);
  expect(first.id).not.toBe(second.id);
  await page.locator('.vscode-tab').filter({ hasText: firstName }).click();
  await expect(page.getByLabel('Paragraph 2', { exact: true })).toHaveValue('First document body');
  await page.locator('.vscode-tab').filter({ hasText: secondName }).click();
  await expect(page.getByLabel('Paragraph 2', { exact: true })).toHaveValue('Second document body');
});

test('spreadsheet edits recalculate dependencies and save the latest formula before blur', async ({ page }) => {
  await createFile(page, 'excel', filename('workbook', 'xlsx'));
  await expect(page.getByRole('button', { name: 'Workbook explorer', exact: true })).toHaveAttribute('data-document-type', 'excel');
  await workbookCell(page, 'A1', '10');
  await workbookCell(page, 'A2', '20');
  await workbookCell(page, 'B1', '=SUM(A1:A2)');
  await workbookCell(page, 'B2', '=B1*2');
  await expect(page.locator('[data-cell="B2"]')).toHaveText('60');
  await workbookCell(page, 'A1', '40');
  await expect(page.locator('[data-cell="B1"]')).toHaveText('60');
  await expect(page.locator('[data-cell="B2"]')).toHaveText('120');
  await page.locator('[data-cell="C1"]').click();
  await page.getByLabel('Formula bar', { exact: true }).fill('=AVERAGE(A1:A2)');
  const record = await save(page);
  expect(record.document.data.sheets[0].cells.C1.value).toBe(30);
  expect(record.document.data.sheets[0].cells.B2.value).toBe(120);
  await page.getByRole('button', { name: 'Settings', exact: true }).click();
  await page.getByRole('combobox', { name: 'Color theme', exact: true }).selectOption('light');
  await expect(page.locator('[data-cell="B2"]')).toHaveCSS('color', 'rgb(51, 51, 51)');
  await page.screenshot({ path: '.verification/browser-excel-light.png', fullPage: true });
  await page.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await expect(page.locator('.sheet-tab.active')).toContainText('Sheet 2');
  await page.getByRole('button', { name: '+ Row', exact: true }).click();
  await page.getByRole('button', { name: '+ Column', exact: true }).click();
  await expect(page.locator('.sheet-tabs')).toContainText('26 rows × 9 columns');
});

test('presentation text, notes and slides edit, persist and present the active slide', async ({ page }) => {
  await createFile(page, 'powerpoint', filename('presentation', 'pptx'));
  await expect(page.getByRole('button', { name: 'Slide navigator', exact: true })).toHaveAttribute('data-document-type', 'powerpoint');
  await page.getByLabel('Slide object: Your next great idea', { exact: true }).click();
  await page.getByLabel('Object text', { exact: true }).fill('Editable presentation title');
  await page.getByLabel('Speaker notes', { exact: true }).fill('Remember the demo talking points.');
  await page.locator('.office-editor').getByRole('button', { name: 'Duplicate slide', exact: true }).click();
  await expect(page.locator('.ppt-thumb')).toHaveCount(2);
  await page.getByLabel('Slide title', { exact: true }).fill('Second slide');
  await page.getByLabel('Slide object: Editable presentation title', { exact: true }).click();
  await page.getByLabel('Object text', { exact: true }).fill('Second slide edited independently');
  const record = await save(page);
  expect(record.document.data.slides[0].objects[0].text).toBe('Editable presentation title');
  expect(record.document.data.slides[1].objects[0].text).toBe('Second slide edited independently');
  expect(record.document.data.slides[1].notes).toBe('Remember the demo talking points.');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  const slideshow = page.getByRole('dialog', { name: 'Slideshow', exact: true });
  await expect(slideshow).toContainText('Second slide edited independently');
  await expect(slideshow).toContainText('2 / 2');
  await page.screenshot({ path: '.verification/browser-ppt.png', fullPage: true });
  await page.keyboard.press('ArrowLeft');
  await expect(slideshow).toContainText('Editable presentation title');
  await page.keyboard.press('Escape');
  await expect(slideshow).toBeHidden();
});

test('layout plugins enable and disable, change real page geometry, and connect to MCP', async ({ page }) => {
  await createFile(page, 'word', filename('plugins', 'docx'));
  await page.getByRole('button', { name: 'Extensions & MCP (Ctrl+Shift+X)', exact: true }).click();
  await page.getByRole('button', { name: 'Disable Document Layout', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Landscape report', exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: 'Enable Document Layout', exact: true }).click();
  const before = await page.locator('.word-page').boundingBox();
  await page.getByRole('button', { name: 'Landscape report', exact: true }).click();
  await expect(page.locator('.word-page')).toHaveCSS('width', '1056px');
  expect((await page.locator('.word-page').boundingBox())!.width).toBeGreaterThan(before!.width);
  await page.getByRole('button', { name: 'Compact report', exact: true }).click();
  await expect(page.locator('.word-page')).toHaveCSS('padding-left', '36px');
  await page.getByRole('button', { name: 'Inspect document context', exact: true }).click();
  await expect(page.getByText('Current context', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Test connection', exact: true }).click();
  await expect(page.getByText('Connected', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Settings', exact: true }).click();
  await page.getByRole('combobox', { name: 'Color theme', exact: true }).selectOption('light');
  await expect(page.locator('body')).toHaveClass(/theme-light/);
  await page.screenshot({ path: '.verification/browser-light.png', fullPage: true });
});

test('search replacement honors case and whole-word options and invalid regex stays contained', async ({ page }) => {
  await createFile(page, 'code', filename('search', 'md'));
  const editor = page.getByLabel('markdown editor', { exact: true });
  await editor.fill('QaNeedle qaneedle QaNeedles\nOther content');
  await page.keyboard.press('Control+Shift+f');
  await page.getByPlaceholder('Search across all open documents...').fill('QaNeedle');
  await page.getByTitle('Match Case (Alt+C)').click();
  await page.getByTitle('Match Whole Word (Alt+W)').click();
  await expect(page.getByText('Found 1 results in 1 files', { exact: true })).toBeVisible();
  await page.getByPlaceholder('Replace with...').fill('Replacement');
  await page.getByTitle('Replace All across loaded documents').click();
  await page.getByRole('dialog', { name: 'Replace across workspace?' }).getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(editor).toHaveValue('Replacement qaneedle QaNeedles\nOther content');
  await page.getByTitle('Use Regular Expression (Alt+R)').click();
  await page.getByPlaceholder('Search across all open documents...').fill('[');
  await expect(page.getByRole('alert')).toContainText(/regular expression|regex|invalid/i);
  await expect(page.getByTitle('Replace All across loaded documents')).toBeDisabled();
});

test('failed saves retain edits through reload and can be retried', async ({ page }) => {
  await createFile(page, 'code', filename('offline-recovery', 'md'));
  const editor = page.getByLabel('markdown editor', { exact: true });
  await editor.fill('These edits must survive an unavailable storage server.');
  await page.route('**/api/documents/*', route => route.request().method() === 'PUT'
    ? route.fulfill({ status: 503, contentType: 'application/json', body: JSON.stringify({ error: 'Storage temporarily unavailable for browser test.' }) }) : route.continue());
  await page.keyboard.press('Control+s');
  await expect(page.locator('.save-state')).toContainText('Save failed');
  await expect(editor).toHaveValue('These edits must survive an unavailable storage server.');
  await page.reload();
  await expect(editor).toHaveValue('These edits must survive an unavailable storage server.');
  await page.unroute('**/api/documents/*');
  const record = await save(page);
  expect(record.document.data.content).toBe('These edits must survive an unavailable storage server.');
});

test('stale revision conflict retains local work and Save a copy resolves it safely', async ({ page }) => {
  await createFile(page, 'code', filename('conflict-original', 'md'));
  const editor = page.getByLabel('markdown editor', { exact: true });
  await editor.fill('Original saved content.');
  const original = await save(page);
  const remote = structuredClone(original.document);
  remote.data.content = 'A different session saved this content.';
  const remoteSave = await page.request.put(`/api/documents/${original.id}`, { data: { name: original.name, revision: original.revision, document: remote } });
  expect(remoteSave.ok()).toBe(true);
  await editor.fill('Local conflicting edits must remain available.');
  const conflictResponse = page.waitForResponse(response => response.request().method() === 'PUT' && response.url().endsWith(`/api/documents/${original.id}`));
  await page.keyboard.press('Control+s');
  expect((await conflictResponse).status()).toBe(409);
  await expect(page.locator('.save-state')).toContainText('Save failed');
  await expect(editor).toHaveValue('Local conflicting edits must remain available.');
  await page.getByRole('button', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: /Save a copy/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Save a copy', exact: true });
  await dialog.getByLabel('Filename').fill(filename('conflict-copy', 'md'));
  const copyResponse = page.waitForResponse(response => response.request().method() === 'PUT' && response.url().includes('/api/documents/'));
  await dialog.getByRole('button', { name: 'Save copy', exact: true }).click();
  const copied = await (await copyResponse).json();
  expect(copied.id).not.toBe(original.id);
  expect(copied.document.data.content).toBe('Local conflicting edits must remain available.');
  const remoteRecord = await (await page.request.get(`/api/documents/${original.id}`)).json();
  expect(remoteRecord.document.data.content).toBe('A different session saved this content.');
});

test('autosave persists edits and reload keeps the current document', async ({ page }) => {
  const name = filename('autosave', 'docx');
  await createFile(page, 'word', name);
  await page.getByRole('button', { name: 'Settings', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Auto Save to local server', exact: true }).check();
  const paragraph = page.getByLabel('Paragraph 2', { exact: true });
  const saved = page.waitForResponse(response => response.request().method() === 'PUT' && response.url().includes('/api/documents/') && response.request().postDataJSON()?.name === name);
  await paragraph.fill('Autosave persists the final typed sentence.');
  const record = await (await saved).json();
  expect(record.document.data.paragraphs[1].text).toBe('Autosave persists the final typed sentence.');
  await expect(page.locator('.save-state')).toContainText('Saved');
  await page.reload();
  await expect(page.locator('.vscode-tab.active')).toContainText(name);
  await expect(paragraph).toHaveValue('Autosave persists the final typed sentence.');
});

test('undo, redo, rename and discard preserve document content and identity', async ({ page }) => {
  const originalName = filename('history', 'docx');
  const renamed = filename('history-renamed', 'docx');
  await createFile(page, 'word', originalName);
  const paragraph = page.getByLabel('Paragraph 2', { exact: true });
  await paragraph.fill('Saved baseline.');
  const original = await save(page);
  await paragraph.fill('Changed after saving.');
  await page.getByTitle('Undo', { exact: true }).click();
  await expect(paragraph).toHaveValue('Saved baseline.');
  await page.getByTitle('Redo', { exact: true }).click();
  await expect(paragraph).toHaveValue('Changed after saving.');
  await page.getByRole('button', { name: 'File', exact: true }).click();
  await page.getByRole('menuitem', { name: /^Rename/ }).click();
  const dialog = page.getByRole('dialog', { name: 'Rename file', exact: true });
  await dialog.getByLabel('Filename').fill(renamed);
  await dialog.getByRole('button', { name: 'Rename', exact: true }).click();
  const record = await save(page);
  expect(record.id).toBe(original.id);
  expect(record.name).toBe(renamed);
  await paragraph.fill('These changes will be discarded.');
  await page.keyboard.press('Control+w');
  await page.getByRole('dialog', { name: 'Save changes before closing?' }).getByRole('button', { name: 'Discard changes', exact: true }).click();
  await page.getByTitle(`Open ${renamed}`, { exact: true }).click();
  await expect(paragraph).toHaveValue('Changed after saving.');
  await expect(page.locator('.save-state')).toContainText('Saved');
});
