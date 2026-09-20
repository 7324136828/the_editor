import test from 'node:test';
import assert from 'node:assert/strict';
import { importTypeScript } from './plugins-test-utils.mjs';

const { defaultViewPreferences, normalizeViewPreferences, readViewProfiles, setViewVisible, moveView } = await importTypeScript(new URL('../src/views/preferences.ts', import.meta.url));

test('view defaults are specific to each application and include shared file navigation', () => {
  for (const [profile, expected] of [['word', 'word-outline'], ['excel', 'excel-sheets'], ['powerpoint', 'ppt-slides'], ['code', 'code-symbols']]) {
    const defaults = defaultViewPreferences(profile);
    assert.ok(defaults.left.includes('workspace-files'));
    assert.ok(defaults.left.includes(expected));
    assert.deepEqual(defaults.right, ['inspector', 'copilot', 'minimap', 'metadata']);
  }
});

test('hidden and reordered views survive storage without changing another application', () => {
  const profiles = readViewProfiles(null);
  profiles.word = setViewVisible(profiles.word, 'left', 'word-outline', false);
  profiles.word = moveView(profiles.word, 'left', 'word-stats', -1);
  const recovered = readViewProfiles(JSON.stringify({ version: 1, profiles }));
  assert.equal(recovered.word.left.includes('word-outline'), false);
  assert.deepEqual(recovered.word.left, ['open-editors', 'workspace-files', 'word-pages', 'word-stats', 'word-comments']);
  assert.deepEqual(recovered.excel, defaultViewPreferences('excel'));
  assert.deepEqual(recovered.code, defaultViewPreferences('code'));
});

test('empty visible areas are intentional and do not restore hidden defaults', () => {
  const preferences = normalizeViewPreferences({ left: [], right: [], custom: [] }, 'word');
  assert.deepEqual(preferences.left, []);
  assert.deepEqual(preferences.right, []);
});

test('custom notes are retained when hidden and recovered in their correct area', () => {
  const custom = { id: 'custom:checklist', title: 'Editing checklist', kind: 'notes', area: 'right', notes: 'Keep this draft\nReview figures' };
  let preferences = normalizeViewPreferences({ left: ['custom:checklist'], right: ['custom:checklist', 'inspector'], custom: [custom] }, 'word');
  assert.deepEqual(preferences.left, []);
  preferences = setViewVisible(preferences, 'right', custom.id, false);
  const profiles = readViewProfiles(JSON.stringify({ version: 1, profiles: { word: preferences } }));
  assert.deepEqual(profiles.word.right, ['inspector']);
  assert.deepEqual(profiles.word.custom, [custom]);
  assert.deepEqual(setViewVisible(profiles.word, 'right', custom.id, true).right, ['inspector', 'custom:checklist']);
});

test('corrupt or incompatible browser data safely restores defaults', () => {
  for (const value of [null, '{broken', '[]', '{}', '{"version":2,"profiles":{"word":{"left":[]}}}']) {
    assert.deepEqual(readViewProfiles(value).word, defaultViewPreferences('word'));
  }
});

test('validation discards duplicate, wrong-profile, unknown, and unsafe custom identifiers', () => {
  const valid = { id: 'custom:context', title: ' My context ', kind: 'context', area: 'left', notes: '' };
  const preferences = normalizeViewPreferences({
    left: ['word-outline', 'word-outline', 'excel-sheets', 'unknown', 42, 'custom:context'],
    right: ['metadata', 'custom:context'],
    custom: [valid, valid, { ...valid, id: '__proto__' }, { ...valid, id: 'custom:bad', kind: 'script' }, { ...valid, id: 'custom:none', title: '' }],
  }, 'word');
  assert.deepEqual(preferences.left, ['word-outline', 'custom:context']);
  assert.deepEqual(preferences.right, ['metadata']);
  assert.deepEqual(preferences.custom, [{ ...valid, title: 'My context' }]);
});

test('view ordering stops at boundaries and re-enabling an existing view is idempotent', () => {
  const preferences = defaultViewPreferences('word');
  assert.deepEqual(moveView(preferences, 'left', preferences.left[0], -1), preferences);
  assert.deepEqual(moveView(preferences, 'right', 'metadata', 1), preferences);
  assert.deepEqual(moveView(preferences, 'left', 'missing', 1), preferences);
  assert.deepEqual(setViewVisible(preferences, 'right', 'inspector', true), preferences);
});
