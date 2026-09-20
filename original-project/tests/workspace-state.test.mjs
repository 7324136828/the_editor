import test from 'node:test';
import assert from 'node:assert/strict';
import { importTypeScript } from './plugins-test-utils.mjs';

const { recoverWorkspace, isOfficeDocument, isStoredDocument, mergeWorkspaceRecords } = await importTypeScript(new URL('../src/hooks/workspaceState.ts', import.meta.url));
const samples = await importTypeScript(new URL('../src/data/sampleDocuments.ts', import.meta.url));
const document = (content = 'Saved text', id = 'code-one') => ({ type: 'code', data: { id, title: 'Notes.md', language: 'markdown', content, symbols: [] } });
const entry = (content = 'Saved text', id = 'code-one') => {
  const doc = document(content, id);
  return { id, name: doc.data.title, document: doc, revision: 1, saved: JSON.stringify(doc) };
};
const state = item => ({ entries: { [item.id]: item }, openIds: [item.id], activeId: item.id });
const record = (content = 'Server text', revision = 2, id = 'code-one') => ({ id, name: 'Notes.md', document: document(content, id), revision, updatedAt: '2026-09-19T01:00:00Z' });

test('recovery validates all four supported sample document models', () => {
  for (const [type, value] of [['word', samples.sampleWordDocument], ['excel', samples.sampleExcelDocument], ['powerpoint', samples.samplePptDocument], ['code', samples.sampleCodeDocument]]) {
    assert.equal(isOfficeDocument({ type, data: value }), true, `${type} sample must remain recoverable`);
  }
});

test('recovery preserves valid drafts, clears interrupted saving, and repairs tab references', () => {
  const first = entry(); first.document = document('Unsaved work'); first.saving = true; first.name = 'Stale name.md';
  const recovered = recoverWorkspace({ entries: { 'code-one': first, broken: { id: 'broken', document: { type: 'word', data: {} } } }, openIds: ['missing', 'code-one', 'code-one'], activeId: 'missing' });
  assert.deepEqual(Object.keys(recovered.entries), ['code-one']);
  assert.equal(recovered.entries['code-one'].document.data.content, 'Unsaved work');
  assert.equal(recovered.entries['code-one'].saving, false);
  assert.equal(recovered.entries['code-one'].name, 'Notes.md');
  assert.deepEqual(recovered.openIds, ['code-one']);
  assert.equal(recovered.activeId, 'code-one');
});

test('mismatched document identities and incomplete models cannot become editor entries', () => {
  const wrong = entry(); wrong.document.data.id = 'another-document';
  assert.equal(recoverWorkspace(state(wrong)), null);
  assert.equal(isOfficeDocument({ type: 'code', data: { ...document().data, content: {} } }), false);
  assert.equal(isOfficeDocument({ type: 'word', data: { id: 'word-one', title: 'broken.docx' } }), false);
  assert.equal(isStoredDocument({ ...record(), document: document('text', 'different') }), false);
  assert.equal(isStoredDocument({ ...record(), revision: 0 }), false);
  const tooLarge = structuredClone(samples.sampleExcelDocument);
  tooLarge.sheets[0].rowCount = 1000000;
  assert.equal(isOfficeDocument({ type: 'excel', data: tooLarge }), false);
});

test('damaged saved baseline cannot overwrite the retained current draft during discard', () => {
  const item = entry('Important draft'); item.saved = '{corrupt';
  const recovered = recoverWorkspace(state(item));
  assert.equal(recovered.entries[item.id].saved, null);
  assert.equal(recovered.entries[item.id].document.data.content, 'Important draft');
  const other = entry('Important draft'); other.saved = JSON.stringify(document('Wrong file', 'another-id'));
  assert.equal(recoverWorkspace(state(other)).entries[other.id].saved, null);
});

test('an intentionally empty recovered workspace stays empty', () => {
  const result = recoverWorkspace({ entries: {}, openIds: [], activeId: 'gone' });
  assert.deepEqual(Object.keys(result.entries), []);
  assert.deepEqual(result.openIds, []);
  assert.equal(result.activeId, '');
});

test('refresh hydrates a pristine seeded file but preserves a recovered dirty draft', () => {
  const seed = entry(); seed.saved = null; seed.revision = null; seed.seeded = true;
  const initial = state(seed);
  assert.equal(mergeWorkspaceRecords(initial, initial, [record()]).entries[seed.id].document.data.content, 'Server text');
  const dirty = entry(); dirty.document = document('Unsaved work');
  const recovered = state(dirty);
  assert.equal(mergeWorkspaceRecords(recovered, recovered, [record()]).entries[dirty.id].document.data.content, 'Unsaved work');
});

test('refresh cannot replace edits or completed saves made since the list request began', () => {
  const initial = state(entry());
  const edited = state({ ...initial.entries['code-one'], document: document('New typing') });
  assert.equal(mergeWorkspaceRecords(edited, initial, [record()]).entries['code-one'], edited.entries['code-one']);
  const afterSave = state({ ...entry('Newly saved text'), revision: 3 });
  assert.equal(mergeWorkspaceRecords(afterSave, initial, [record('Older list result', 2)]).entries['code-one'].revision, 3);
  assert.equal(mergeWorkspaceRecords(afterSave, initial, [record('Older list result', 2)]).entries['code-one'].document.data.content, 'Newly saved text');
});

test('refresh does not touch in-flight saves or roll back a newer revision', () => {
  const current = state(entry());
  assert.equal(mergeWorkspaceRecords(current, current, [record()], new Set(['code-one'])).entries['code-one'], current.entries['code-one']);
  current.entries['code-one'].revision = 5;
  assert.equal(mergeWorkspaceRecords(current, current, [record('Older server snapshot', 4)]).entries['code-one'].revision, 5);
});

test('refresh does not resurrect an entry discarded during the request', () => {
  const initial = state(entry());
  const emptied = { entries: {}, openIds: [], activeId: '' };
  assert.equal(mergeWorkspaceRecords(emptied, initial, [record()]).entries['code-one'], undefined);
  const fresh = mergeWorkspaceRecords(emptied, emptied, [record()]);
  assert.equal(fresh.entries['code-one'].document.data.content, 'Server text');
});
