import test from 'node:test';
import assert from 'node:assert/strict';
import { importTypeScript } from './plugins-test-utils.mjs';
import { editorialLayoutPlugin } from '../examples/plugins/editorial-layout.js';

const { createPluginHost } = await importTypeScript(new URL('../src/plugins/host.ts', import.meta.url));
const { builtinPlugins } = await importTypeScript(new URL('../src/plugins/builtins.ts', import.meta.url));
const { collectDocumentContext, getDocumentDiagnostics } = await importTypeScript(new URL('../src/plugins/context.ts', import.meta.url));

const word = () => ({ type: 'word', data: { id: 'word-1', title: 'Notes.docx', paragraphs: [{ id: 'p1', type: 'heading-1', text: 'Real heading' }, { id: 'p2', type: 'body', text: 'Current editor content.' }], headings: [], sections: [{ id: 'section-1', title: 'Document', pageNumber: 1, orientation: 'portrait' }], comments: [], typography: { fontFamily: 'Arial', fontSize: 14, lineHeight: 1.5, marginSize: 'normal', textColor: '#000000' }, stats: { words: 999 } } });
const code = () => ({ type: 'code', data: { id: 'code-1', title: 'main.json', language: 'json', content: '{"hello":true}', symbols: [] } });

test('layout plugin changes the model without mutating the source document', () => {
  const document = word();
  const host = createPluginHost({ plugins: builtinPlugins, storage: null });
  const output = host.execute('studio.layout.landscape', document);
  assert.equal(output.document.data.sections[0].orientation, 'landscape');
  assert.equal(document.data.sections[0].orientation, 'portrait');
  assert.equal(host.execute('studio.layout.compact', document).document.data.typography.marginSize, 'narrow');
  assert.equal(host.commands(code()).filter(command => command.pluginId === 'studio.layout').length, 0);
  assert.throws(() => host.execute('studio.layout.business', code()), /does not support/);
});

test('SmartArt plugin inserts structured editable diagrams only in Word documents', () => {
  const document = word();
  const host = createPluginHost({ plugins: builtinPlugins, storage: null });
  const commands = host.commands(document).filter(command => command.pluginId === 'studio.smartart');
  assert.deepEqual(commands.map(command => command.id), [
    'studio.smartart.insert-process',
    'studio.smartart.insert-cycle',
    'studio.smartart.insert-hierarchy',
    'studio.smartart.insert-pyramid',
  ]);
  assert.equal(host.commands(code()).some(command => command.pluginId === 'studio.smartart'), false);
  const output = host.execute('studio.smartart.insert-process', document);
  const inserted = output.document.data.paragraphs.at(-1);
  assert.equal(inserted.type, 'smartart');
  assert.equal(inserted.smartArt.layout, 'process');
  assert.deepEqual(inserted.smartArt.items.map(item => item.text), ['Plan', 'Build', 'Review', 'Deliver']);
  assert.equal(document.data.paragraphs.length, 2, 'plugin execution must not mutate its input');
  assert.throws(() => host.execute('studio.smartart.insert-process', code()), /does not support/);
});

test('disable persists, prevents execution, and emits change notifications', () => {
  const storage = { saved: '', getItem() { return this.saved; }, setItem(key, value) { this.saved = value; } };
  const host = createPluginHost({ plugins: builtinPlugins, storage });
  let changes = 0;
  const unsubscribe = host.subscribe(() => changes++);
  host.setEnabled('studio.layout', false);
  assert.equal(changes, 1);
  unsubscribe();
  assert.throws(() => host.execute('studio.layout.compact', word()), /disabled/);
  assert.equal(createPluginHost({ plugins: builtinPlugins, storage }).list()[0].enabled, false);
  host.setEnabled('studio.layout', true);
  assert.equal(changes, 1);
});

test('malformed/unavailable persistence does not break plugin execution', () => {
  const host = createPluginHost({ plugins: builtinPlugins, storage: { getItem: () => '{broken', setItem: () => { throw new Error('Quota exceeded'); } } });
  host.setEnabled('studio.layout', false);
  assert.equal(host.list()[0].enabled, false);
});

test('trusted JavaScript example registers and edits through the actual host contract', () => {
  const host = createPluginHost({ plugins: [...builtinPlugins, editorialLayoutPlugin], storage: null });
  const output = host.execute('example.editorial.apply', word());
  assert.equal(output.document.data.typography.fontFamily, 'Georgia');
  assert.equal(output.document.data.typography.lineHeight, 1.75);
});

test('host enforces read-only permission and identity, and isolates failed mutations', () => {
  const plugin = { id: 'test.readonly', name: 'Read only', version: '1', description: '', permissions: ['document:read'], commands: [{ id: 'test.readonly.write', title: 'Bad writer', documentTypes: ['word'], run(document) { document.data.title = 'Changed'; return { message: 'changed', document }; } }] };
  const document = word();
  const host = createPluginHost({ plugins: [plugin], storage: null });
  assert.throws(() => host.execute('test.readonly.write', document), /document:write permission/);
  assert.equal(document.data.title, 'Notes.docx');
  const unregister = host.register({ ...plugin, id: 'test.identity', permissions: ['document:read', 'document:write'], commands: [{ ...plugin.commands[0], id: 'test.identity.write', run(document) { document.data.id = 'another'; return { message: 'changed', document }; } }] });
  assert.throws(() => host.execute('test.identity.write', document), /identity/);
  unregister();
  assert.throws(() => host.execute('test.identity.write', document), /Unknown/);
});

test('plugin ids and command namespaces cannot collide', () => {
  const host = createPluginHost({ plugins: builtinPlugins, storage: null });
  assert.throws(() => host.register(builtinPlugins[0]), /already registered/);
  assert.throws(() => host.register({ ...builtinPlugins[0], id: 'other.plugin' }), /must start with/);
});

test('context computes statistics from edited text and declares truncation', () => {
  const document = word();
  const context = collectDocumentContext(document, 8);
  assert.equal(context.statistics.words, 5);
  assert.equal(context.source, 'current-editor');
  assert.deepEqual(context.outline, ['Real heading']);
  assert.equal(context.text, 'Real hea');
  assert.equal(context.truncated, true);
  assert.equal(collectDocumentContext(document, -2).text, '');
});

test('context includes table text and current formulas without inventing totals', () => {
  const document = word();
  document.data.paragraphs = [{ id: 'p1', type: 'table', tableData: [['Item', 'Amount'], ['Tea', '12']] }];
  assert.equal(collectDocumentContext(document).text, 'Item\tAmount\nTea\t12');
  const sheet = { type: 'excel', data: { id: 'sheet', title: 'Budget.xlsx', sheets: [{ name: 'Costs', rowCount: 2, colCount: 2, cells: { A1: { value: 2 }, A2: { value: 4, formula: '=A1*2' } } }] } };
  const context = collectDocumentContext(sheet);
  assert.equal(context.statistics.formulas, 1);
  assert.match(context.text, /A2: =A1\*2 → 4/);
  assert.equal(context.statistics.populatedCells, 2);
});

test('diagnostics report actual JSON and heading issues', () => {
  assert.equal(getDocumentDiagnostics(code()).length, 0);
  const invalid = code(); invalid.data.content = '{oops';
  assert.equal(getDocumentDiagnostics(invalid)[0].severity, 'error');
  const document = word(); document.data.paragraphs[0].type = 'heading-3';
  assert.match(getDocumentDiagnostics(document)[0].message, /jumps/);
});
