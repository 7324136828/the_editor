import test from 'node:test';
import assert from 'node:assert/strict';
import { importTypeScript } from './plugins-test-utils.mjs';

const { searchDocuments, replaceInDocument, createMatcher } = await importTypeScript(new URL('../src/services/search.ts', import.meta.url));
const code = content => ({ type: 'code', data: { id: 'code-1', title: 'example.md', language: 'markdown', content, symbols: [] } });
const word = paragraphs => ({ type: 'word', data: { id: 'word-1', title: 'example.docx', paragraphs, headings: [], sections: [], comments: [], typography: { fontFamily: 'Arial', fontSize: 14, lineHeight: 1.5, marginSize: 'normal', textColor: '#000000' }, stats: {} } });
const excel = cells => ({ type: 'excel', data: { id: 'excel-1', title: 'example.xlsx', activeSheetId: 'sheet-1', formulasAudit: [], sheets: [{ id: 'sheet-1', name: 'Sheet1', rowCount: 20, colCount: 3, columns: [], cells }] } });
const files = (...documents) => documents.map(doc => ({ id: doc.data.id, name: doc.data.title, doc }));

test('plain search escapes regular-expression punctuation', () => {
  const matches = searchDocuments(files(code('price = $1.50; [x] (a+b)')), '$1.50');
  assert.equal(matches.length, 1);
  assert.equal(matches[0].matchLength, 5);
  assert.equal(searchDocuments(files(code('axb')), 'a.b').length, 0);
});

test('case and whole-word choices apply consistently to search and replacement', () => {
  const doc = code('Cat cat scatter catfish');
  const matches = searchDocuments(files(doc), 'cat', { matchCase: true, wholeWord: true });
  assert.equal(matches.length, 1);
  const replaced = replaceInDocument(doc, 'cat', 'dog', { matchCase: true, wholeWord: true });
  assert.equal(replaced.count, 1);
  assert.equal(replaced.document.data.content, 'Cat dog scatter catfish');
});

test('regex syntax errors and overlong queries are reported', () => {
  assert.throws(() => createMatcher('[', { useRegex: true }));
  assert.throws(() => createMatcher('a'.repeat(201)));
  assert.deepEqual(searchDocuments(files(code('hello')), ''), []);
});

test('replacement text is literal, including dollar replacement tokens', () => {
  const doc = code('value value');
  const replaced = replaceInDocument(doc, 'value', '$&$1');
  assert.equal(replaced.document.data.content, '$&$1 $&$1');
  assert.equal(replaced.count, 2);
  assert.equal(doc.data.content, 'value value');
});

test('anchored code regex uses the same line scope for find and replace', () => {
  const document = code('cat\nother\ncat');
  assert.equal(searchDocuments(files(document), '^cat$', { useRegex: true }).length, 2);
  const replaced = replaceInDocument(document, '^cat$', 'dog', { useRegex: true });
  assert.equal(replaced.count, 2);
  assert.equal(replaced.document.data.content, 'dog\nother\ndog');
});

test('search reports every occurrence with original text offsets', () => {
  const matches = searchDocuments(files(code('cat cat CAT')), 'cat');
  assert.equal(matches.length, 3);
  assert.deepEqual(matches.map(match => match.matchIndex), [0, 4, 8]);
  assert.deepEqual(matches.map(match => match.targetRef.lineNumber), [1, 1, 1]);
});

test('search visits Word tables, spreadsheet formulas, and presentation notes', () => {
  const documents = files(
    word([{ id: 'p1', type: 'table', tableData: [['target', 'unrelated']] }]),
    excel({ A1: { value: 2 }, A2: { value: 2, formula: '=SUM(A1)' } }),
    { type: 'powerpoint', data: { id: 'ppt-1', title: 'Deck.pptx', slides: [{ id: 's1', title: 'A title', notes: 'target speaker notes', objects: [] }] } },
  );
  const matches = searchDocuments(documents, 'target');
  assert.equal(matches.length, 2);
  assert.equal(matches[0].targetRef.headingId, 'p1');
  assert.equal(matches[1].targetRef.slideIndex, 1);
  assert.equal(searchDocuments(documents, 'SUM')[0].targetRef.cellCoord, 'A2');
});

test('search and replace include SmartArt titles and item labels', () => {
  const document = word([{ id: 'diagram', type: 'smartart', smartArt: { title: 'Launch target', layout: 'process', accentColor: '#2b6cb0', items: [{ id: 'step', text: 'Review target metrics' }] } }]);
  const matches = searchDocuments(files(document), 'target');
  assert.deepEqual(matches.map(match => match.location), ['SmartArt 1: title', 'SmartArt 1, item 1']);
  const result = replaceInDocument(document, 'target', 'outcome');
  assert.equal(result.count, 2);
  assert.equal(result.document.data.paragraphs[0].smartArt.title, 'Launch outcome');
  assert.equal(result.document.data.paragraphs[0].smartArt.items[0].text, 'Review outcome metrics');
  assert.equal(document.data.paragraphs[0].smartArt.title, 'Launch target');
});

test('spreadsheet text replacement retains string types and leading zeros', () => {
  const document = excel({ A1: { value: '00123' }, A2: { value: '123' } });
  const result = replaceInDocument(document, '123', '456');
  assert.equal(result.document.data.sheets[0].cells.A1.value, '00456');
  assert.equal(result.document.data.sheets[0].cells.A2.value, '456');
  assert.equal(document.data.sheets[0].cells.A1.value, '00123');
});

test('Word replacement preserves formatting in untouched runs', () => {
  const document = word([{ id: 'p1', type: 'body', text: 'Hello world', runs: [{ text: 'Hello ', bold: true }, { text: 'world', italic: true }] }]);
  const result = replaceInDocument(document, 'Hello', 'Welcome');
  assert.equal(result.count, 1);
  assert.equal(result.document.data.paragraphs[0].runs.map(run => run.text).join(''), 'Welcome world');
  assert.ok(result.document.data.paragraphs[0].runs.some(run => run.text === 'world' && run.italic));
  assert.equal(document.data.paragraphs[0].runs[0].text, 'Hello ');
});

test('unsafe regex constructs are rejected before searching document text', () => {
  // These would otherwise execute inside React's synchronous render path.
  for (const pattern of ['(a+)+$', '(a|aa)+$', '(a)\\1+', 'a+a+$', 'a+$']) {
    assert.throws(() => createMatcher(pattern, { useRegex: true }), undefined, `Expected an unsafe-pattern error for ${pattern}`);
  }
});
