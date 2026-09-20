import { test } from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { DOMParser } from '@xmldom/xmldom';
import { serializeDocument } from '../src/services/exportDocument.ts';

async function packageFrom(document) {
  const output = await serializeDocument(document);
  assert.deepEqual([...output.data.subarray(0, 2)], [80, 75], 'Office document must be a ZIP package');
  const zip = await JSZip.loadAsync(output.data);
  assert.ok(zip.file('[Content_Types].xml'));
  assert.ok(zip.file('_rels/.rels'));
  const parser = new DOMParser({ onError: (level, message) => { throw new Error(`${level}: ${message}`); } });
  for (const [name, file] of Object.entries(zip.files)) {
    if (name.endsWith('.xml') || name.endsWith('.rels')) assert.ok(parser.parseFromString(await file.async('text'), 'application/xml').documentElement, `${name} must be valid XML`);
  }
  // Every internal OPC relationship must resolve to a real package part.
  for (const [name, file] of Object.entries(zip.files)) {
    if (!name.endsWith('.rels')) continue;
    const sourceDir = name === '_rels/.rels' ? '' : path.posix.dirname(path.posix.dirname(name));
    const text = await file.async('text');
    for (const target of text.matchAll(/Target="([^"]+)"/g)) {
      const resolved = path.posix.normalize(path.posix.join(sourceDir, target[1]));
      assert.ok(zip.file(resolved), `${name} target ${target[1]} should exist`);
    }
  }
  return { zip, output };
}

test('DOCX exports edited text, escaping, rich runs, headings, tables and layout', async () => {
  const { zip, output } = await packageFrom({ type: 'word', data: {
    id: 'word-test', title: 'Report.docx', author: 'Author', createdAt: '', modifiedAt: '', headings: [], sections: [{ id: 'section-1', title: 'Main', pageNumber: 1, orientation: 'landscape' }], comments: [],
    paragraphs: [{ id: 'p1', type: 'heading-1', text: 'Edited <Report> & Summary' }, { id: 'p2', type: 'body', runs: [{ text: 'Grüße', bold: true }, { text: '\nSecond line', italic: true }] }, { id: 'table', type: 'table', tableData: [['Region', 'Total'], ['East', '120']] }],
    typography: { fontFamily: 'Aptos', fontSize: 16, lineHeight: 1.5, marginSize: 'narrow', textColor: '#123456' }, stats: {},
  } });
  assert.equal(output.extension, 'docx');
  const body = await zip.file('word/document.xml').async('text');
  assert.match(body, /Edited &lt;Report&gt; &amp; Summary/);
  assert.match(body, /w:pStyle w:val="Heading1"/);
  assert.match(body, /<w:b\/>/);
  assert.match(body, /<w:i\/>/);
  assert.match(body, /<w:br\/>/);
  assert.match(body, /<w:tbl>/);
  assert.match(body, /w:orient="landscape"/);
  assert.match(body, /w:top="720"/);
});

test('DOCX exports SmartArt as visible editable cells with structured round-trip metadata', async () => {
  const { zip } = await packageFrom({ type: 'word', data: {
    id: 'word-smartart', title: 'Diagram.docx', author: 'Author', createdAt: '', modifiedAt: '', headings: [], sections: [{ id: 'section-1', title: 'Main', pageNumber: 1, orientation: 'portrait' }], comments: [],
    paragraphs: [{ id: 'diagram-1', type: 'smartart', smartArt: { title: 'Release process', layout: 'process', accentColor: '#2b6cb0', items: [{ id: 'step-1', text: 'Plan' }, { id: 'step-2', text: 'Ship & learn' }] } }],
    typography: { fontFamily: 'Aptos', fontSize: 14, lineHeight: 1.5, marginSize: 'normal', textColor: '#123456' }, stats: {},
  } });
  const body = await zip.file('word/document.xml').async('text');
  assert.match(body, /Release process/);
  assert.match(body, /Ship &amp; learn/);
  assert.match(body, /w:shd w:fill="2B6CB0"/);
  const metadata = await zip.file('customXml/officeStudioSmartArt.xml').async('text');
  assert.match(metadata, /officeStudioSmartArt/);
  assert.match(metadata, /&quot;layout&quot;:&quot;process&quot;/);
});

test('XLSX round-trips values and editable formulas in multiple sheets', async () => {
  const makeSheet = (id, name) => ({ id, name, rowCount: 10, colCount: 3, columns: [{ key: 'A', label: 'A', width: 140 }], cells: { A1: { value: 'Revenue ✓' }, A2: { value: 30 }, A3: { value: 12.5, format: { numberFormat: 'currency', decimalPlaces: 2 } }, B2: { value: 42.5, formula: '=SUM(A2:A3)' }, C2: { value: true } } });
  const { output } = await packageFrom({ type: 'excel', data: { id: 'xlsx-test', title: 'Book.xlsx', activeSheetId: 'one', formulasAudit: [], sheets: [makeSheet('one', 'Sales'), makeSheet('two', 'Sales')] } });
  const workbook = XLSX.read(output.data, { type: 'array', cellFormula: true, cellNF: true });
  assert.deepEqual(workbook.SheetNames, ['Sales', 'Sales (2)']);
  assert.equal(workbook.Sheets.Sales.A1.v, 'Revenue ✓');
  assert.equal(workbook.Sheets.Sales.A2.v, 30);
  assert.equal(workbook.Sheets.Sales.C2.v, true);
  assert.equal(workbook.Sheets.Sales.B2.f, 'SUM(A2:A3)');
  assert.equal(workbook.Sheets.Sales.B2.v, 42.5);
  assert.equal(workbook.Sheets.Sales.A3.z, '$#,##0.00');
});

test('PPTX packages editable shapes, coordinates, notes, master, theme and correct aspect ratio', async () => {
  const { zip, output } = await packageFrom({ type: 'powerpoint', data: { id: 'ppt-test', title: 'Deck.pptx', activeSlideId: 's1', themeName: 'Office Studio', accentColor: '#007acc', slides: [{ id: 's1', slideNumber: 1, title: 'Edited title', layout: 'blank', background: '#ffffff', transition: 'fade', notes: 'Remember <this> & that', objects: [{ id: 'o1', kind: 'text', x: 50, y: 80, width: 300, height: 100, text: 'Edited presentation', fontSize: 24, color: '#123456', fill: 'transparent', zIndex: 1 }] }] } });
  assert.equal(output.extension, 'pptx');
  const slide = await zip.file('ppt/slides/slide1.xml').async('text');
  assert.match(slide, /Edited presentation/);
  assert.match(slide, /<a:off x="476250" y="762000"\/>/);
  assert.match(slide, /<p:fade\/>/);
  const notes = await zip.file('ppt/notesSlides/notesSlide1.xml').async('text');
  assert.match(notes, /Remember &lt;this&gt; &amp; that/);
  const presentation = await zip.file('ppt/presentation.xml').async('text');
  assert.match(presentation, /cx="8382000" cy="4714875" type="screen16x9"/);
  assert.ok(zip.file('ppt/theme/theme1.xml'));
  assert.ok(zip.file('ppt/slideMasters/slideMaster1.xml'));
});

test('text documents export exact UTF-8 source rather than a model wrapper', async () => {
  const content = 'const greeting = "hello ✓";\n';
  const output = await serializeDocument({ type: 'code', data: { id: 'code-test', title: 'test.ts', language: 'typescript', content, symbols: [] } });
  assert.equal(new TextDecoder().decode(output.data), content);
  assert.equal(output.extension, 'ts');
});
