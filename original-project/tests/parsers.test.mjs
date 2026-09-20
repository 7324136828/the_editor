import { test } from 'node:test';
import assert from 'node:assert/strict';
import { registerHooks } from 'node:module';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { serializeDocument } from '../src/services/exportDocument.ts';

// Browser builds resolve these extensionless TS imports; teach Node's native TS loader the same.
registerHooks({ resolve(specifier, context, nextResolve) { try { return nextResolve(specifier, context); } catch (error) { if (error.code === 'ERR_MODULE_NOT_FOUND' && specifier.startsWith('.')) return nextResolve(specifier + '.ts', context); throw error; } } });
globalThis.DOMParser = DOMParser;
const { parseDocxFile } = await import('../src/parsers/docxParser.ts');
const { parsePptxFile } = await import('../src/parsers/pptxParser.ts');
const { parseXlsxFile } = await import('../src/parsers/xlsxParser.ts');
const file = (data, name) => new File([data], name);

test('DOCX export/import preserves rich text, tables, typography, orientation and comments', async () => {
  const source = { type: 'word', data: { id: 'doc-word', title: 'test.docx', headings: [], author: 'Test', createdAt: '', modifiedAt: '', stats: {}, sections: [{ id: 's', title: 'Main', pageNumber: 1, orientation: 'landscape' }],
    typography: { fontFamily: 'Aptos', fontSize: 16, lineHeight: 1.5, marginSize: 'narrow', textColor: '#123456' },
    paragraphs: [{ id: 'p1', type: 'heading-1', text: 'Heading', align: 'center' }, { id: 'p2', type: 'body', runs: [{ text: 'Bold & ', bold: true }, { text: 'italic\nsecond line', italic: true }] }, { id: 'p3', type: 'table', tableData: [['A', 'B'], ['1', '2']] }],
    comments: [{ id: 'c1', author: 'Reviewer', avatarColor: '#fff', timestamp: '', selectedText: 'Bold', comment: 'Check the numbers.', resolved: false }],
  } };
  const output = await serializeDocument(source);
  const imported = await parseDocxFile(file(output.data, 'test.docx'));
  assert.equal(imported.paragraphs.length, 3, 'comment anchors must not duplicate body text');
  assert.equal(imported.paragraphs[0].type, 'heading-1');
  assert.equal(imported.paragraphs[0].align, 'center');
  assert.equal(imported.paragraphs[1].runs[0].bold, true);
  assert.equal(imported.paragraphs[1].runs[1].italic, true);
  assert.equal(imported.paragraphs[1].text, 'Bold & italic\nsecond line');
  assert.deepEqual(imported.paragraphs[2].tableData, [['A', 'B'], ['1', '2']]);
  assert.equal(imported.sections[0].orientation, 'landscape');
  assert.deepEqual(imported.typography, source.data.typography);
  assert.equal(imported.comments[0].comment, 'Check the numbers.');
  assert.equal(imported.comments[0].author, 'Reviewer');
  assert.match(imported.comments[0].selectedText, /Bold/);
});

test('DOCX export/import retains editable SmartArt structure and text', async () => {
  const source = { type: 'word', data: { id: 'doc-smartart', title: 'diagram.docx', headings: [], author: 'Test', createdAt: '', modifiedAt: '', stats: {}, sections: [{ id: 's', title: 'Main', pageNumber: 1, orientation: 'portrait' }],
    typography: { fontFamily: 'Aptos', fontSize: 14, lineHeight: 1.5, marginSize: 'normal', textColor: '#123456' }, comments: [],
    paragraphs: [{ id: 'diagram-1', type: 'smartart', smartArt: { title: 'Launch plan', layout: 'cycle', accentColor: '#0f766e', items: [{ id: 'one', text: 'Discover' }, { id: 'two', text: 'Deliver' }] } }],
  } };
  const output = await serializeDocument(source);
  const imported = await parseDocxFile(file(output.data, 'diagram.docx'));
  assert.equal(imported.paragraphs.length, 1);
  assert.deepEqual(imported.paragraphs[0], source.data.paragraphs[0]);
  assert.equal(imported.stats.words, 4);
});

test('PPTX export/import preserves shape positions, color, title, text and notes', async () => {
  const source = { type: 'powerpoint', data: { id: 'doc-ppt', title: 'test.pptx', activeSlideId: 's1', themeName: 'Office Studio', accentColor: '#007ACC', slides: [{ id: 's1', slideNumber: 1, title: 'A real title', layout: 'title', background: '#EFEFEF', transition: 'fade', notes: 'Speaker notes\nSecond line', objects: [{ id: 'o1', kind: 'text', x: 25, y: 45, width: 420, height: 60, text: 'Changed title\nSubtitle', fill: 'transparent', color: '#123456', fontSize: 24, fontWeight: 'bold', align: 'center', rotation: 15, zIndex: 1 }] }] } };
  const output = await serializeDocument(source);
  const imported = await parsePptxFile(file(output.data, 'test.pptx'));
  assert.equal(imported.slides.length, 1);
  const slide = imported.slides[0];
  assert.equal(slide.title, source.data.slides[0].title);
  assert.equal(slide.notes, 'Speaker notes\nSecond line');
  assert.equal(slide.transition, 'fade');
  assert.equal(slide.background, '#EFEFEF');
  const object = slide.objects[0];
  for (const key of ['x', 'y', 'width', 'height', 'text', 'color', 'fill', 'fontSize', 'fontWeight', 'align', 'rotation']) assert.equal(object[key], source.data.slides[0].objects[0][key], key);
});

test('PPTX importer respects explicit slide relationship order', async () => {
  const source = { type: 'powerpoint', data: { id: 'ppt', title: 'test', themeName: '', accentColor: '#000000', activeSlideId: 's1', slides: [1, 2].map(number => ({ id: `s${number}`, slideNumber: number, title: `Slide ${number}`, layout: 'blank', objects: [], notes: '', background: '#ffffff', transition: 'none' })) } };
  const output = await serializeDocument(source);
  const zip = await JSZip.loadAsync(output.data);
  const xml = await zip.file('ppt/presentation.xml').async('text');
  zip.file('ppt/presentation.xml', xml.replace('<p:sldId id="256" r:id="rId4"/><p:sldId id="257" r:id="rId5"/>', '<p:sldId id="257" r:id="rId5"/><p:sldId id="256" r:id="rId4"/>'));
  const imported = await parsePptxFile(file(await zip.generateAsync({ type: 'uint8array' }), 'test.pptx'));
  assert.deepEqual(imported.slides.map(slide => slide.title), ['Slide 2', 'Slide 1']);
});

test('XLSX used ranges retain absolute D30 position, formulas and column widths', async () => {
  const workbook = XLSX.utils.book_new();
  const sheet = { D30: { t: 'n', v: 42 }, E30: { t: 'n', v: 84, f: 'D30*2', z: '$#,##0.00' }, '!ref': 'D30:E30', '!cols': [{ wpx: 150 }] };
  XLSX.utils.book_append_sheet(workbook, sheet, 'Offset');
  const imported = await parseXlsxFile(file(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }), 'offset.xlsx'));
  assert.equal(imported.sheets[0].rowCount, 30);
  assert.equal(imported.sheets[0].columns[0].key, 'A');
  assert.equal(imported.sheets[0].cells.D30.value, 42);
  assert.equal(imported.sheets[0].cells.E30.formula, '=D30*2');
  assert.equal(imported.sheets[0].cells.E30.format.numberFormat, 'currency');
  assert.equal(imported.sheets[0].columns[0].width, 150);
});

test('invalid packages and huge sheet ranges fail clearly instead of fabricating content', async () => {
  await assert.rejects(parseDocxFile(file('not a zip', 'bad.docx')), /not a readable Office/);
  await assert.rejects(parsePptxFile(file('not a zip', 'bad.pptx')), /not a readable Office/);
  await assert.rejects(parseXlsxFile(file('pretend office', 'bad.xlsx')), /do not match/);
  const empty = new JSZip(); empty.file('[Content_Types].xml', '<Types/>');
  const emptyBytes = await empty.generateAsync({ type: 'uint8array' });
  await assert.rejects(parseDocxFile(file(emptyBytes, 'empty.docx')), /missing its required word/);
  await assert.rejects(parsePptxFile(file(emptyBytes, 'empty.pptx')), /missing its required ppt/);
  const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, { A1001: { t: 's', v: 'Too far' }, '!ref': 'A1:A1001' }, 'Huge');
  await assert.rejects(parseXlsxFile(file(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' }), 'huge.xlsx')), /exceeds the editor limit/);
});
