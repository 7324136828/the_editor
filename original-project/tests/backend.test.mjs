import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import http from 'node:http';
import ts from 'typescript';
import { createOfficeServer } from '../server/index.mjs';

const document = (id = 'doc-test', content = 'Hello world') => ({ type: 'code', data: { id, title: 'hello.ts', language: 'typescript', content, symbols: [] } });
async function fixture(t, withAssets = false) {
  const dataDir = await mkdtemp(path.join(tmpdir(), 'office-storage-test-'));
  const staticDir = withAssets ? path.join(dataDir, 'dist') : null;
  if (staticDir) {
    await mkdir(staticDir);
    await writeFile(path.join(staticDir, 'index.html'), '<!doctype html><title>Office Studio</title>');
    await writeFile(path.join(staticDir, 'app.js'), 'console.log("application");');
  }
  let server;
  const start = async () => {
    server = createOfficeServer({ dataDir, staticDir });
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    return `http://127.0.0.1:${server.address().port}`;
  };
  const stop = async () => {
    server.closeAllConnections();
    await new Promise(resolve => server.close(resolve));
  };
  let base = await start();
  t.after(async () => { await stop(); await rm(dataDir, { recursive: true, force: true }); });
  return {
    dataDir,
    request: (url, options) => fetch(base + url, options),
    rawStatus: (url, headers) => new Promise((resolve, reject) => { const request = http.get(base + url, { headers }, response => { response.resume(); resolve(response.statusCode); }); request.on('error', reject); }),
    put: (id, model, revision = null, name = 'hello.ts') => fetch(`${base}/api/documents/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, document: model, revision }) }),
    restart: async () => { await stop(); base = await start(); },
  };
}

test('save survives server restart; history preserves previous contents', async t => {
  const app = await fixture(t);
  const saved = await app.put('doc-test', document());
  assert.equal(saved.status, 201);
  assert.equal((await saved.json()).revision, 1);
  const updated = await app.put('doc-test', document('doc-test', 'Revised content'), 1);
  assert.equal(updated.status, 200);
  await app.restart();
  const all = await (await app.request('/api/documents')).json();
  assert.equal(all.documents.length, 1);
  assert.equal(all.documents[0].revision, 2);
  assert.equal(all.documents[0].document.data.content, 'Revised content');
  const versions = await (await app.request('/api/documents/doc-test/history')).json();
  assert.deepEqual(versions.versions.map(item => item.revision), [2, 1]);
  const original = await (await app.request('/api/documents/doc-test/history/1')).json();
  assert.equal(original.document.data.content, 'Hello world');
  const restored = await app.put('doc-test', original.document, 2);
  assert.equal((await restored.json()).revision, 3);
  assert.deepEqual(await readdir(app.dataDir), ['doc-test.json']);
});

test('concurrent updates serialize and reject stale revisions without lost writes', async t => {
  const app = await fixture(t);
  await app.put('doc-test', document());
  const responses = await Promise.all([app.put('doc-test', document('doc-test', 'First'), 1), app.put('doc-test', document('doc-test', 'Second'), 1)]);
  assert.deepEqual(responses.map(result => result.status).sort(), [200, 409]);
  const conflict = await responses.find(result => result.status === 409).json();
  assert.equal(conflict.current.revision, 2);
  const staleNew = await app.put('doc-test', document());
  assert.equal(staleNew.status, 409);
  const history = await (await app.request('/api/documents/doc-test/history')).json();
  assert.equal(history.versions.length, 2);
});

test('invalid paths, origins, hosts, content types, models and JSON cannot write storage', async t => {
  const app = await fixture(t);
  assert.equal((await app.put('..%2Foutside', document())).status, 400);
  assert.equal((await app.put('doc-test', document('mismatched'))).status, 400);
  assert.equal((await app.put('doc-test', document(), null, '../escape.ts')).status, 400);
  assert.equal((await app.put('doc-test', { type: 'word', data: { id: 'doc-test', title: 'Broken' } })).status, 400);
  assert.equal((await app.request('/api/documents', { headers: { Origin: 'https://untrusted.example' } })).status, 403);
  assert.equal(await app.rawStatus('/api/documents', { Host: 'untrusted.example' }), 403);
  assert.equal((await app.request('/api/documents/doc-test', { method: 'PUT', headers: { 'Content-Type': 'text/plain' }, body: '{}' })).status, 415);
  assert.equal((await app.request('/api/documents/doc-test', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: '{bad' })).status, 400);
  assert.equal((await app.request('/api/documents/missing')).status, 404);
  assert.deepEqual(await readdir(app.dataDir), []);
});

test('history is bounded to the latest 50 successful revisions', async t => {
  const app = await fixture(t);
  for (let revision = 0; revision < 52; revision++) {
    const result = await app.put('doc-test', document('doc-test', `Revision ${revision + 1}`), revision || null);
    assert.ok(result.ok);
  }
  const history = await (await app.request('/api/documents/doc-test/history')).json();
  assert.equal(history.versions.length, 50);
  assert.equal(history.versions[0].revision, 52);
  assert.equal(history.versions[49].revision, 3);
  assert.equal((await app.request('/api/documents/doc-test/history/1')).status, 404);
});

test('production static server serves built app, routes and assets without exposing storage', async t => {
  const app = await fixture(t, true);
  await app.put('doc-test', document());
  const root = await app.request('/');
  assert.equal(root.status, 200);
  assert.match(root.headers.get('content-type'), /text\/html/);
  assert.match(await root.text(), /Office Studio/);
  const asset = await app.request('/app.js');
  assert.equal(asset.status, 200);
  assert.match(asset.headers.get('content-type'), /javascript/);
  assert.match(await asset.text(), /application/);
  assert.equal((await app.request('/workspace/example')).status, 200);
  assert.equal((await app.request('/app.js', { method: 'HEAD' })).status, 200);
  assert.equal((await app.request('/doc-test.json')).status, 404);
  assert.equal((await app.request('/..%2fdoc-test.json')).status, 400);
  assert.equal((await app.request('/api/unknown')).status, 404);
});

test('all shipped Office model fixtures pass backend validation and persist without dropping fields', async t => {
  const app = await fixture(t);
  const source = await readFile(new URL('../src/data/sampleDocuments.ts', import.meta.url), 'utf8');
  const code = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.ESNext } }).outputText;
  const samples = await import('data:text/javascript;base64,' + Buffer.from(code).toString('base64'));
  for (const [type, key] of [['word', 'sampleWordDocument'], ['excel', 'sampleExcelDocument'], ['powerpoint', 'samplePptDocument']]) {
    const model = { type, data: samples[key] };
    const response = await app.put(model.data.id, model, null, model.data.title);
    const result = await response.json();
    assert.equal(response.status, 201, result.error);
    assert.deepEqual(result.document, model);
  }
});
