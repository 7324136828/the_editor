import test from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { createMcpHandler, dispatchMcp } from '../server/mcp.mjs';

const document = { type: 'code', data: { id: 'notes-1', title: 'Notes.md', language: 'markdown', content: '# Current saved title\nActual saved content.', symbols: [] } };
const dependencies = { listDocuments: async () => [structuredClone(document)] };
const request = (method, params, id = 1) => ({ jsonrpc: '2.0', id, method, params });

test('MCP initialize negotiates protocol and advertises only available capabilities', async () => {
  const response = await dispatchMcp(request('initialize', { protocolVersion: '2025-11-25', clientInfo: { name: 'test', version: '1' }, capabilities: {} }), dependencies);
  assert.equal(response.result.protocolVersion, '2025-11-25');
  assert.ok(response.result.capabilities.tools);
  assert.equal(response.result.capabilities.resources.subscribe, false);
  assert.equal(response.result.capabilities.prompts, undefined);
  assert.equal((await dispatchMcp(request('initialize', {}), dependencies)).error.code, -32602);
});

test('MCP lists read-only tools and reads actual saved documents', async () => {
  const listing = await dispatchMcp(request('tools/list'), dependencies);
  assert.equal(listing.result.tools.length, 3);
  assert.ok(listing.result.tools.every(tool => tool.annotations.readOnlyHint));
  const response = await dispatchMcp(request('tools/call', { name: 'get_document', arguments: { id: 'notes-1' } }), dependencies);
  assert.deepEqual(response.result.structuredContent, document);
  assert.deepEqual(JSON.parse(response.result.content[0].text), document);
});

test('MCP resource list/read use exact URI matching with no file-path access', async () => {
  const listing = await dispatchMcp(request('resources/list'), dependencies);
  assert.equal(listing.result.resources[0].uri, 'workspace://document/notes-1');
  const read = await dispatchMcp(request('resources/read', { uri: listing.result.resources[0].uri }), dependencies);
  assert.deepEqual(JSON.parse(read.result.contents[0].text), document);
  assert.equal((await dispatchMcp(request('resources/read', { uri: 'file:///etc/passwd' }), dependencies)).error.code, -32002);
});

test('MCP context is explicitly saved, bounded, and has live content', async () => {
  const response = await dispatchMcp(request('tools/call', { name: 'document_context', arguments: { id: 'notes-1', maxCharacters: 9 } }), dependencies);
  const context = response.result.structuredContent;
  assert.equal(context.source, 'saved-workspace');
  assert.equal(context.text, '# Current');
  assert.equal(context.truncated, true);
});

test('MCP distinguishes invalid requests, bad arguments, and missing document tool results', async () => {
  assert.equal((await dispatchMcp([], dependencies)).error.code, -32600);
  assert.equal((await dispatchMcp(request('arbitrary/execute'), dependencies)).error.code, -32601);
  assert.equal((await dispatchMcp(request('tools/call', { name: 'get_document', arguments: { id: '../secret', script: 'bad' } }), dependencies)).error.code, -32602);
  assert.equal((await dispatchMcp(request('tools/call', { name: 'document_context', arguments: { id: 'notes-1', maxCharacters: -1 } }), dependencies)).error.code, -32602);
  assert.equal((await dispatchMcp(request('tools/call', { name: 'list_documents', arguments: { constructor: 'unexpected' } }), dependencies)).error.code, -32602);
  assert.equal((await dispatchMcp(request('tools/list', null), dependencies)).error.code, -32602);
  assert.equal((await dispatchMcp(request('tools/call', { name: 'get_document', arguments: { id: 'missing' } }), dependencies)).result.isError, true);
  assert.equal(await dispatchMcp({ jsonrpc: '2.0', method: 'notifications/initialized' }, dependencies), null);
  assert.equal((await dispatchMcp(request('tools/list', null, null), dependencies)).error.code, -32600);
});

test('MCP returns a sanitized error if saved workspace access fails', async () => {
  const response = await dispatchMcp(request('resources/list'), { listDocuments() { throw new Error('secret filesystem path'); } });
  assert.equal(response.error.code, -32603);
  assert.doesNotMatch(response.error.message, /secret/);
});

test('Streamable HTTP handles negotiation, notifications, invalid JSON, and Origin protection', async t => {
  const server = createServer(createMcpHandler(dependencies));
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => new Promise(resolve => { server.closeAllConnections(); server.close(resolve); }));
  const url = `http://127.0.0.1:${server.address().port}/mcp`;
  const headers = { 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-11-25' };
  const post = (body, extra = {}) => fetch(url, { method: 'POST', headers: { ...headers, ...extra }, body: typeof body === 'string' ? body : JSON.stringify(body) });
  const response = await post(request('tools/list'));
  assert.equal(response.status, 200);
  assert.equal((await response.json()).result.tools.length, 3);
  const notification = await post({ jsonrpc: '2.0', method: 'notifications/initialized' });
  assert.equal(notification.status, 202);
  assert.equal(await notification.text(), '');
  assert.equal((await fetch(url)).status, 405);
  assert.equal((await post('{malformed')).status, 400);
  assert.equal((await post(request('tools/list'), { Origin: 'https://hostile.example' })).status, 403);
  assert.equal((await post(request('tools/list'), { Accept: 'application/json' })).status, 406);
  assert.equal((await post(request('tools/list'), { 'MCP-Protocol-Version': 'bogus' })).status, 400);
});
