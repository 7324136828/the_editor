import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createServer } from 'vite';
import { createOfficeServer } from '../server/index.mjs';

const prefix = path.join(tmpdir(), 'office-studio-e2e-');
const dataDir = await mkdtemp(prefix);
const backend = createOfficeServer({ dataDir, allowedOrigins: ['http://127.0.0.1:5185'] });
await new Promise(resolve => backend.listen(3102, '127.0.0.1', resolve));
const frontend = await createServer({ server: { host: '127.0.0.1', port: 5185, strictPort: true, proxy: { '/api': 'http://127.0.0.1:3102', '/mcp': 'http://127.0.0.1:3102' } } });
await frontend.listen();
console.log('Isolated browser test workspace ready at http://127.0.0.1:5185');
let closing = false;
async function shutdown() {
  if (closing) return;
  closing = true;
  await frontend.close();
  await new Promise(resolve => backend.close(resolve));
  const resolved = path.resolve(dataDir);
  if (resolved.startsWith(path.resolve(prefix)) && path.dirname(resolved) === path.resolve(tmpdir())) await rm(resolved, { recursive: true, force: true });
  process.exit(0);
}
process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
