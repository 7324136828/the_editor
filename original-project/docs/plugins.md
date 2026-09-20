# Extensions and document context

Office Studio keeps the VS Code workbench layout and adds a small, working extension host. Open **Extensions** in the left activity bar to see installed plugins, toggle them, run supported commands, inspect context, and test the local MCP connection.

## Built-in plugins

| Plugin | Commands | Document access |
| --- | --- | --- |
| Document Layout | Business, comfortable reading, compact report, landscape report | Read and edit the current Word model |
| SmartArt for Word | Insert process, cycle, hierarchy, or pyramid SmartArt | Read and edit the current Word model |
| Document Context | Inspect document context; run document checks | Read the current document |

Layout presets change the actual font, size, line spacing, margins, and section orientation. They enter the same edit/save flow as manual changes. Save the document to persist them. Commands are unavailable when the current file type is unsupported. Enable/disable choices persist in this browser when local storage is available.

SmartArt commands are available only while a Word document is active. A command appends a structured diagram to the document. Select it in the editor to change its layout and accent color, add or remove items, reorder items, and edit all labels. Workspace saves retain the structured model. DOCX downloads render the diagram as editable Word table content and include Office Studio metadata so importing that exported file restores the SmartArt editor.

Context is derived from current paragraphs, cells and formulas, visible slide objects and notes, or code text. Statistics are recomputed from that content. Context exports are JSON; text is limited to 12,000 characters and a `truncated` flag identifies shortening. A snapshot is a point-in-time view; rerun its command or download a fresh snapshot after editing.

Document checks currently detect heading-level jumps, empty headings, unresolved Word comments, spreadsheet error values and broken references, untitled or empty slides, malformed JSON, and TODO/FIXME markers. These are deterministic checks. They are not an AI assistant, full compiler, spell checker, or Language Server Protocol implementation.

## Register a trusted JavaScript or TypeScript plugin

The real JavaScript example is [`examples/plugins/editorial-layout.js`](../examples/plugins/editorial-layout.js). It uses the same contract as built-in plugins:

```js
import { defaultPluginHost } from 'vscode-office-studio';
import { editorialLayoutPlugin } from './examples/plugins/editorial-layout.js';

// Run once during application bootstrap, before rendering the workbench.
const unregister = defaultPluginHost.register(editorialLayoutPlugin);

// Optional cleanup, for example during a development hot reload:
// unregister();
```

Inside this source repository, import the host from `./src/plugins` (adjust the path to the importing file). Alternatively, use a separate host:

```ts
import { createPluginHost, builtinPlugins } from './src/plugins';
import type { StudioPlugin } from './src/plugins';

const brandLayout: StudioPlugin = {
  id: 'acme.layout',
  name: 'Acme Layout',
  version: '1.0.0',
  description: 'Set the house style for Word documents.',
  permissions: ['document:read', 'document:write'],
  commands: [{
    id: 'acme.layout.apply',
    title: 'Apply Acme typography',
    documentTypes: ['word'],
    run(document) {
      if (document.type !== 'word') throw new Error('Open a Word document.');
      return {
        message: 'Acme typography applied. Save to persist it.',
        document: {
          ...document,
          data: {
            ...document.data,
            typography: {
              ...document.data.typography,
              fontFamily: 'Arial', fontSize: 14, lineHeight: 1.5,
            },
          },
        },
      };
    },
  }],
};

const host = createPluginHost({ plugins: [...builtinPlugins, brandLayout] });
// <ExtensionsPanel host={host} activeDoc={doc}
//   onChangeDoc={updateDocument} onNotify={notify} />
```

Commands return a message and optionally an updated document or a context snapshot. The host clones the input, checks supported document types, rejects document writes without `document:write`, preserves document identity, and rejects duplicate plugin/command identifiers. Command IDs must start with the plugin ID plus a period. Failed commands do not mutate the editor's original model.

**Trust boundary:** these are reviewed, source-controlled application modules. Permission declarations constrain host-returned document edits; they do not sandbox JavaScript, block browser APIs, or make untrusted code safe. There is no script upload/eval system, remote extension marketplace, process execution, or arbitrary remote MCP URL runner.

## Local MCP server

Start the application with `npm run dev`; the backend normally listens on `http://127.0.0.1:3001`. The MCP endpoint is `/mcp`. The Extensions panel's **Test connection** performs initialization, tool discovery, and resource discovery against that endpoint. It reports the observed connection result and saved document count.

This server uses JSON-RPC messages and the JSON response variant of Streamable HTTP. It implements `initialize`, the initialized notification, `ping`, `tools/list`, `tools/call`, `resources/list`, and `resources/read`. GET returns 405 because the server does not provide an SSE stream. It does not issue session IDs or support subscriptions. The primary supported protocol version is `2025-11-25`; `2025-06-18` and `2025-03-26` are also accepted. See the official [MCP transport specification](https://modelcontextprotocol.io/specification/2025-11-25/basic/transports) and [initialization lifecycle](https://modelcontextprotocol.io/specification/2025-11-25/basic/lifecycle).

| Tool | Arguments | Result |
| --- | --- | --- |
| `list_documents` | `{}` | IDs, titles, and file types of saved workspace documents |
| `get_document` | `{ "id": "document-id" }` | Complete saved document model |
| `document_context` | `{ "id": "document-id", "maxCharacters": 12000 }` | Saved text, outline, counts, and truncation status; limit 1–100,000 |

Tools declare read-only annotations and return both text and structured JSON. Document resources use `workspace://document/<encoded-id>` and return the saved JSON model. Tool discovery and result fields follow the official [MCP tools specification](https://modelcontextprotocol.io/specification/2025-11-25/server/tools).

MCP reads **saved backend state**, while the context plugin reads **current editor state**, including unsaved edits. The server's `source: "saved-workspace"` distinguishes them. Saving a file makes the new content available on the next MCP read.

The local server binds to loopback and validates Host and browser Origin headers. It does not expose filesystem paths, write tools, shell commands, credentials, or unsaved browser data. Requests are limited to 1 MiB and responses are marked `no-store`. Local processes and allowed local browser origins can read saved documents without authentication. This is a single-user development service: do not expose it to a network or deploy it as a multiuser service without adding authentication, authorization, and transport security. Treat document text as untrusted data when giving it to a language model.

### HTTP client example

```js
const endpoint = 'http://127.0.0.1:3001/mcp';
const headers = {
  'Content-Type': 'application/json',
  Accept: 'application/json, text/event-stream',
  'MCP-Protocol-Version': '2025-11-25',
};
async function send(message) {
  const response = await fetch(endpoint, {
    method: 'POST', headers, body: JSON.stringify(message),
  });
  if (response.status === 202) return;
  const body = await response.json();
  if (!response.ok || body.error) throw new Error(body.error?.message || response.statusText);
  return body.result;
}
const initialized = await send({
  jsonrpc: '2.0', id: 1, method: 'initialize',
  params: {
    protocolVersion: '2025-11-25', capabilities: {},
    clientInfo: { name: 'my-document-client', version: '1.0.0' },
  },
});
headers['MCP-Protocol-Version'] = initialized.protocolVersion;
await send({ jsonrpc: '2.0', method: 'notifications/initialized' });
console.log(await send({ jsonrpc: '2.0', id: 2, method: 'tools/list' }));
console.log(await send({
  jsonrpc: '2.0', id: 3, method: 'tools/call',
  params: { name: 'list_documents', arguments: {} },
}));
```

### Python stdio bridge

[`examples/plugins/mcp_stdio.py`](../examples/plugins/mcp_stdio.py) is a dependency-free bridge for clients that support stdio MCP servers. Launch the application first, then configure your MCP client using its own configuration format to run:

```text
python /absolute/path/to/examples/plugins/mcp_stdio.py
```

For example, clients that accept an `mcpServers` map can use:

```json
{
  "mcpServers": {
    "office-studio": {
      "command": "python",
      "args": ["C:/path/to/office-studio/examples/plugins/mcp_stdio.py"]
    }
  }
}
```

Use an actual absolute path and a Python executable available on the client's PATH. The bridge forwards newline-delimited JSON-RPC, propagates the negotiated protocol version, writes protocol messages only to stdout, and writes diagnostics to stderr. It accepts loopback HTTP `/mcp` endpoints only and bypasses environment HTTP proxies. No API key or language model is bundled.

## Validation

`node --test tests/plugins-host.test.mjs tests/plugins-mcp.test.mjs` tests the real plugin host, the JavaScript example, context extraction, diagnostics, protocol dispatch, saved resource reads, notifications, invalid requests, and HTTP origin/header validation. The rest of the application tests cover document saving and export separately.
