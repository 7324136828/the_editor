"""MCP (Model Context Protocol) JSON-RPC handler for Code Office Studio."""

from __future__ import annotations

from typing import Any
from .document_store import DocumentStore


class McpService:
    """Handles JSON-RPC requests for the /mcp endpoint."""

    def __init__(self, doc_store: DocumentStore) -> None:
        self.doc_store = doc_store

    def handle_request(self, payload: dict[str, Any]) -> dict[str, Any] | None:
        """Process incoming MCP JSON-RPC message."""
        req_id = payload.get("id")
        method = payload.get("method")
        params = payload.get("params", {})

        if method == "notifications/initialized":
            return None

        if method == "initialize":
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "protocolVersion": "2025-11-25",
                    "serverInfo": {
                        "name": "vscode-office-studio-python",
                        "version": "1.0.0",
                    },
                    "capabilities": {
                        "tools": {"listChanged": False},
                        "resources": {"listChanged": False},
                    },
                },
            }

        if method == "tools/list":
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "tools": [
                        {
                            "name": "list_workspace_files",
                            "description": "List saved Word, Excel, PowerPoint, and code files.",
                            "inputSchema": {"type": "object", "properties": {}},
                        },
                        {
                            "name": "get_document_model",
                            "description": "Get full document JSON model by document ID.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {"id": {"type": "string"}},
                                "required": ["id"],
                            },
                        },
                        {
                            "name": "get_document_outline",
                            "description": "Get structured outline, headings, sheets, or slides.",
                            "inputSchema": {
                                "type": "object",
                                "properties": {"id": {"type": "string"}},
                                "required": ["id"],
                            },
                        },
                    ]
                },
            }

        if method == "tools/call":
            tool_name = params.get("name")
            args = params.get("arguments", {})
            return self._call_tool(req_id, tool_name, args)

        if method == "resources/list":
            records = self.doc_store.list_documents()
            resources = [
                {
                    "uri": f"office://documents/{r['id']}",
                    "name": r["name"],
                    "mimeType": "application/json",
                }
                for r in records
            ]
            return {"jsonrpc": "2.0", "id": req_id, "result": {"resources": resources}}

        if method == "resources/read":
            uri = params.get("uri", "")
            prefix = "office://documents/"
            if uri.startswith(prefix):
                doc_id = uri[len(prefix):]
                try:
                    record = self.doc_store.get_document(doc_id)
                    import json
                    return {
                        "jsonrpc": "2.0",
                        "id": req_id,
                        "result": {
                            "contents": [
                                {
                                    "uri": uri,
                                    "mimeType": "application/json",
                                    "text": json.dumps(record.get("document", {}), indent=2),
                                }
                            ]
                        },
                    }
                except Exception:
                    pass
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32602, "message": "Resource not found."},
            }

        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "error": {"code": -32601, "message": f"Method '{method}' not found."},
        }

    def _call_tool(self, req_id: Any, name: str, args: dict[str, Any]) -> dict[str, Any]:
        import json
        if name == "list_workspace_files":
            records = self.doc_store.list_documents()
            text = json.dumps([{"id": r["id"], "name": r["name"], "updatedAt": r.get("updatedAt")} for r in records], indent=2)
            return {"jsonrpc": "2.0", "id": req_id, "result": {"content": [{"type": "text", "text": text}]}}

        if name in ("get_document_model", "get_document_outline"):
            doc_id = args.get("id")
            if not doc_id:
                return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32602, "message": "Argument 'id' is required."}}
            try:
                rec = self.doc_store.get_document(doc_id)
                doc = rec.get("document", {})
                if name == "get_document_outline":
                    doc_type = doc.get("type", "")
                    data = doc.get("data", {})
                    if doc_type == "word":
                        outline = {"headings": data.get("headings", []), "sections": data.get("sections", [])}
                    elif doc_type == "excel":
                        outline = {"sheets": [s.get("name") for s in data.get("sheets", [])]}
                    elif doc_type == "powerpoint":
                        outline = {"slides": [{"number": s.get("slideNumber"), "title": s.get("title")} for s in data.get("slides", [])]}
                    else:
                        outline = {"language": data.get("language")}
                    text = json.dumps(outline, indent=2)
                else:
                    text = json.dumps(doc, indent=2)
                return {"jsonrpc": "2.0", "id": req_id, "result": {"content": [{"type": "text", "text": text}]}}
            except Exception as e:
                return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32603, "message": str(e)}}

        return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32601, "message": f"Tool '{name}' not found."}}

