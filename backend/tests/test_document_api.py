"""Unit and integration tests for Code Office Studio backend."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from fastapi.testclient import TestClient

from backend.app.main import create_app


class DocumentApiTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.storage_dir = Path(self.temp_dir.name) / "storage"
        self.db_path = Path(self.temp_dir.name) / "test_office.db"
        self.app = create_app(storage_dir=self.storage_dir, db_path=self.db_path)
        self.client = TestClient(self.app)

    def tearDown(self) -> None:
        import gc
        gc.collect()
        try:
            self.temp_dir.cleanup()
        except Exception:
            pass

    def test_health(self) -> None:
        res = self.client.get("/api/health")
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["status"], "ok")
        self.assertEqual(data["storage"], "local-disk")

    def test_word_document_crud_and_conflict(self) -> None:
        word_doc = {
            "type": "word",
            "data": {
                "id": "doc-1",
                "title": "Quarterly Report",
                "author": "Analyst",
                "createdAt": "2026-09-20T00:00:00Z",
                "modifiedAt": "2026-09-20T00:00:00Z",
                "headings": [{"id": "h1", "text": "Overview", "level": 1, "pageIndex": 0}],
                "sections": [{"id": "s1", "title": "Main", "orientation": "portrait", "pageNumber": 1}],
                "comments": [],
                "typography": {
                    "fontFamily": "Segoe UI",
                    "textColor": "#000000",
                    "fontSize": 12,
                    "lineHeight": 1.5,
                    "marginSize": "normal",
                },
                "stats": {
                    "words": 150,
                    "characters": 900,
                    "charactersNoSpaces": 750,
                    "paragraphs": 5,
                    "pages": 1,
                    "readingTimeMinutes": 1,
                    "readabilityScore": "Standard",
                },
                "paragraphs": [
                    {"id": "p1", "type": "heading-1", "text": "Overview"},
                    {"id": "p2", "type": "body", "text": "This is the executive summary."},
                ],
            },
        }

        # 1. Save new document
        res = self.client.put("/api/documents/doc-1", json={"name": "Report.docx", "document": word_doc, "revision": None})
        self.assertEqual(res.status_code, 201)
        saved = res.json()
        self.assertEqual(saved["revision"], 1)

        # 2. Get document
        res = self.client.get("/api/documents/doc-1")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["name"], "Report.docx")

        # 3. Update document with correct revision
        word_doc["data"]["title"] = "Updated Report"
        res = self.client.put("/api/documents/doc-1", json={"name": "Report.docx", "document": word_doc, "revision": 1})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["revision"], 2)

        # 4. Conflict on stale revision
        res = self.client.put("/api/documents/doc-1", json={"name": "Report.docx", "document": word_doc, "revision": 1})
        self.assertEqual(res.status_code, 409)

        # 5. Check history
        res = self.client.get("/api/documents/doc-1/history")
        self.assertEqual(res.status_code, 200)
        history = res.json()["versions"]
        self.assertEqual(len(history), 2)
        self.assertEqual(history[0]["revision"], 2)

        # 6. Delete document
        res = self.client.delete("/api/documents/doc-1")
        self.assertEqual(res.status_code, 200)
        self.assertTrue(res.json()["deleted"])

    def test_excel_document_crud(self) -> None:
        excel_doc = {
            "type": "excel",
            "data": {
                "id": "sheet-1",
                "title": "Financial Model",
                "activeSheetId": "s1",
                "sheets": [
                    {
                        "id": "s1",
                        "name": "Revenue",
                        "rowCount": 20,
                        "colCount": 5,
                        "columns": [
                            {"key": "A", "label": "Item", "type": "text", "width": 120},
                            {"key": "B", "label": "Q1", "type": "currency", "width": 100},
                            {"key": "C", "label": "Q2", "type": "currency", "width": 100},
                            {"key": "D", "label": "Q3", "type": "currency", "width": 100},
                            {"key": "E", "label": "Total", "type": "currency", "width": 110},
                        ],
                        "cells": {
                            "A1": {"value": "Software Sales"},
                            "B1": {"value": 50000},
                            "C1": {"value": 60000},
                            "D1": {"value": 75000},
                            "E1": {"value": 185000, "formula": "=SUM(B1:D1)"},
                        },
                    }
                ],
            },
        }

        res = self.client.put("/api/documents/sheet-1", json={"name": "Model.xlsx", "document": excel_doc, "revision": None})
        self.assertEqual(res.status_code, 201)

        res = self.client.get("/api/documents")
        self.assertEqual(res.status_code, 200)
        docs = res.json()["documents"]
        self.assertEqual(len(docs), 1)
        self.assertEqual(docs[0]["id"], "sheet-1")

    def test_chat_assistant_offline_engine(self) -> None:
        # Test chat with Word document context
        word_doc = {
            "type": "word",
            "data": {
                "id": "w1",
                "title": "Business Plan",
                "headings": [{"text": "Executive Summary", "level": 1}],
                "stats": {"words": 500, "paragraphs": 6, "readingTimeMinutes": 2, "readabilityScore": "Standard"},
                "paragraphs": [{"id": "p1", "type": "body", "text": "Antigravity helps developers innovate."}],
            },
        }
        res = self.client.post("/api/chat", json={"query": "Summarize this document", "document": word_doc})
        self.assertEqual(res.status_code, 200)
        reply = res.json()
        self.assertIn("Business Plan", reply["message"])
        self.assertIn("Structure", reply["message"])
        self.assertTrue(len(reply["suggestions"]) > 0)

        # Test chat with Excel document context
        excel_doc = {
            "type": "excel",
            "data": {
                "id": "x1",
                "title": "Budget",
                "activeSheetId": "s1",
                "sheets": [
                    {
                        "id": "s1",
                        "name": "Expenses",
                        "cells": {
                            "A1": {"value": "Rent"},
                            "B1": {"value": 2000},
                            "A2": {"value": "Utilities"},
                            "B2": {"value": 300},
                            "B3": {"value": 2300, "formula": "=SUM(B1:B2)"},
                        },
                    }
                ],
            },
        }
        res = self.client.post("/api/chat", json={"query": "Explain formulas in this sheet", "document": excel_doc})
        self.assertEqual(res.status_code, 200)
        reply = res.json()
        self.assertIn("=SUM(B1:B2)", reply["message"])

    def test_conversion_and_jobs_lifecycle(self) -> None:
        sample_csv = b"Product,Price,Quantity\nLaptop,1200,5\nMouse,25,20\nKeyboard,75,15\n"
        res = self.client.post(
            "/api/convert",
            files={"file": ("inventory.csv", sample_csv, "text/csv")},
        )
        self.assertEqual(res.status_code, 200)
        job_info = res.json()
        job_id = job_info["job_id"]

        # Check job status
        res = self.client.get(f"/api/jobs/{job_id}")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["filename"], "inventory.csv")

        # Discard job
        res = self.client.post(f"/api/jobs/{job_id}/discard")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["status"], "discarded")

        # Verify job is marked discarded
        res = self.client.get(f"/api/jobs/{job_id}")
        self.assertEqual(res.json()["status"], "discarded")

    def test_mcp_endpoint(self) -> None:
        # Initialize
        res = self.client.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "initialize"})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.json()["result"]["serverInfo"]["name"], "vscode-office-studio-python")

        # Tools list
        res = self.client.post("/mcp", json={"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        self.assertEqual(res.status_code, 200)
        tools = res.json()["result"]["tools"]
        tool_names = [t["name"] for t in tools]
        self.assertIn("list_workspace_files", tool_names)
        self.assertIn("get_document_model", tool_names)


if __name__ == "__main__":
    unittest.main()

