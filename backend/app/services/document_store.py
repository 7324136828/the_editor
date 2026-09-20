"""Document store service with atomic writes and version tracking."""

from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from ..schemas.document import (
    ID_PATTERN,
    DocumentEnvelope,
    DocumentRecord,
    ValidationError,
    validate_document_model,
)

MAX_VERSIONS = 50


class DocumentConflictError(Exception):
    def __init__(self, message: str, current: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.status_code = 409
        self.current = current


class DocumentNotFoundError(Exception):
    def __init__(self, message: str = "Document not found.") -> None:
        super().__init__(message)
        self.status_code = 404


class DocumentStore:
    """Atomic local disk storage for Code Office Studio documents."""

    def __init__(self, storage_dir: Path | str | None = None) -> None:
        raw_dir = storage_dir or os.environ.get("OFFICE_DATA_DIR", ".office-data")
        self.root = Path(raw_dir).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self._locks: dict[str, threading.Lock] = {}
        self._global_lock = threading.Lock()

    def _get_lock(self, doc_id: str) -> threading.Lock:
        with self._global_lock:
            if doc_id not in self._locks:
                self._locks[doc_id] = threading.Lock()
            return self._locks[doc_id]

    def _file_path(self, doc_id: str) -> Path:
        if not ID_PATTERN.match(doc_id):
            raise ValidationError("Invalid document ID. Use letters, numbers, underscores, and hyphens.")
        return self.root / f"{doc_id}.json"

    def read_envelope(self, doc_id: str) -> dict[str, Any] | None:
        target = self._file_path(doc_id)
        if not target.is_file():
            return None
        try:
            return json.loads(target.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None

    def list_documents(self) -> list[dict[str, Any]]:
        self.root.mkdir(parents=True, exist_ok=True)
        records: list[dict[str, Any]] = []
        for file in self.root.glob("*.json"):
            doc_id = file.stem
            if not ID_PATTERN.match(doc_id):
                continue
            envelope = self.read_envelope(doc_id)
            if envelope and "head" in envelope:
                records.append(envelope["head"])
        records.sort(key=lambda r: r.get("updatedAt", ""), reverse=True)
        return records

    def get_document(self, doc_id: str) -> dict[str, Any]:
        envelope = self.read_envelope(doc_id)
        if not envelope or "head" not in envelope:
            raise DocumentNotFoundError()
        return envelope["head"]

    def get_history(self, doc_id: str) -> list[dict[str, Any]]:
        envelope = self.read_envelope(doc_id)
        if not envelope:
            raise DocumentNotFoundError()
        versions = envelope.get("versions", [])
        return [
            {
                "revision": v["revision"],
                "name": v["name"],
                "updatedAt": v["updatedAt"],
            }
            for v in reversed(versions)
        ]

    def get_version(self, doc_id: str, revision: int) -> dict[str, Any]:
        envelope = self.read_envelope(doc_id)
        if not envelope:
            raise DocumentNotFoundError()
        for v in envelope.get("versions", []):
            if v.get("revision") == revision:
                return v
        raise DocumentNotFoundError("Version is no longer available.")

    def save_document(
        self,
        doc_id: str,
        name: str,
        document: dict[str, Any],
        revision: int | None,
    ) -> tuple[dict[str, Any], bool]:
        """Validate, check conflict, and atomically persist document envelope."""
        doc_type = document.get("type", "")
        doc_data = document.get("data", {})
        validate_document_model(doc_type, doc_data, doc_id)

        clean_name = name.strip()
        if not clean_name or len(clean_name) > 240 or any(c in clean_name for c in ("\x00", "\x1f", "\\", "/")):
            raise ValidationError("A filename of 1–240 characters, without path separators, is required.")

        lock = self._get_lock(doc_id)
        with lock:
            existing = self.read_envelope(doc_id)
            existing_rev = existing["head"]["revision"] if existing and "head" in existing else None

            if existing_rev != revision:
                raise DocumentConflictError(
                    "This document was changed elsewhere. Reload the saved version or save a copy.",
                    current=existing["head"] if existing else None,
                )

            new_rev = (existing_rev or 0) + 1
            now_iso = datetime.now(timezone.utc).isoformat()
            record = {
                "id": doc_id,
                "name": clean_name,
                "document": document,
                "revision": new_rev,
                "updatedAt": now_iso,
            }

            versions = list(existing.get("versions", [])) if existing else []
            versions.append(record)
            if len(versions) > MAX_VERSIONS:
                versions = versions[-MAX_VERSIONS:]

            envelope = {
                "schemaVersion": 1,
                "head": record,
                "versions": versions,
            }

            # Atomic file commit
            temp_path = self.root / f"{doc_id}.{uuid.uuid4().hex}.tmp"
            try:
                temp_path.write_text(json.dumps(envelope, ensure_ascii=False, indent=2), encoding="utf-8")
                target_path = self._file_path(doc_id)
                os.replace(temp_path, target_path)
            finally:
                if temp_path.exists():
                    temp_path.unlink(missing_ok=True)

            return record, existing is None

    def delete_document(self, doc_id: str) -> bool:
        lock = self._get_lock(doc_id)
        with lock:
            target = self._file_path(doc_id)
            if target.exists():
                target.unlink(missing_ok=True)
                return True
            return False

