from __future__ import annotations

import copy
import json
import os
import sqlite3
from pathlib import Path

from .models import DocumentState, utc_now

SCHEMA = """
CREATE TABLE IF NOT EXISTS recovery (
    doc_id TEXT PRIMARY KEY,
    payload TEXT NOT NULL,
    path TEXT,
    updated TEXT NOT NULL,
    dirty INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS versions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id TEXT NOT NULL,
    label TEXT NOT NULL,
    created TEXT NOT NULL,
    payload TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""

MAX_VERSIONS_LISTED = 50


def default_data_root() -> Path:
    override = os.environ.get("FOLIO_DATA_DIR")
    if override:
        return Path(override)
    try:
        from PySide6.QtCore import QStandardPaths

        location = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.AppLocalDataLocation
        )
        if location:
            return Path(location)
    except Exception:
        pass
    return Path.home() / ".folio"


class StorageError(RuntimeError):
    pass


class LocalStore:
    def __init__(self, root: Path | None = None):
        self.root = Path(root) if root is not None else default_data_root()
        self.root.mkdir(parents=True, exist_ok=True)
        self._db = sqlite3.connect(self.root / "folio.db")
        self._db.execute("PRAGMA journal_mode=WAL")
        self._db.execute("PRAGMA synchronous=NORMAL")
        with self._db:
            self._db.executescript(SCHEMA)

    def save_recovery(self, state: DocumentState, path: str | None, dirty: bool) -> None:
        payload = json.dumps(state.to_dict(), ensure_ascii=False)
        with self._db:
            self._db.execute(
                "INSERT INTO recovery (doc_id, payload, path, updated, dirty) "
                "VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(doc_id) DO UPDATE SET "
                "payload=excluded.payload, path=excluded.path, "
                "updated=excluded.updated, dirty=excluded.dirty",
                (state.id, payload, path, utc_now(), 1 if dirty else 0),
            )

    def recoverable(self) -> list[dict]:
        rows = self._db.execute(
            "SELECT doc_id, payload, path, updated, dirty FROM recovery "
            "WHERE dirty = 1 ORDER BY updated DESC"
        ).fetchall()
        results: list[dict] = []
        for doc_id, payload, path, updated, dirty in rows:
            try:
                data = json.loads(payload)
                title = data.get("state", data).get("title") or data.get("title") or "Untitled"
            except Exception:
                title = "Untitled"
            results.append({
                "doc_id": doc_id,
                "title": title,
                "path": path,
                "updated": updated,
                "dirty": bool(dirty),
                "payload": payload,
            })
        return results

    def load_recovery(self, doc_id: str) -> tuple[DocumentState, str | None] | None:
        row = self._db.execute(
            "SELECT payload, path FROM recovery WHERE doc_id = ?", (doc_id,)
        ).fetchone()
        if not row:
            return None
        state = DocumentState.from_dict(json.loads(row[0]))
        return state, row[1]

    def mark_clean(self, doc_id: str) -> None:
        with self._db:
            self._db.execute(
                "UPDATE recovery SET dirty = 0 WHERE doc_id = ?", (doc_id,)
            )

    def discard_recovery(self, doc_id: str) -> None:
        with self._db:
            self._db.execute("DELETE FROM recovery WHERE doc_id = ?", (doc_id,))

    def add_version(self, state: DocumentState, label: str) -> int:
        snapshot = copy.deepcopy(state)
        payload = json.dumps(snapshot.to_dict(), ensure_ascii=False)
        with self._db:
            cursor = self._db.execute(
                "INSERT INTO versions (doc_id, label, created, payload) VALUES (?, ?, ?, ?)",
                (state.id, label, utc_now(), payload),
            )
        return int(cursor.lastrowid)

    def versions(self, doc_id: str) -> list[dict]:
        rows = self._db.execute(
            "SELECT id, label, created FROM versions WHERE doc_id = ? "
            "ORDER BY id DESC LIMIT ?",
            (doc_id, MAX_VERSIONS_LISTED),
        ).fetchall()
        return [{"id": row[0], "label": row[1], "created": row[2]} for row in rows]

    def load_version(self, version_id: int) -> DocumentState:
        row = self._db.execute(
            "SELECT payload FROM versions WHERE id = ?", (version_id,)
        ).fetchone()
        if not row:
            raise StorageError(f"Version {version_id} not found")
        return DocumentState.from_dict(json.loads(row[0]))

    def get_setting(self, key: str, default=None):
        row = self._db.execute(
            "SELECT value FROM settings WHERE key = ?", (key,)
        ).fetchone()
        if not row:
            return default
        try:
            return json.loads(row[0])
        except Exception:
            return default

    def set_setting(self, key: str, value) -> None:
        with self._db:
            self._db.execute(
                "INSERT INTO settings (key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, json.dumps(value)),
            )

    def close(self) -> None:
        try:
            self._db.close()
        except Exception:
            pass
