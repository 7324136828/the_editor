"""Job state persistence store using SQLite."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generator

from ..schemas.job import JobRecord, JobStatus


class JobStore:
    """Persistent SQLite store for background and conversion jobs."""

    def __init__(self, db_path: Path | str) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _connection(self) -> Generator[sqlite3.Connection, None, None]:
        conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def _init_db(self) -> None:
        with self._connection() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    id TEXT PRIMARY KEY,
                    filename TEXT NOT NULL,
                    file_size INTEGER DEFAULT 0,
                    status TEXT NOT NULL,
                    progress INTEGER DEFAULT 0,
                    created_at TEXT NOT NULL,
                    completed_at TEXT,
                    error_message TEXT,
                    temp_dir TEXT,
                    zip_path TEXT,
                    summary_json TEXT
                )
                """
            )
            conn.commit()

    def create_job(self, job: JobRecord) -> JobRecord:
        with self._connection() as conn:
            conn.execute(
                """
                INSERT INTO jobs (
                    id, filename, file_size, status, progress,
                    created_at, completed_at, error_message, temp_dir, zip_path, summary_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    job.id,
                    job.filename,
                    job.file_size,
                    job.status,
                    job.progress,
                    job.created_at,
                    job.completed_at,
                    job.error_message,
                    job.temp_dir,
                    job.zip_path,
                    json.dumps(job.summary),
                ),
            )
            conn.commit()
        return job

    def update_job(self, job: JobRecord) -> None:
        with self._connection() as conn:
            conn.execute(
                """
                UPDATE jobs SET
                    status = ?,
                    progress = ?,
                    completed_at = ?,
                    error_message = ?,
                    temp_dir = ?,
                    zip_path = ?,
                    summary_json = ?
                WHERE id = ?
                """,
                (
                    job.status,
                    job.progress,
                    job.completed_at,
                    job.error_message,
                    job.temp_dir,
                    job.zip_path,
                    json.dumps(job.summary),
                    job.id,
                ),
            )
            conn.commit()

    def get_job(self, job_id: str) -> JobRecord | None:
        with self._connection() as conn:
            cursor = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,))
            row = cursor.fetchone()
            if not row:
                return None
            return self._row_to_job(row)

    def list_jobs(self) -> list[JobRecord]:
        with self._connection() as conn:
            cursor = conn.execute("SELECT * FROM jobs ORDER BY created_at DESC")
            return [self._row_to_job(r) for r in cursor.fetchall()]

    def _row_to_job(self, row: sqlite3.Row) -> JobRecord:
        summary = {}
        if row["summary_json"]:
            try:
                summary = json.loads(row["summary_json"])
            except Exception:
                pass
        return JobRecord(
            id=row["id"],
            filename=row["filename"],
            file_size=row["file_size"],
            status=row["status"],
            progress=row["progress"],
            created_at=row["created_at"],
            completed_at=row["completed_at"],
            error_message=row["error_message"],
            temp_dir=row["temp_dir"],
            zip_path=row["zip_path"],
            summary=summary,
        )
