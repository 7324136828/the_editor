"""Job and conversion pipeline schemas."""

from __future__ import annotations

from typing import Any, Literal
from pydantic import BaseModel, Field

JobStatus = Literal["queued", "in_progress", "completed", "failed", "discarded"]


class JobRecord(BaseModel):
    id: str
    filename: str
    file_size: int = 0
    status: JobStatus = "queued"
    progress: int = 0
    created_at: str
    completed_at: str | None = None
    error_message: str | None = None
    temp_dir: str | None = None
    zip_path: str | None = None
    summary: dict[str, Any] = Field(default_factory=dict)


class JobResponse(BaseModel):
    job_id: str
    status: JobStatus
    detail: str | None = None

