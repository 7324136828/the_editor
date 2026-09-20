"""FastAPI application entrypoint for Code Office Studio."""

from __future__ import annotations

import os
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import (
    BackgroundTasks,
    Depends,
    FastAPI,
    File,
    HTTPException,
    Request,
    Response,
    UploadFile,
    status,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from .schemas.chat import ChatRequest, ChatResponse
from .schemas.document import DocumentPayload, ValidationError
from .schemas.job import JobRecord, JobResponse
from .services.chat_service import ChatService
from .services.conversion_pipeline import ConversionPipeline
from .services.document_store import (
    DocumentConflictError,
    DocumentNotFoundError,
    DocumentStore,
)
from .services.job_store import JobStore
from .services.mcp_service import McpService
from .utils.temp_manager import TempJobManager

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def configured_cors_origins() -> list[str]:
    raw = os.environ.get(
        "CORS_ORIGINS",
        "http://localhost:5173,http://127.0.0.1:5173,http://localhost:5174,http://127.0.0.1:5174,http://localhost:4173,http://127.0.0.1:4173",
    )
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


def create_app(
    storage_dir: Path | str | None = None,
    db_path: Path | str | None = None,
) -> FastAPI:
    doc_store = DocumentStore(storage_dir)
    job_store = JobStore(db_path or (PROJECT_ROOT / "data" / "office.db"))
    temp_manager = TempJobManager()
    pipeline = ConversionPipeline(temp_manager)
    chat_service = ChatService()
    mcp_service = McpService(doc_store)

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        app.state.doc_store = doc_store
        app.state.job_store = job_store
        app.state.temp_manager = temp_manager
        app.state.pipeline = pipeline
        app.state.chat_service = chat_service
        app.state.mcp_service = mcp_service
        yield

    app = FastAPI(
        title="Code Office Studio API",
        version="1.0.0",
        description="Production API for Word, Excel, PowerPoint document viewer and AI chat",
        lifespan=lifespan,
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=configured_cors_origins(),
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # ------------------ Health & Storage ------------------
    @app.get("/api/health")
    def health() -> dict[str, Any]:
        return {
            "status": "ok",
            "storage": "local-disk",
            "maxVersions": 50,
        }

    # ------------------ Document CRUD ------------------
    @app.get("/api/documents")
    def list_documents() -> dict[str, Any]:
        return {"documents": doc_store.list_documents()}

    @app.get("/api/documents/{doc_id}")
    def get_document(doc_id: str) -> dict[str, Any]:
        try:
            return doc_store.get_document(doc_id)
        except DocumentNotFoundError:
            raise HTTPException(status_code=404, detail="Document not found.")
        except ValidationError as e:
            raise HTTPException(status_code=e.status_code, detail=str(e))

    @app.put("/api/documents/{doc_id}")
    def save_document(doc_id: str, payload: DocumentPayload) -> Response:
        try:
            record, created = doc_store.save_document(
                doc_id=doc_id,
                name=payload.name,
                document=payload.document,
                revision=payload.revision,
            )
            return JSONResponse(
                status_code=status.HTTP_201_CREATED if created else status.HTTP_200_OK,
                content=record,
            )
        except DocumentConflictError as e:
            return JSONResponse(
                status_code=status.HTTP_409_CONFLICT,
                content={"error": str(e), "current": e.current},
            )
        except ValidationError as e:
            raise HTTPException(status_code=e.status_code, detail=str(e))
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))

    @app.delete("/api/documents/{doc_id}")
    def delete_document(doc_id: str) -> dict[str, Any]:
        deleted = doc_store.delete_document(doc_id)
        return {"id": doc_id, "deleted": deleted}

    @app.get("/api/documents/{doc_id}/history")
    def get_document_history(doc_id: str) -> dict[str, Any]:
        try:
            return {"versions": doc_store.get_history(doc_id)}
        except DocumentNotFoundError:
            raise HTTPException(status_code=404, detail="Document not found.")

    @app.get("/api/documents/{doc_id}/history/{revision}")
    def get_document_version(doc_id: str, revision: int) -> dict[str, Any]:
        try:
            return doc_store.get_version(doc_id, revision)
        except DocumentNotFoundError:
            raise HTTPException(status_code=404, detail="Version is no longer available.")

    # ------------------ AI Document Chat ------------------
    @app.post("/api/chat", response_model=ChatResponse)
    async def chat(request: ChatRequest) -> ChatResponse:
        return await chat_service.generate_response(
            messages=request.messages,
            query=request.query,
            document=request.document,
        )

    # ------------------ Ingestion & Conversion Pipeline ------------------
    def _run_conversion_task(job_id: str, staged_file: Path, job_dir: Path) -> None:
        job = job_store.get_job(job_id)
        if not job:
            return
        pipeline.process_job(job, staged_file, job_dir)
        job_store.update_job(job)

    @app.post("/api/convert")
    async def convert_document(
        background_tasks: BackgroundTasks,
        file: UploadFile = File(...),
    ) -> JobResponse:
        content = await file.read()
        filename = file.filename or "uploaded_document"

        job_id, job_dir = temp_manager.create_job_dir()
        staged_path = temp_manager.stage_file(job_dir, filename, content)

        job = JobRecord(
            id=job_id,
            filename=filename,
            file_size=len(content),
            status="in_progress",
            progress=10,
            created_at=datetime.now(timezone.utc).isoformat(),
            temp_dir=str(job_dir),
        )
        job_store.create_job(job)

        background_tasks.add_task(_run_conversion_task, job_id, staged_path, job_dir)
        return JobResponse(job_id=job_id, status="in_progress")

    @app.get("/api/jobs")
    def list_jobs() -> list[JobRecord]:
        return job_store.list_jobs()

    @app.get("/api/jobs/{job_id}")
    def get_job(job_id: str) -> JobRecord:
        job = job_store.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        return job

    @app.post("/api/jobs/{job_id}/discard")
    def discard_job(job_id: str) -> dict[str, Any]:
        job = job_store.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")

        # Purge isolated temp folder
        if job.temp_dir:
            temp_manager.cleanup_job(job.temp_dir)

        job.status = "discarded"
        job.progress = 0
        job.error_message = "Job discarded by user."
        job_store.update_job(job)
        return {"job_id": job_id, "status": "discarded", "detail": "Temporary folder purged"}

    @app.get("/api/jobs/{job_id}/download-zip")
    def download_job_zip(job_id: str):
        job = job_store.get_job(job_id)
        if not job or job.status != "completed":
            raise HTTPException(status_code=400, detail="Job is not completed")
        if not job.zip_path or not Path(job.zip_path).is_file():
            raise HTTPException(status_code=404, detail="ZIP archive not found")

        download_name = f"{Path(job.filename).stem}_converted.zip"
        return FileResponse(
            path=job.zip_path,
            filename=download_name,
            media_type="application/zip",
        )

    # ------------------ MCP JSON-RPC ------------------
    @app.post("/mcp")
    async def mcp_endpoint(request: Request) -> Response:
        try:
            body = await request.json()
        except Exception:
            return JSONResponse(
                status_code=400,
                content={"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Invalid JSON."}},
            )
        response_payload = mcp_service.handle_request(body)
        if response_payload is None:
            return Response(status_code=204)
        return JSONResponse(content=response_payload)

    return app


app = create_app()
