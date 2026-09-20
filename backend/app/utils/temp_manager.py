"""System temporary directory manager for isolated conversion jobs.

Complies with productionization Pillar 3:
- Isolated folders under system temp directory (tempfile.gettempdir() / 'office_prod_jobs' / <job_id>)
- Standard job structure: inputs/, work/, outputs/, archive/
- Packaging outputs into ZIP archive
- Robust purge and lifecycle management
"""

from __future__ import annotations

import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path


class TempJobManager:
    """Manages isolated temp folder lifecycles for conversion jobs."""

    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = base_dir or (Path(tempfile.gettempdir()) / "office_prod_jobs")
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def create_job_dir(self, job_id: str | None = None) -> tuple[str, Path]:
        """Create an isolated folder hierarchy for a job."""
        jid = job_id or str(uuid.uuid4())
        job_dir = self.base_dir / jid
        (job_dir / "inputs").mkdir(parents=True, exist_ok=True)
        (job_dir / "work").mkdir(parents=True, exist_ok=True)
        (job_dir / "outputs").mkdir(parents=True, exist_ok=True)
        (job_dir / "archive").mkdir(parents=True, exist_ok=True)
        return jid, job_dir

    def stage_file(self, job_dir: Path, filename: str, content: bytes) -> Path:
        """Stage an incoming uploaded or pasted file into inputs/."""
        safe_name = Path(filename).name or "document.bin"
        target_path = job_dir / "inputs" / safe_name
        target_path.write_bytes(content)
        return target_path

    def package_zip(self, job_dir: Path, zip_name: str = "conversion_output.zip") -> Path:
        """Package all files in outputs/ into archive/<zip_name>."""
        outputs_dir = job_dir / "outputs"
        archive_dir = job_dir / "archive"
        archive_dir.mkdir(parents=True, exist_ok=True)
        zip_path = archive_dir / zip_name

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in outputs_dir.rglob("*"):
                if file.is_file():
                    zf.write(file, arcname=file.relative_to(outputs_dir))
        return zip_path

    def cleanup_job(self, job_dir: Path | str) -> None:
        """Purge isolated system temp folder."""
        path = Path(job_dir)
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)

