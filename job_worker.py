"""
job_worker.py
=============
Background worker pool for RPVE async job processing.

Architecture:
  - A Python stdlib queue.Queue holds queued job_ids.
  - N daemon threads (default: 3) each pull from the queue and call run_job().
  - run_job() is the synchronous orchestrator — safe to call from a thread.
  - Worker threads are started once at FastAPI lifespan startup.
  - No Redis, no Celery — stdlib only.

Thread safety:
  - Each thread has its own SQLite connection (via job_store functions).
  - Each job writes to its own isolated job_dir — no shared file paths.
  - The in-memory queue is thread-safe by Python's stdlib design.
"""

import logging
import queue
import threading
from pathlib import Path
from typing import Optional

import job_store

# Number of concurrent processing threads.
# Each thread handles one PDF at a time (OCR + LLM calls are I/O-heavy,
# so 3 threads is a safe default that won't overload the OpenAI rate limit).
WORKER_COUNT = 3

# The in-memory FIFO queue of job_ids awaiting processing.
_job_queue: queue.Queue[str] = queue.Queue()
_workers_started = False
_workers_lock = threading.Lock()

# Root directory for all job workspaces (resolved at import time)
JOBS_DIR: Path = Path(__file__).parent / "jobs"
JOBS_DIR.mkdir(exist_ok=True)


def get_job_dir(job_id: str) -> Path:
    """Return (and create) the workspace directory for a given job."""
    job_dir = JOBS_DIR / job_id
    (job_dir / "input").mkdir(parents=True, exist_ok=True)
    (job_dir / "work").mkdir(parents=True, exist_ok=True)
    (job_dir / "output").mkdir(parents=True, exist_ok=True)
    (job_dir / "logs").mkdir(parents=True, exist_ok=True)
    return job_dir


def _make_job_logger(job_id: str, job_dir: Path) -> logging.Logger:
    """Create a per-job logger that writes to a file AND the console."""
    log_path = job_dir / "logs" / "job.log"
    logger = logging.getLogger(f"rpve.job.{job_id}")
    logger.setLevel(logging.DEBUG)
    
    # Avoid adding duplicate handlers if the logger already exists
    if not logger.handlers:
        formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        
        # 1. File Handler (saves to job directory)
        fh = logging.FileHandler(str(log_path), encoding="utf-8")
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        
        # 2. Console Handler (prints to terminal)
        sh = logging.StreamHandler()
        sh.setFormatter(formatter)
        logger.addHandler(sh)
        
    return logger


def enqueue_job(job_id: str) -> None:
    """Add a job_id to the processing queue."""
    _job_queue.put(job_id)


def _worker_loop() -> None:
    """Main loop run by each worker thread."""
    while True:
        job_id = _job_queue.get()  # blocks until a job is available
        try:
            _execute_job(job_id)
        except Exception as exc:
            # Last-resort catch — _execute_job has its own try/finally,
            # but just in case something really unexpected happens.
            logging.getLogger("rpve.worker").exception(
                "Unhandled exception in worker for job %s: %s", job_id, exc
            )
            try:
                job_store.update_status(job_id, "failed", error=str(exc))
            except Exception:
                pass
        finally:
            _job_queue.task_done()


def _execute_job(job_id: str) -> None:
    """
    Orchestrate a single RPVE job end-to-end.

    This function runs inside a worker thread. It imports flow_orchestrator
    lazily to avoid circular-import issues at module load time.
    """
    job_dir = get_job_dir(job_id)
    logger = _make_job_logger(job_id, job_dir)
    logger.info("=== Job %s started ===", job_id)

    try:
        # Import lazily to avoid circular import at startup
        import flow_orchestrator  # noqa: PLC0415

        # Retrieve job info from DB to get the input file paths
        meta = job_store.get_job(job_id)
        if meta is None:
            raise ValueError(f"Job {job_id} not found in database")

        # Input files are stored in jobs/{job_id}/input/
        input_dir = job_dir / "input"
        pdf_files  = list(input_dir.glob("*.pdf"))
        xlsx_files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xls"))

        if not pdf_files:
            raise ValueError("No PDF file found in job input directory")
        if not xlsx_files:
            raise ValueError("No Excel template found in job input directory")

        pdf_path       = pdf_files[0]
        template_path  = str(xlsx_files[0])
        ref_census     = str(xlsx_files[1]) if len(xlsx_files) > 1 else None

        logger.info(
            "Inputs — PDF: %s | Template: %s | Ref: %s",
            pdf_path.name,
            Path(template_path).name,
            Path(ref_census).name if ref_census else "None",
        )

        # Run the synchronous pipeline.
        # flow_orchestrator.run_job() writes status transitions via the callback.
        def _status(phase: str, ttype: Optional[str] = None) -> None:
            job_store.update_status(job_id, phase, template_type=ttype)
            logger.info("Status → %s", phase)

        result = flow_orchestrator.run_job(
            job_id=job_id,
            pdf_path=pdf_path,
            template_path=template_path,
            ref_census_path=ref_census,
            job_dir=job_dir,
            status_callback=_status,
            logger=logger,
        )

        # Persist the final result
        final_path = result.get("final_report_path") or result.get("excel_path", "")
        job_store.set_final_output(job_id, final_path, result)
        job_store.update_status(job_id, "completed")
        logger.info("=== Job %s completed successfully ===", job_id)

    except Exception as exc:
        import traceback
        logger.error("Job %s FAILED: %s\n%s", job_id, exc, traceback.format_exc())
        job_store.update_status(job_id, "failed", error=str(exc))


def start_workers(n: int = WORKER_COUNT) -> None:
    """
    Start N background worker daemon threads.
    Safe to call multiple times — threads are only started once.
    """
    global _workers_started
    with _workers_lock:
        if _workers_started:
            return
        for i in range(n):
            t = threading.Thread(
                target=_worker_loop,
                name=f"rpve-worker-{i}",
                daemon=True,
            )
            t.start()
            logging.getLogger("rpve.worker").info("Started worker thread: %s", t.name)
        _workers_started = True
