"""
job_models.py
=============
Dataclass definitions for RPVE async job tracking.

Each uploaded invoice creates one JobMeta record.
The job_id is purely backend-internal — it is never
exposed to the frontend.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

VALID_STATES = [
    "queued",
    "preprocessing",
    "classifying",
    "phase1",
    "phase2",
    "phase3",
    "phase4",
    "completed",
    "failed",
]


@dataclass
class JobMeta:
    """Lifecycle metadata for a single RPVE processing job."""

    job_id: str
    status: str = "queued"

    # Template type detected during classification (type1 / type2 / type3)
    template_type: Optional[str] = None

    # Timestamps (ISO-8601 UTC strings)
    created_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())

    # Error message — only populated on failure
    error: Optional[str] = None

    # Per-phase output file paths (absolute)
    phase1_output: Optional[str] = None   # raw extraction XLSX
    phase2_output: Optional[str] = None   # filled census XLSX
    phase3_output: Optional[str] = None   # validated census XLSX
    phase4_output: Optional[str] = None   # LLM-resolved census XLSX
    final_output: Optional[str] = None    # primary download target (highest available phase)

    # Rich JSON result dict from process_invoice_data() — serialised as text
    result_json: Optional[str] = None

    def is_terminal(self) -> bool:
        return self.status in ("completed", "failed")
