"""Strict column allowlists for OASIS sources, applied BEFORE encryption/upload.

This is data minimization, not anonymization. Educator feedback keeps the two
comment-question responses in Answer text; those may include identifying text.
Student assessments retain matching/completion metadata, never scores/comments.
No unknown/new column is kept automatically. Update the allowlist deliberately
if a future feature actually needs a new source field.
"""
from __future__ import annotations
import csv
import io

PRIVACY_VERSION = 1
EDUCATOR_COLUMNS = (
    "Course ID", "Start Date", "End Date", "Evaluator", "Evaluator Username",
    "Evaluator External ID", "Evaluator Email", "Evaluation", "Form Record",
    "Question ID", "Question", "Answer text", "Multiple Choice Value",
    "Multiple Choice Label", "Submit Date",
)
STUDENT_COLUMNS = (
    "Course ID", "Start Date", "End Date", "Student", "Student External ID",
    "Evaluator", "Evaluator Username", "Evaluator Email", "Evaluation",
    "Form Record", "Submit Date",
)
# Keep existing source-validation semantics: optional missing identities are
# still surfaced by report checks, not invented when a CSV is archived.
STUDENT_REQUIRED = (
    "Course ID", "Start Date", "End Date", "Evaluator", "Evaluation",
    "Form Record", "Student", "Submit Date",
)


def minimize_oasis_csv(raw: bytes, kind: str) -> dict:
    """Select only allowed columns; preserve retained cell values and row order.

    Reads the whole source with size/quoting/header/row-width checks. Re-encodes
    as UTF-8 BOM with deterministic CSV quoting. Missing optional columns are
    NOT fabricated. Changing a removed field cannot create a new source snapshot.
    """
    from schedule_app.services.oasis_evaluations import (
        inspect_oasis_csv, OASISArchiveError, OASIS_REQUIRED_COLUMNS,
        OASIS_MAX_FIELD_CHARS, _CSV_LOCK,
    )
    if kind not in ("educator", "student"):
        raise OASISArchiveError("Unknown evaluation upload type; nothing was saved.")
    allowed = STUDENT_COLUMNS if kind == "student" else EDUCATOR_COLUMNS
    required = STUDENT_REQUIRED if kind == "student" else OASIS_REQUIRED_COLUMNS
    original = inspect_oasis_csv(raw, required_columns=required)
    codec = {"UTF-8": "utf-8-sig", "UTF-16": "utf-16", "Windows-1252": "cp1252"}[original["encoding"]]
    output = io.StringIO(newline="")
    with _CSV_LOCK:
        before = csv.field_size_limit()
        csv.field_size_limit(OASIS_MAX_FIELD_CHARS)
        try:
            reader = csv.reader(io.StringIO(raw.decode(codec), newline=""), strict=True)
            headers = [x.strip().lstrip("\ufeff") for x in next(reader)]
            kept = [x for x in allowed if x in headers]
            indexes = [headers.index(x) for x in kept]
            writer = csv.writer(output, lineterminator="\r\n")
            writer.writerow(kept)
            for row in reader:
                if not row or all(not x.strip() for x in row):
                    continue
                writer.writerow([row[i] for i in indexes])
        except (csv.Error, UnicodeError, StopIteration):
            raise OASISArchiveError("The evaluation CSV could not be minimized safely. Nothing was saved.") from None
        finally:
            csv.field_size_limit(before)
    safe = output.getvalue().encode("utf-8-sig")
    # This second check also protects against malformed minimized output.
    details = inspect_oasis_csv(safe, required_columns=required)
    removed = [x for x in headers if x not in allowed]
    return {"raw": safe, "details": details, "privacy": {
        "version": PRIVACY_VERSION, "kind": kind,
        "source_column_count": len(headers), "retained_column_count": len(kept),
        "removed_column_count": len(removed), "retained_columns": kept,
        "removed_columns": removed, "needs_minimization": raw != safe,
    }}
