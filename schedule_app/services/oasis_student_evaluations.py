"""Student-assessment metadata archives, isolated from educator feedback.

Only fields needed for identity matching, completion counts and source/date checks
are saved. Grades, questions and student assessment comments are not persisted.
"""
from __future__ import annotations

import base64
from collections import Counter
import csv
import hashlib
import hmac
import io
import re
from typing import Any

from schedule_app.services.oasis_evaluations import (
    GitHubOASISEvaluations, OASISArchiveError, OASIS_MAX_FIELD_CHARS,
    _CSV_LOCK, inspect_oasis_csv, oasis_export_label,
)

STUDENT_ARCHIVE_VERSION = 2
STUDENT_SUBFOLDER = "oasis_student_evaluations"
STUDENT_REQUIRED_COLUMNS = ("Student", "Submit Date")
# Add a new title here only after confirming it assesses students, not educators.
STUDENT_FORM_TITLES = (
    "*Clinical Assessment of Student",
    "*PEDS Handoff",
    "*PEDS History Taking & Physical Exam",
)
_STUDENT_ID_CONTEXT = b"schedule-app:oasis-student-assessment-export-id:v1"


def _form_key(value: str) -> str:
    """Normalize label formatting only; never infer a form's educational role."""
    return re.sub(r"\s+", " ", value.strip().lstrip("*").strip()).casefold()


_STUDENT_FORMS = {_form_key(title): title for title in STUDENT_FORM_TITLES}


def _form_metadata(raw: bytes, *, student_archive: bool) -> dict[str, Any]:
    """Validate the whole CSV, then inspect form labels without retaining people.

    The general inspector bounds bytes, rows and field length and checks quoting
    and row widths. Do not weaken these checks for an assessment export.
    """
    from schedule_app.services.oasis_privacy import STUDENT_REQUIRED
    details = inspect_oasis_csv(raw, required_columns=STUDENT_REQUIRED if student_archive else None)
    codec = {"UTF-8": "utf-8-sig", "UTF-16": "utf-16",
             "Windows-1252": "cp1252"}[details["encoding"]]
    forms: set[tuple[str, str, str]] = set()
    type_rows: Counter[str] = Counter()
    type_forms: dict[str, set[tuple[str, str, str]]] = {}
    missing_ids = blank_submissions = 0
    # The CSV field limit is process-wide, so reuse the base inspector's lock.
    with _CSV_LOCK:
        old_limit = csv.field_size_limit()
        csv.field_size_limit(OASIS_MAX_FIELD_CHARS)
        try:
            reader = csv.reader(io.StringIO(raw.decode(codec), newline=""), strict=True)
            headers = [cell.strip().lstrip("\ufeff") for cell in next(reader)]
            if student_archive:
                missing = [name for name in STUDENT_REQUIRED_COLUMNS if name not in headers]
                if missing:
                    raise OASISArchiveError("This is not the expected student-evaluation export. "
                                           "Missing columns: " + ", ".join(missing) + ".")
            form_index = headers.index("Evaluation")
            course_index = headers.index("Course ID")
            record_index = headers.index("Form Record")
            submitted_index = headers.index("Submit Date") if "Submit Date" in headers else None
            record_number = 1
            for values in reader:
                if not values or all(not value.strip() for value in values):
                    continue
                record_number += 1
                key = _form_key(values[form_index])
                title = _STUDENT_FORMS.get(key)
                if student_archive and title is None:
                    # Neither question text, comments, names nor untrusted labels
                    # appear in the error. Stop before contacting GitHub.
                    raise OASISArchiveError(
                        f"CSV record {record_number} has a form type not supported by the student-evaluation archive. "
                        "Use an export containing Clinical Assessment of Student, PEDS Handoff, and/or "
                        "PEDS History Taking & Physical Exam. Put feedback about educators under "
                        "Evaluations of educators. No rows have been saved or dropped."
                    )
                if not student_archive and title is not None:
                    raise OASISArchiveError(
                        "This CSV contains evaluations OF STUDENTS. Select Evaluations of students "
                        "at the top of OER and upload it there. Student assessments "
                        "cannot be added to the educator-feedback archive. Nothing was saved."
                    )
                if not student_archive:
                    continue
                record_id = values[record_index].strip()
                if record_id:
                    identity = (values[course_index].strip(), key, record_id)
                    forms.add(identity)
                    type_forms.setdefault(title, set()).add(identity)
                else:
                    missing_ids += 1
                type_rows[title] += 1
                if submitted_index is not None and not values[submitted_index].strip():
                    blank_submissions += 1
        except (csv.Error, UnicodeError, StopIteration):
            raise OASISArchiveError("The CSV could not be inspected safely. Re-export the original OASIS CSV.") from None
        finally:
            csv.field_size_limit(old_limit)
    if student_archive:
        details.update({
            "archive_kind": "student_evaluations",
            "form_count": len(forms),
            "form_types": [title for title in STUDENT_FORM_TITLES if title in type_rows],
            "form_type_counts": [
                {"evaluation_form": title, "question_rows": type_rows[title],
                 "forms_in_file": len(type_forms.get(title, set()))}
                for title in STUDENT_FORM_TITLES if title in type_rows
            ],
            "missing_form_record_rows": missing_ids,
            "blank_submit_date_rows": blank_submissions,
        })
    return details


def inspect_student_oasis_csv(raw: bytes) -> dict[str, Any]:
    """Return aggregate metadata for the supplied student-assessment format."""
    return _form_metadata(raw, student_archive=True)


def validate_educator_upload_kind(raw: bytes) -> None:
    """Prevent the known student forms entering the existing educator pipeline."""
    _form_metadata(raw, student_archive=False)


def student_export_label(filename: str) -> str:
    """Course dates and opaque export ID only; no student or preceptor names."""
    return oasis_export_label(filename)


class GitHubOASISStudentEvaluations(GitHubOASISEvaluations):
    """Reuse verified ciphertext storage in a different, non-recursive folder.

    The original educator service always reads oasis_evaluations/; it never scans
    this sibling folder. A separate HMAC domain also prevents identifier reuse
    between the two archives. Keep current/previous encryption keys unchanged.
    """

    def __init__(self, archive):
        super().__init__(archive)
        self.folder = f"{self.config.folder}/{STUDENT_SUBFOLDER}"

    def _candidate_names(self, raw: bytes, details: dict[str, Any]) -> list[str]:
        names = []
        for encoded_key in (self.config.encryption_key, *self.config.previous_encryption_keys):
            key = base64.urlsafe_b64decode(encoded_key.encode("ascii"))
            id_key = hmac.new(key, _STUDENT_ID_CONTEXT, hashlib.sha256).digest()
            identity = hmac.new(id_key, raw, hashlib.sha256).hexdigest()
            filename = f"OASIS_{details['coverage']}_{identity}.csv.enc"
            if filename not in names:
                names.append(filename)
        return names

    def _inspect(self, raw: bytes) -> dict[str, Any]:
        return inspect_student_oasis_csv(raw)

    def _minimize(self, raw: bytes) -> dict[str, Any]:
        from schedule_app.services.oasis_privacy import minimize_oasis_csv
        inspect_student_oasis_csv(raw)  # Reject foreign/mixed forms before saving.
        return minimize_oasis_csv(raw, "student")
