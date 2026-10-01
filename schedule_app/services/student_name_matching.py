"""Exact student-name matching after removing explicit program designations.

This module does not fuzzy-match names or change encrypted catalog keys.  The
existing student_name_key remains the storage key so older confirmed matches
(including names ending in '(MD)') remain readable, editable and removable.
"""
from __future__ import annotations
from collections import defaultdict
from dataclasses import dataclass
import re
from typing import Mapping, Sequence

from schedule_app.services.student_assessment_links import student_name_key

# Only recognized, trailing labels are ignored. Never remove arbitrary text in
# parentheses, family-name suffixes, a middle name, or other identity details.
_PROGRAM = r"(?:M\.?\s*D\.?|P\.?\s*A\.?|D\.?\s*O\.?)"
_PROGRAM_WITH_YEAR = rf"{_PROGRAM}(?:\s*\d{{4}})?"
_TRAILING_DESIGNATION = re.compile(
    rf"(?:\s*\(\s*{_PROGRAM_WITH_YEAR}\s*\)|\s*;\s*{_PROGRAM_WITH_YEAR})\s*$",
    re.IGNORECASE,
)
_MISSING_IDS = {"", "nan", "n/a", "none", "null", "-", "--"}


def student_name_without_designations(value: object) -> str:
    """Remove explicit trailing (MD)/(PA)/(DO) and matching class-year labels.

    Supports '; MD2028', '(MD 2028)' and stacked labels. Case/spacing and comma
    spacing are normalized by student_matching_key; the original sources and
    saved display names are never rewritten by this function.
    """
    text = re.sub(r"\s+", " ", str(value or "").strip())
    while True:
        cleaned = _TRAILING_DESIGNATION.sub("", text).strip()
        if cleaned == text:
            return text
        text = cleaned


def student_matching_key(value: object) -> str:
    """Exact name key for matching ONLY; not a new persistent catalog key."""
    return student_name_key(student_name_without_designations(value))


@dataclass(frozen=True)
class StudentNameMatch:
    # Keep the existing key for the correction form and encrypted saved entry.
    name_key: str
    candidate_ids: tuple[str, ...]
    match_status: str


class StudentNameMatcher:
    """Use the same identity decision for calculation and missing-name alerts.

    A previously confirmed exact catalog entry has priority. Otherwise collect
    ALL IDs with the same designation-free name; resolve only when there is one.
    Two people whose names collapse to the same key remain ambiguous. This does
    not infer student usernames, strip real-name differences or write to GitHub.
    """

    def __init__(self, name_ids: Mapping[str, Sequence[str]],
                 saved_entries: Mapping[str, Mapping[str, str]]):
        by_name = defaultdict(set)
        for name, ids in name_ids.items():
            key = student_matching_key(name)
            if not key:
                continue
            for sid in ids:
                sid = str(sid).strip()
                if sid.casefold() not in _MISSING_IDS:
                    by_name[key].add(sid)
        self._name_ids = {key: tuple(sorted(ids)) for key, ids in by_name.items()}
        self._saved_entries = saved_entries

    def resolve(self, opd_name: object) -> StudentNameMatch:
        key = student_name_key(opd_name)
        saved_id = str(self._saved_entries.get(key, {}).get("external_id", "")).strip()
        if saved_id.casefold() not in _MISSING_IDS:
            return StudentNameMatch(key, (saved_id,), "Saved match")
        matching_key = student_matching_key(opd_name)
        candidates = self._name_ids.get(matching_key, ()) if matching_key else ()
        if len(candidates) != 1:
            status = "Needs review"
        elif key != matching_key:
            status = "Exact name match (designation ignored)"
        else:
            status = "Exact normalized name match"
        return StudentNameMatch(key, candidates, status)
