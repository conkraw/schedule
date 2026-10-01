"""Documented completion policy: absence is not an identity-confirmation task.

Similar names are review suggestions ONLY, never inferred identity links. A
student missing from OASIS remains an OPD-derived denominator member. Existing
explicit, encrypted links remain authoritative. No new identifiers are saved.
"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
import re

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.student_name_matching import StudentNameMatcher, student_matching_key
from schedule_app.services.teaching_analysis import teaching_local_today

MATCHED = "Matched"
ABSENT = "No assessment on file"
REVIEW = "Name review needed"


def assessment_as_of(value=None):
    """Use a local calendar date. An explicit cutoff cannot be in the future."""
    if value is None:
        value = teaching_local_today()
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, str):
        try:
            value = date.fromisoformat(value)
        except ValueError:
            raise OPDArchiveError("Choose a valid Assessments as of date.") from None
    if not isinstance(value, date) or value > teaching_local_today():
        raise OPDArchiveError("Assessments as of must be a valid date on or before today.")
    return value


def completion_window(start, end, as_of):
    """An empty intersection is valid; it contains no eligible shifts or forms."""
    return start.isoformat(), min(end, assessment_as_of(as_of)).isoformat()


def plausible_name_difference(left, right):
    """Conservative review cue, not fuzzy auto-matching or proof of identity.

    Catch token reordering, an initial/full-name difference with the rest exact,
    or a small typo in one token while every other token matches. Shared surnames
    alone must not send clearly different students into the confirmation queue.
    """
    a = sorted(re.findall(r"[^\W_]+", student_matching_key(left), flags=re.UNICODE))
    b = sorted(re.findall(r"[^\W_]+", student_matching_key(right), flags=re.UNICODE))
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) != len(b) or len(a) < 2:
        return False
    rest_a, rest_b = list(a), list(b)
    for token in a:
        if token in rest_b:
            rest_a.remove(token)
            rest_b.remove(token)
    if len(rest_a) != 1 or len(rest_b) != 1:
        return False
    x, y = rest_a[0], rest_b[0]
    if len(x) == 1 or len(y) == 1:
        return x[0] == y[0]
    # Limit the change to a short typo. A loose ratio alone matches many distinct
    # given names (Eva/Ava, for example), so require at least four letters here.
    return (min(len(x), len(y)) >= 4 and abs(len(x) - len(y)) <= 2
            and SequenceMatcher(None, x, y).ratio() >= 0.75)


@dataclass(frozen=True)
class AssessmentStudent:
    name_key: str             # Existing encrypted-catalog key, not modified.
    identity_key: tuple      # Run-local grouping; never written to GitHub.
    candidate_ids: tuple
    status: str
    suggestions: tuple = ()  # UI-only names; no ID is inferred from these.


class AssessmentIdentityResolver:
    def __init__(self, prepared, saved_entries):
        self.matcher = StudentNameMatcher(prepared.get("name_ids", {}), saved_entries)
        self.names = prepared.get("student_names", {})
        self.known = {student_matching_key(k): v for k, v in self.names.items()}
        for key in prepared.get("name_ids", {}):
            self.known.setdefault(student_matching_key(key), self.names.get(key, key))
        self.cache = {}

    def resolve(self, opd_name):
        # Cache per exact input: a saved designation-specific override has priority.
        text = str(opd_name)
        if text in self.cache:
            return self.cache[text]
        match = self.matcher.resolve(opd_name)
        ids = match.candidate_ids
        normalized = student_matching_key(opd_name)
        if len(ids) == 1:
            item = AssessmentStudent(match.name_key, ("external_id", ids[0]), ids, MATCHED)
        else:
            suggestions = tuple(sorted({name for key, name in self.known.items()
                                        if key == normalized or plausible_name_difference(normalized, key)},
                                       key=str.casefold))
            # A known name without a usable ID or with multiple IDs is a genuine
            # source-identity issue. Complete absence needs no confirmation.
            status = REVIEW if len(ids) > 1 or normalized in self.known or suggestions else ABSENT
            item = AssessmentStudent(match.name_key, ("opd_name", normalized), ids, status, suggestions)
        self.cache[text] = item
        return item


def has_documented_percentage(row):
    status = str(row.get("assessment_status", ""))
    return status == "Calculated" or status.startswith("Provisional:")
