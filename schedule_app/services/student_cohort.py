"""One in-memory student cohort for continuity and assessment eligibility.

Identities come from the existing confirmed-name/OASIS resolver. Only aggregate
counts leave this module for reports; student keys and names are never exported
or saved by this calculation. The caller supplies OPD assignments AFTER the
existing ignore-list and outpatient-over-nursery rules.
"""
from collections import defaultdict
from datetime import date

from schedule_app.services.opd_archive import OPDArchiveError

STUDENT_COHORT_VERSION = 1


def group_student_assignments(assignments, resolver, active_names, start_date, end_date):
    """Group unique date/AM-PM slots using a single reconciled student identity."""
    active = set(active_names)
    pairs, unresolved = defaultdict(set), {}
    for item in assignments:
        if item['preceptor_name'] not in active or not start_date <= item['date'] <= end_date:
            continue
        # Refuse malformed source dates/periods rather than creating another slot.
        try:
            day = date.fromisoformat(item['date']).isoformat()
            if day != item['date'] or item['shift'] not in ('AM', 'PM'):
                raise ValueError('invalid slot')
        except (ValueError, TypeError):
            raise OPDArchiveError('A retained student assignment has an invalid date or AM/PM shift. Refresh the OPDs.') from None
        match = resolver.resolve(item['student'])
        pair = (item['preceptor_name'], match.identity_key)
        pairs[pair].add((day, item['shift']))
        if match.status != 'Matched':
            unresolved.setdefault(pair, {})[match.name_key] = (item['student'], match)
    return pairs, unresolved


def student_cohort_counts(pairs, preceptor_name, minimum_shifts):
    """Derive all three counts from the SAME student-to-slot sets."""
    groups = {identity: slots for (name, identity), slots in pairs.items()
              if name == preceptor_name and slots}
    eligible = {identity for identity, slots in groups.items() if len(slots) >= minimum_shifts}
    days3 = sum(len({day for day, _ in slots}) >= 3 for slots in groups.values())
    counts = {'unique_students': len(groups), 'unique_students_3plus_days': days3,
              'eligible_students': len(eligible)}
    validate_student_cohort_counts(counts, minimum_shifts)
    return counts, eligible


def validate_student_cohort_counts(row, minimum_shifts):
    """Catch mismatched dates/identities BEFORE publishing conflicting counts.

    Three distinct days require at least three AM/PM slots, so with a minimum
    of 1, 2 or 3 shifts the eligible count cannot be below the three-day count.
    This inequality is deliberately NOT imposed when the user selects 4+ shifts.
    """
    values = [row.get(key) for key in ('unique_students', 'unique_students_3plus_days', 'eligible_students')]
    if any(type(value) is not int or value < 0 for value in values):
        raise OPDArchiveError('Student continuity/eligibility counts are missing or invalid. Refresh evaluation completeness.')
    unique, days3, eligible = values
    if (days3 > unique or eligible > unique
            or (minimum_shifts is not None and minimum_shifts <= 3 and eligible < days3)):
        raise OPDArchiveError(
            'Student continuity and assessment eligibility do not reconcile for '
            f"{row.get('preceptor_name', 'a preceptor')}. Recalculate evaluation completeness: "
            'the same student identities and cutoff must be used in both sections.')
